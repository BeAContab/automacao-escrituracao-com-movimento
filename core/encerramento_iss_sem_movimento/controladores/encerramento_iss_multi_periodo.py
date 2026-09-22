"""Encerramento ISS (Sem Movimento) — várias competências numa só execução.

CÓPIA INDEPENDENTE de `encerramento_iss.py` (a função "Encerramento ISS" de uma competência
só, que já funciona e NÃO deve ser alterada). Em vez de uma única `competencia: date`, esta
versão recebe uma lista de competências e, para cada empresa da planilha, busca a inscrição
**uma vez** e processa todas as competências antes de passar para a próxima empresa — evita
repetir a busca da inscrição a cada mês.

Reaproveita (sem editar) as peças genéricas do módulo original: `EmpresaSemMovimentoISSFortaleza`,
`CredenciaisISSFortaleza`, `ChromeDriver`, `retry_on_exception`, utilitários de CNPJ/data, e os
tipos `StatusEncerramentoISS`/`CodigoEncerramentoISS` (importados, nunca modificados). O restante
da lógica de navegação/portal foi copiado e adaptado, porque o original amarra fortemente
"1 problema = pula para a próxima empresa", o que aqui precisa ser "1 problema = pula só para a
próxima competência da mesma empresa".

Diferenças deliberadas em relação ao original:
- pasta de saída organizada por `<ano>-<mês>` (não só `<mês>`), para não colidir competências do
  mesmo mês em anos diferentes (ex.: Jan/2025 e Jan/2026);
- um relatório de execução por competência, não um relatório único;
- a planilha de entrada NÃO tem uma coluna por competência: a data/status de cada competência é
  ACRESCENTADA na mesma coluna (Y) já usada pelo original, separada por "; " — decisão do usuário,
  ciente de que a célula fica com várias informações concatenadas;
  ao final da célula, sem apagar o que já estava lá;
- um erro inesperado numa competência específica não derruba a execução inteira: fica registrado
  como problema daquela competência e o robô segue para a próxima (competência ou empresa).

ATENÇÃO — duplicação consciente: uma correção futura no fluxo do portal (seletores, navegação)
pode precisar ser feita nos dois módulos (este e `encerramento_iss.py`).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from math import floor
from pathlib import Path
from pprint import pformat
from typing import Any, Callable

from openpyxl.cell import MergedCell
from pydantic import ValidationError
from loguru import logger
from selenium.common import NoSuchElementException, TimeoutException
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
)

from python_spreadsheet_reader.readers import SpreadsheetIsLockedException, XLSXReader
from python_spreadsheet_reader.styler import HorizontalAlignment, VerticalAlignment

from core.encerramento_iss_sem_movimento.controladores.encerramento_iss import (
    AUTORES_RELATORIO,
    VERSAO_RELATORIO,
    CodigoEncerramentoISS,
    StatusEncerramentoISS,
)
from core.encerramento_iss_sem_movimento.iss.credenciais import CredenciaisISSFortaleza
from core.encerramento_iss_sem_movimento.iss.empresa import EmpresaSemMovimentoISSFortaleza, MensagemISS
from core.encerramento_iss_sem_movimento.python_webdriver.driver.chrome import ChromeDriver
from core.encerramento_iss_sem_movimento.python_webdriver.functions import retry_on_exception
from core.encerramento_iss_sem_movimento.utils.classes import UserStoppedThreadException
from core.encerramento_iss_sem_movimento.utils.constants import FORMATO_DATA
from core.encerramento_iss_sem_movimento.utils.functions import (
    aplicar_mascara_cnpj,
    remover_pontuacao_cnpj,
    safe_division,
    timestamp_as_file_name,
)


# ---------------------------------------------------------------------------
# Resultados (um registro por empresa POR competência, ao contrário do original
# que guarda o resultado direto no objeto Empresa — aqui um mesmo objeto Empresa
# passa por várias competências, então cada uma precisa do seu próprio registro)
# ---------------------------------------------------------------------------


@dataclass
class ResultadoMesEmpresa:
    empresa: EmpresaSemMovimentoISSFortaleza
    competencia: date
    encerrada: bool
    dt_encerramento: date | None = None
    problemas: list[str] = field(default_factory=list)
    caminho_certificado: Path | None = None


class ResultadoEncerramentoISSMultiPeriodo:
    def __init__(
        self,
        status: StatusEncerramentoISS,
        resultados: list[ResultadoMesEmpresa],
        report_paths: list[Path],
    ) -> None:
        self.status = status
        self.resultados = resultados
        self.report_paths = report_paths

    @property
    def processadas(self) -> int:
        return len(self.resultados)

    @property
    def encerradas(self) -> int:
        return sum(1 for r in self.resultados if r.encerrada)

    @property
    def problemas(self) -> int:
        return sum(1 for r in self.resultados if not r.encerrada)


class SemEmpresasProcessadasException(Exception):
    """Nenhuma empresa foi processada para nenhuma competência (nada a gerar de relatório)."""


# ---------------------------------------------------------------------------
# Controlador
# ---------------------------------------------------------------------------


class ControladorEncerramentoISSMultiPeriodo:
    def __init__(
        self,
        caminho_planilha: Path,
        caminho_saida: Path,
        caminho_webdriver: Path,
        url_iss_fortaleza: str,
        credenciais: CredenciaisISSFortaleza,
        competencias: list[date],
        caminho_template_relatorio: Path,
        check_thread_stopped_callback: Callable[[], None],
    ) -> None:
        if not competencias:
            raise ValueError("Informe ao menos uma competência.")
        self.caminho_planilha = caminho_planilha
        self.caminho_saida = caminho_saida
        self.caminho_webdriver = caminho_webdriver
        self.url_iss_fortaleza = url_iss_fortaleza
        self.credenciais = credenciais
        # Ordem cronológica e sem duplicatas, independente da ordem em que o operador as adicionou.
        self.competencias: list[date] = sorted(set(competencias))
        self.caminho_template_relatorio = caminho_template_relatorio
        self.check_thread_stopped_callback = check_thread_stopped_callback
        self.resultados: list[ResultadoMesEmpresa] = []

    # -- fluxo principal -------------------------------------------------

    def executar_processo(self) -> ResultadoEncerramentoISSMultiPeriodo:
        rotulos = ", ".join(c.strftime("%m/%Y") for c in self.competencias)
        logger.info(
            f"Vou percorrer a planilha em busca das empresas sem movimento para encerrar o ISS delas "
            f"em {len(self.competencias)} competência(s): {rotulos}."
        )

        driver = ChromeDriver(driver_path=self.caminho_webdriver)
        driver.start_driver(options=("--start-maximized",))

        reader = XLSXReader(self.caminho_planilha)

        status = StatusEncerramentoISS()

        try:
            self.check_thread_stopped_callback()
            self._login_iss_fortaleza(driver)

            for row_number, row in reader.lazy_load_sheet():
                self.check_thread_stopped_callback()
                try:
                    self._processar_linha(driver, reader, row_number, row)
                except UserStoppedThreadException:
                    raise
                except Exception as exc:  # noqa: BLE001 — um erro inesperado numa empresa não pode abortar a planilha inteira
                    logger.error(f"Erro inesperado na linha {row_number} da planilha — vou pular para a próxima. Detalhe: {exc}")

            msg = "Terminei de percorrer a planilha para todas as competências solicitadas."
            logger.success(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.SUCCESS, message=msg)
        except UserStoppedThreadException:
            msg = "Recebi o pedido de cancelamento do operador e parei o processo aqui."
            logger.warning(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.USER_ENDED_PROCESS, message=msg)
        except FileNotFoundError as fnf_err:
            msg = f"Não foi possível achar o arquivo: {fnf_err.filename}"
            logger.error(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.ERROR_FILE_NOT_FOUND, message=msg)
        except SpreadsheetIsLockedException as locked:
            logger.error(locked)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.ERROR_SPREADSHEET_LOCKED, message=str(locked))
        except Exception as exc:  # noqa: BLE001
            logger.error(exc)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION, message=str(exc))
        finally:
            reader.close_workbook()
            driver.quit_driver()

        report_paths: list[Path] = []
        try:
            report_paths = self._gerar_relatorios_de_execucao()
        except SemEmpresasProcessadasException as sem_empresa:
            logger.warning(sem_empresa)

        return ResultadoEncerramentoISSMultiPeriodo(status, self.resultados, report_paths)

    def _processar_linha(self, driver: ChromeDriver, reader: XLSXReader, row_number: int, row: dict) -> None:
        try:
            empresa = EmpresaSemMovimentoISSFortaleza(
                codigo=int(row[f"A{row_number}"].value),
                nome=row[f"B{row_number}"].value,
                cnpj=remover_pontuacao_cnpj(row[f"D{row_number}"].value),
                responsavel=row[f"H{row_number}"].value,
                municipio=row[f"W{row_number}"].value,
                linha_planilha_fiscal=row_number,
            )
        except KeyError as k_err:
            logger.error(
                f"Essa linha da planilha não tem a coluna `{k_err}` que eu esperava, então vou pular para a próxima. "
                f"Detalhe: {pformat(k_err)}"
            )
            return
        except ValidationError as v_err:
            logger.debug(f"Os dados dessa empresa não passaram na validação — vou pular essa linha. Detalhe:\n{v_err.json()}")
            return
        except (ValueError, TypeError):
            return  # Campo None ou de outro tipo inesperado

        self.check_thread_stopped_callback()
        empresa.cnpj_m = aplicar_mascara_cnpj(empresa.cnpj)

        if not self._procurar_inscricao_empresa(driver, empresa):
            logger.warning(
                f"Não encontrei procuração para a empresa {empresa} no portal — vou marcar como 'sem inscrição' "
                "para todas as competências e seguir para a próxima."
            )
            msg = "S/ INSCRIÇÃO"
            for competencia in self.competencias:
                self.resultados.append(ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg]))
            self._acrescentar_na_planilha_fiscal(reader, competencia=None, valor=msg, row_number=empresa.linha_planilha_fiscal)
            return  # Sem inscrição: não há como trocar de empresa pelo modal (ele não abriu)

        self._dar_ciencia_nas_mensagens_nao_lidas(driver, empresa)
        self._navegar_tela_escrituracao(driver)

        for competencia in self.competencias:
            self.check_thread_stopped_callback()
            try:
                resultado = self._processar_competencia(driver, reader, empresa, competencia)
            except UserStoppedThreadException:
                raise
            except (NoSuchElementException, TimeoutException) as exc:
                msg = f"Não consegui processar essa competência no portal: {type(exc).__name__}"
                logger.error(f"{empresa} — {competencia.strftime('%m/%Y')}: {msg}")
                resultado = ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg])
                self._acrescentar_na_planilha_fiscal(reader, competencia, msg, empresa.linha_planilha_fiscal)
            self.resultados.append(resultado)

        self._abrir_modal_de_alteracao_de_inscricao(driver)

    # -- login / navegação (cópia do original; não dependem da competência) ----

    def _login_iss_fortaleza(self, driver: ChromeDriver) -> None:
        logger.info("Estou entrando no portal da ISS Fortaleza com o CPF e a senha informados...")

        driver.goto(self.url_iss_fortaleza)
        driver.click(driver.find_element().by_attribute(attr_name="href", attr_value="/grpfor/oauth2/login"))

        driver.type(element=driver.find_element().by_id("username"), text=self.credenciais.cpf.get_secret_value())
        driver.type(element=driver.find_element().by_id("password"), text=self.credenciais.senha.get_secret_value())

        self.check_thread_stopped_callback()
        driver.find_element().by_id("kc-form-login").get_element().submit()

    def _procurar_inscricao_empresa(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza) -> bool:
        """Procura a inscrição da empresa no portal. Devolve False (sem levantar) se não achar."""
        logger.info(f"Agora é a vez da empresa {empresa}: vou procurar a inscrição dela no portal.")

        self.check_thread_stopped_callback()

        driver.click(driver.find_element().by_xpath('//table[@id="alteraInscricaoForm:tipoPesquisa"]//input[@value="CNPJ"]'))
        driver.explicit_wait(1)  # AJAX reaplica a máscara — digitar rápido demais apaga o texto

        driver.clear(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"))
        driver.type(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"), empresa.cnpj)
        driver.explicit_wait(1)

        record_xpath = f'//tbody[@id="alteraInscricaoForm:empresaDataTable:tb"]//a[normalize-space()="{empresa.cnpj_m}"]'
        bln_empresa_listada = bool(driver.find_element().by_xpath(record_xpath).get_all_elements())

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("alteraInscricaoForm:btnPesquisar"))

        try:
            if bln_empresa_listada:
                driver.explicit_wait(5)
            driver.click(driver.find_element().by_xpath(record_xpath), max_retries=5)
        except (TimeoutException, NoSuchElementException):
            return False

        self._confirmar_alteracao_de_inscricao(driver)
        logger.info(f"Encontrei a procuração da empresa {empresa}. Entrando na página dela agora...")
        return True

    def _dar_ciencia_nas_mensagens_nao_lidas(self, driver: ChromeDriver, emp: EmpresaSemMovimentoISSFortaleza) -> None:
        try:

            def _wait_for_messages_modal_to_popup():
                return driver.find_element().by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr").get_element()

            retry_on_exception(
                func=_wait_for_messages_modal_to_popup, max_attempts=15, polling_seconds=1, exception=NoSuchElementException
            )

            logger.info("Essa empresa tem mensagens não lidas no portal — vou abrir cada uma e dar ciência antes de continuar.")

            driver.drag_and_drop_by_offset(driver.find_element().by_id("mensagensModalHeader"), x_offset=0, y_offset=-200)

            messages = driver.find_element().by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr").get_all_elements()
            for _msg in messages:
                self.check_thread_stopped_callback()

                driver.click(driver.find_element().by_xpath('//a[contains(@id, "mensagensForm:mensagemDataTable:0:linkTitulo")]'))

                def _collect_message_data() -> MensagemISS:
                    title = driver.find_element().by_id("mensagensForm:titulo").get_element()
                    data = driver.find_element().by_id("mensagensForm:dataRegistro").get_element()
                    conteudo = driver.find_element().by_id("mensagensForm:descricao").get_element()
                    return MensagemISS(title=title.text, date=datetime.strptime(data.text, "%d/%m/%Y").date(), content=conteudo.text)

                m = retry_on_exception(func=_collect_message_data, max_attempts=15, polling_seconds=1)
                emp.mensagens.append(m)

                try:
                    anexos = driver.find_element().by_xpath(
                        '//*[@id="mensagensForm:divAnexos"]//a[contains(@id, "linkVisualizarAnexo")]'
                    ).get_all_elements()
                    logger.info("Essa mensagem tem anexos — vou baixar todos antes de dar ciência.")
                    for idx, _anexo in enumerate(anexos):
                        self.check_thread_stopped_callback()
                        try:
                            seletor_anexo = f"mensagensForm:anexos:{idx}:linkVisualizarAnexo"
                            # 1º caso: o anexo abre num modal de leitura antes de baixar.
                            driver.click(driver.find_element().by_id(seletor_anexo), max_retries=5)
                            driver.click(driver.find_element().by_id("mensagensForm:botaoBaixarAnexo"))
                            driver.explicit_wait(0.5)
                            driver.click(driver.find_element().by_id("mensagensForm:botaoVoltarAnexo"))
                        except (ElementClickInterceptedException, NoSuchElementException, StaleElementReferenceException):
                            # 2º caso: o link já baixa o anexo direto, sem modal.
                            seletor_anexo = f"mensagensForm:anexos:{idx}:linkBaixarAnexo"
                            driver.click(driver.find_element().by_id(seletor_anexo), max_retries=5)

                        driver.explicit_wait(1)
                        driver.hide_elements((f"#{seletor_anexo}",))
                except (TimeoutException, StaleElementReferenceException, NoSuchElementException):
                    pass  # Sem anexo

                self.check_thread_stopped_callback()
                driver.click(driver.find_element().by_id("mensagensForm:botaoDarCiencia"))
        except NoSuchElementException:
            pass

    def _navegar_tela_escrituracao(self, driver: ChromeDriver) -> None:
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_xpath('//ul[contains(@class, "navbar-nav")]//a[normalize-space()="Escrituração"]'))
        driver.click(
            driver.find_element().by_xpath(
                '//ul[contains(@class, "navbar-nav")]//li[@class="dropdown open"]//ul[@class="dropdown-menu"]//a[normalize-space()="Manter Escrituração"]'
            )
        )

    @staticmethod
    def _confirmar_alteracao_de_inscricao(driver: ChromeDriver) -> None:
        try:
            driver.click(driver.find_element().by_id("alteraInscricaoForm:botaoOk"), max_retries=5)
        except (ElementNotInteractableException, NoSuchElementException, StaleElementReferenceException):
            pass  # Modal não apareceu

    def _abrir_modal_de_alteracao_de_inscricao(self, driver: ChromeDriver) -> None:
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_attribute(attr_name="title", attr_value="Alterar Inscrição Atual"))

    # -- por competência (adaptado do original: recebe a competência como parâmetro
    #    e devolve o resultado em vez de gravar direto em `self` e abrir o modal) ----

    def _preencher_widget_calendario(self, driver: ChromeDriver, competencia: date) -> None:
        logger.debug(f"Abrindo o calendário do portal para selecionar a competência {competencia.strftime('%m/%Y')}...")

        self.check_thread_stopped_callback()

        driver.click(driver.find_element().by_class("rich-calendar-tool-btn"))
        driver.click(driver.find_element().by_id(f"manterEscrituracaoForm:dataInicialDateEditorLayoutM{competencia.month - 1}"))
        driver.click(
            driver.find_element().by_xpath(
                f'//div[contains(@id, "manterEscrituracaoForm:dataInicialDateEditorLayoutY") and normalize-space()="{competencia.year}"]'
            )
        )

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("manterEscrituracaoForm:dataInicialDateEditorButtonOk"))

    def _verificar_servicos_prestados(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza) -> str | None:
        """Devolve o motivo do problema (sem gravar nada), ou None se não houver serviços prestados."""
        logger.info("Antes de encerrar, vou conferir se essa empresa tem serviços prestados no período...")

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("abaServicosPrestados_lbl"))

        def _aguardar_elemento_carregar():
            return driver.find_element().by_xpath(
                "//table[@id='abaEncerramentoForm:dataTableServicosPrestados']"
                "//td[normalize-space(text())='Somatório']/following-sibling::td[1]"
            ).get_element()

        quantidade_cell = retry_on_exception(func=_aguardar_elemento_carregar, max_attempts=60, polling_seconds=1)

        motivo: str | None = None
        try:
            quantidade = int(quantidade_cell.text.strip())
            if quantidade > 0:
                logger.warning(f"A empresa {empresa} tem serviços prestados no período — vou marcar como problema.")
                motivo = "SERVIÇOS PRESTADOS"
        except (TypeError, ValueError):
            logger.error(
                f"Não consegui ler direito a quantidade de serviços prestados da empresa {empresa} — "
                f"o portal retornou um valor que eu não esperava: {quantidade_cell.text!r}"
            )
            motivo = "ERRO AO LER SERVIÇOS PRESTADOS"

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("abaEncerramento_lbl"))
        return motivo

    def _verificar_servicos_pendentes(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza) -> str | None:
        """Devolve o motivo do problema (sem gravar nada), ou None se não houver serviços pendentes."""
        logger.info("Também vou conferir se há serviços pendentes registrados para essa empresa...")

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("aba_servicos_pendentes_lbl"))

        def _aguardar_elemento_carregar():
            return driver.find_element().by_id("servicos_pendentes_form:table_servico_tomados_pendente:tb").get_element()

        servicos_pendentes_table = retry_on_exception(func=_aguardar_elemento_carregar, max_attempts=60, polling_seconds=1)
        row_count = servicos_pendentes_table.get_property("childElementCount")

        motivo: str | None = None
        if not isinstance(row_count, int):
            logger.error(f"Não consegui confirmar se a empresa {empresa} tem serviços pendentes — tipo de dado inesperado.")
            motivo = "ERRO AO LER SERVIÇOS PENDENTES"
        elif row_count > 0:
            logger.warning(f"A empresa {empresa} tem serviços pendentes — vou marcar como problema.")
            motivo = "SERVIÇOS PENDENTES"

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("abaEncerramento_lbl"))
        return motivo

    def _processar_competencia(
        self, driver: ChromeDriver, reader: XLSXReader, empresa: EmpresaSemMovimentoISSFortaleza, competencia: date
    ) -> ResultadoMesEmpresa:
        self._preencher_widget_calendario(driver, competencia)

        logger.debug("Consultando no portal se há escriturações em aberto para essa competência...")
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("manterEscrituracaoForm:btnConsultar"))

        competencia_str = competencia.strftime("%m/%Y")

        def _locate_element():
            element = driver.find_element().by_xpath(
                f'//tbody[@id="manterEscrituracaoForm:dataTable:tb"]/tr/td/span[normalize-space()="{competencia_str}"]/../../self::tr'
            )
            element.get_element()
            return element

        search_result_row = retry_on_exception(func=_locate_element, max_attempts=60, polling_seconds=1, exception=NoSuchElementException)
        link_escriturar = driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "linkEscriturar")]')
        link_escriturar_title: str | None = link_escriturar.get_element().get_dom_attribute("title")

        if not link_escriturar_title:
            msg = "O portal não me deu informação suficiente para saber se a escrituração pode ser encerrada."
            logger.error(f"{empresa} — {competencia_str}: {msg}")
            self._acrescentar_na_planilha_fiscal(reader, competencia, msg, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg])

        match link_escriturar_title.upper().strip():
            case "ESCRITURAÇÃO ENCERRADA":
                return self._processar_ja_encerrada(driver, reader, empresa, competencia, search_result_row)
            case "ESCRITURAR":
                return self._processar_a_encerrar(driver, reader, empresa, competencia, link_escriturar)
            case _:
                msg = f"Título inesperado do botão de escriturar: {link_escriturar_title!r}"
                logger.error(f"{empresa} — {competencia_str}: {msg}")
                self._acrescentar_na_planilha_fiscal(reader, competencia, msg, empresa.linha_planilha_fiscal)
                return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg])

    def _processar_ja_encerrada(
        self,
        driver: ChromeDriver,
        reader: XLSXReader,
        empresa: EmpresaSemMovimentoISSFortaleza,
        competencia: date,
        search_result_row: Any,
    ) -> ResultadoMesEmpresa:
        competencia_str = competencia.strftime("%m/%Y")
        logger.warning(f"A escrituração da empresa {empresa} já tinha sido encerrada antes ({competencia_str}) — vou só baixar o certificado.")

        self.check_thread_stopped_callback()
        driver.click(driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "certificado")]'))
        driver.click(driver.find_element().by_id("formMenuTopo"))  # aguarda o certificado ficar visível

        caminho_certificado = self._imprimir_declaracao_fechamento_iss(driver, empresa, competencia)

        driver.get_driver().back()  # o botão de mudar empresa fica escondido enquanto o certificado é exibido

        dt_texto: str = driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "dataEncerramento")]').get_element().text
        dt_encerramento = datetime.strptime(dt_texto, FORMATO_DATA).date()

        self._acrescentar_na_planilha_fiscal(reader, competencia, dt_texto, empresa.linha_planilha_fiscal)

        return ResultadoMesEmpresa(empresa, competencia, encerrada=True, dt_encerramento=dt_encerramento, caminho_certificado=Path(caminho_certificado))

    def _processar_a_encerrar(
        self,
        driver: ChromeDriver,
        reader: XLSXReader,
        empresa: EmpresaSemMovimentoISSFortaleza,
        competencia: date,
        link_escriturar: Any,
    ) -> ResultadoMesEmpresa:
        competencia_str = competencia.strftime("%m/%Y")
        logger.info(f"A empresa {empresa} ainda está com a escrituração em aberto ({competencia_str}) — vou verificar as pendências e encerrar.")

        self.check_thread_stopped_callback()
        driver.click(link_escriturar)

        motivo = self._verificar_servicos_prestados(driver, empresa)
        if motivo is None:
            motivo = self._verificar_servicos_pendentes(driver, empresa)

        if motivo is not None:
            self._acrescentar_na_planilha_fiscal(reader, competencia, motivo, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[motivo])

        btn_encerrar = driver.find_element().by_id("abaEncerramentoForm:btnEncerrarEscrituracao")
        driver.scroll_element_into_view(btn_encerrar)
        driver.click(btn_encerrar)

        logger.debug("Confirmando no portal o encerramento da escrituração...")
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("formEncerramento:btnSim"))

        logger.info("Encerramento confirmado! Agora vou baixar o certificado dessa empresa.")
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_xpath('//input[contains(@id, "btnCertificadoEscrituracao")]'))
        driver.find_element().by_id("docPrincipal")  # aguarda o documento carregar, senão o PDF vem vazio

        caminho_certificado = self._imprimir_declaracao_fechamento_iss(driver, empresa, competencia)

        dt_encerramento = date.today()
        self._acrescentar_na_planilha_fiscal(reader, competencia, dt_encerramento.strftime(FORMATO_DATA), empresa.linha_planilha_fiscal)

        logger.success(f"Pronto! Encerrei a escrituração da empresa {empresa} para a competência {competencia_str}.")
        driver.get_driver().back()  # o botão de mudar empresa fica escondido enquanto o certificado é exibido

        return ResultadoMesEmpresa(empresa, competencia, encerrada=True, dt_encerramento=dt_encerramento, caminho_certificado=Path(caminho_certificado))

    def _imprimir_declaracao_fechamento_iss(self, driver: ChromeDriver, emp: EmpresaSemMovimentoISSFortaleza, competencia: date) -> str:
        self.check_thread_stopped_callback()

        empresa = f"EMP_{emp.codigo}"
        arquivo = f"CERTIFICADO ISS_EMP{emp.codigo}_{timestamp_as_file_name('pdf')}"
        # <ano>-<mês>, diferente do original (só <mês>): evita colidir competências do mesmo mês em anos diferentes.
        diretorio = f"{self.caminho_saida}/{competencia.year}-{competencia.month:02d}/{empresa}"
        driver.page_to_pdf(
            file_path=diretorio,
            file_name=arquivo,
            hide_elements=("div#top", "form#formMenuTopo", "div#div-inscricao", "div#footer", "div.footer"),
            options={"scale": 1, "fitWindow": True, "paperWidth": 8.27, "paperHeight": 11.69},
        )
        return f"{diretorio}/{arquivo}"

    # -- planilha de entrada ----------------------------------------------

    def _acrescentar_na_planilha_fiscal(self, reader: XLSXReader, competencia: date | None, valor: str, row_number: int) -> None:
        """Acrescenta `valor` na coluna Y da planilha de entrada, SEM apagar o que já estava lá — a
        decisão do usuário foi manter tudo numa única célula, mesmo que fique com várias informações
        concatenadas. Cada acréscimo leva o rótulo "MM/AAAA: " na frente, exceto quando `competencia`
        é None (usado só para "S/ INSCRIÇÃO", que vale para todas as competências pedidas)."""
        self.check_thread_stopped_callback()
        if row_number <= 0:
            raise ValueError(f"Número da coluna não pode ser 0 ou menor. Recebido: `{row_number}`")

        cell = reader.get_cell(f"Y{row_number}")
        if isinstance(cell, MergedCell):
            raise TypeError(f"Não é possível atribuir valor `{valor}` à célula mesclada {cell.coordinate}")

        rotulo = f"{competencia.strftime('%m/%Y')}: {valor}" if competencia else valor
        atual = cell.value
        atual_str = atual.strip() if isinstance(atual, str) else ("" if atual is None else str(atual).strip())
        cell.value = f"{atual_str}; {rotulo}" if atual_str else rotulo
        reader.save_spreadsheet(close_workbook=False)

    # -- relatório de execução (um por competência) ------------------------

    def _gerar_relatorios_de_execucao(self) -> list[Path]:
        if not self.resultados:
            raise SemEmpresasProcessadasException("Não é possível gerar relatório: nenhuma empresa processada em nenhuma competência.")

        logger.info("Terminei o encerramento das empresas. Agora vou montar um relatório consolidado por competência.")

        por_competencia: dict[date, list[ResultadoMesEmpresa]] = {}
        for resultado in self.resultados:
            por_competencia.setdefault(resultado.competencia, []).append(resultado)

        caminhos: list[Path] = []
        for competencia in sorted(por_competencia):
            try:
                caminhos.append(self._gerar_relatorio_de_uma_competencia(competencia, por_competencia[competencia]))
            except SpreadsheetIsLockedException:
                logger.error(
                    f"Não consegui gerar o relatório de {competencia.strftime('%m/%Y')} porque o arquivo modelo "
                    "está aberto em outro programa — feche-o e gere esse relatório de novo depois."
                )
        return caminhos

    def _gerar_relatorio_de_uma_competencia(self, competencia: date, registros: list[ResultadoMesEmpresa]) -> Path:
        reader = XLSXReader(workbook_path=self.caminho_template_relatorio)

        now = datetime.now()
        now_str = now.strftime("%d/%m/%Y %H:%M:%S")
        reader.set_cell_value(row=4, col=2, value=now_str)
        reader.set_cell_value(row=5, col=2, value=AUTORES_RELATORIO)
        reader.set_cell_value(row=6, col=2, value=VERSAO_RELATORIO)
        reader.set_cell_value(row=7, col=2, value=f"Competência: {competencia.strftime('%m/%Y')}")

        l_processadas = len(registros)
        l_encerradas = sum(1 for r in registros if r.encerrada)
        l_problemas = l_processadas - l_encerradas

        reader.set_cell_value(row=5, col=3, value=l_processadas)
        reader.set_cell_value(row=5, col=5, value=l_encerradas)
        reader.set_cell_value(row=5, col=7, value=l_problemas)
        reader.set_cell_value(row=5, col=9, value=floor(safe_division(l_encerradas, l_processadas) * 100))

        linha_inicial = 10
        linha_offset = linha_inicial
        for r in registros:
            ws = reader._get_worksheet()
            ws.row_dimensions[linha_offset].height = 30

            detalhe = {
                "codigo": r.empresa.codigo,
                "nome": r.empresa.nome,
                "cnpj": int(remover_pontuacao_cnpj(r.empresa.cnpj)),
                "dt_encerramento": r.dt_encerramento.strftime(FORMATO_DATA) if r.dt_encerramento else None,
                "dt_processamento": now_str,
                "responsavel": r.empresa.responsavel,
                "problemas": " - ".join(r.problemas),
                "mensagens": len(r.empresa.mensagens),
            }
            if r.caminho_certificado:
                detalhe["certificado"] = str(r.caminho_certificado)

            coluna_offset = 1
            for key, det in detalhe.items():
                if det:
                    cell = reader.get_cell(row=linha_offset, col=coluna_offset)
                    cell.value = det
                    (reader.style_cell(cell).font(name="Arial", size=10)
                     .alignment(horizontal=HorizontalAlignment.CENTER, vertical=VerticalAlignment.CENTER))
                    match key.upper():
                        case "CNPJ":
                            reader.style_cell(cell).number_format(r'00"."000"."000"/"0000"-"00')
                        case _:
                            reader.style_cell(cell).number_format("General")
                coluna_offset += 1
            linha_offset += 1

        nome_arquivo = f"relatorio_execucao_automacao_{timestamp_as_file_name('xlsx')}"
        caminho_planilha = reader.save_spreadsheet(f"{self.caminho_saida}/{competencia.year}-{competencia.month:02d}/{nome_arquivo}")
        logger.success(f"Relatório de {competencia.strftime('%m/%Y')} pronto! Salvei em: {caminho_planilha}")
        return caminho_planilha
