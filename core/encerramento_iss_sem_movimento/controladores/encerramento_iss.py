from python_spreadsheet_reader.readers import XLSXReader, SpreadsheetIsLockedException
from python_spreadsheet_reader.styler import HorizontalAlignment, VerticalAlignment
from core.encerramento_iss_sem_movimento.iss.empresa import EmpresaSemMovimentoISSFortaleza, MensagemISS
from core.encerramento_iss_sem_movimento.iss.credenciais import CredenciaisISSFortaleza
from core.encerramento_iss_sem_movimento.python_webdriver.driver.chrome import ChromeDriver
from core.encerramento_iss_sem_movimento.python_webdriver.functions import retry_on_exception
from selenium.common.exceptions import ElementClickInterceptedException, StaleElementReferenceException
from selenium.common import ElementNotInteractableException, TimeoutException, NoSuchElementException
from loguru import logger
from pprint import pformat
from pathlib import Path
from core.encerramento_iss_sem_movimento.utils.functions import (
    remover_pontuacao_cnpj,
    aplicar_mascara_cnpj,
    timestamp_as_file_name,
    safe_division,
)
from core.encerramento_iss_sem_movimento.utils.classes import UserStoppedThreadException
from core.encerramento_iss_sem_movimento.utils.constants import FORMATO_DATA
from datetime import datetime, date, timedelta
from typing import Any
from pydantic import ValidationError
from openpyxl.cell import MergedCell
from enum import IntEnum
from collections.abc import Callable
from math import floor

# Metadados fixos usados no cabeçalho do relatório de execução — no projeto original
# (fusion/encerramento-iss-sem-movimento) esses valores vinham de um `pyproject.toml`
# próprio (via ControladorAmbiente.get_project_metadata()), que não existe mais aqui
# nem sobrevive ao empacotamento com PyInstaller. Fixados como constantes simples.
AUTORES_RELATORIO = "Barreira & Associados - Assessoria Contábil"
VERSAO_RELATORIO = "Automações ISS"


class ControladorEncerramentoISS:
    def __init__(
        self,
        caminho_planilha: Path,
        caminho_saida: Path,
        caminho_webdriver: Path,
        url_iss_fortaleza: str,
        credenciais: CredenciaisISSFortaleza,
        competencia: date,
        caminho_template_relatorio: Path,
        check_thread_stopped_callback: Callable[[], None]
    ) -> None:
        self.caminho_planilha = caminho_planilha
        self.caminho_saida = caminho_saida
        self.caminho_webdriver = caminho_webdriver
        self.url_iss_fortaleza = url_iss_fortaleza
        self.credenciais = credenciais
        self.competencia = competencia
        self.caminho_template_relatorio = caminho_template_relatorio
        self.check_thread_stopped_callback = check_thread_stopped_callback
        # Contadores
        self.processadas: list[EmpresaSemMovimentoISSFortaleza] = []
        self.problemas: list[EmpresaSemMovimentoISSFortaleza] = []
        self.encerradas: list[EmpresaSemMovimentoISSFortaleza] = []

    def executar_processo(self) -> "ResultadoEncerramentoISS":
        """Executa o processo de encerramento do ISS, escrituração das planilhas e geração de relatórios."""
        logger.info(
            "Vou começar a percorrer a planilha em busca das empresas sem movimento para encerrar o ISS delas."
        )

        # Iniciar webdriver
        driver = ChromeDriver(
            driver_path=self.caminho_webdriver
        )
        driver.start_driver(options=("--start-maximized",))

        # Buscar empresas para o fechamento do ISS
        reader = XLSXReader(self.caminho_planilha)

        status = StatusEncerramentoISS()

        try:
            self.check_thread_stopped_callback()
            self._login_iss_fortaleza(driver)

            for row_number, row in reader.lazy_load_sheet():
                self.check_thread_stopped_callback()
                try:
                    try:
                        empresa = EmpresaSemMovimentoISSFortaleza(
                            codigo=int(row[f"A{row_number}"].value),
                            nome=row[f"B{row_number}"].value,
                            cnpj=remover_pontuacao_cnpj(row[f"D{row_number}"].value),
                            responsavel=row[f"H{row_number}"].value,
                            municipio=row[f"W{row_number}"].value,
                            linha_planilha_fiscal=row_number
                        )
                    except KeyError as k_err:
                        logger.error(f"Essa linha da planilha não tem a coluna `{k_err}` que eu esperava, então vou pular para a próxima. Detalhe: {pformat(k_err)}")
                        continue
                    except ValidationError as v_err:
                        logger.debug(f"Os dados dessa empresa não passaram na validação — vou pular essa linha. Detalhe:\n{v_err.json()}")
                        continue
                    except (ValueError, TypeError):
                        continue  # Caso um dos campos da empresa for None ou outro tipo inesperado

                    self.check_thread_stopped_callback()

                    # Aplicar máscara no CNPJ
                    empresa.cnpj_m = aplicar_mascara_cnpj(empresa.cnpj)

                    self._procurar_inscricao_empresa(driver, reader, empresa)

                    self._dar_ciencia_nas_mensagens_nao_lidas(driver, empresa)

                    self._navegar_tela_escrituracao(driver)

                    self._preencher_widget_calendario(driver)

                    self._escriturar_empresa(driver, reader, empresa)

                    # Fim - Abrir modal de mudança de inscrição antes de ir para a próx. empresa
                    self._abrir_modal_de_alteracao_de_inscricao(driver)
                except ProximaEmpresaException:  # Levantado após erro tratado em alguma dos sub-métodos
                    continue

            msg = "Terminei de percorrer a planilha — o encerramento do ISS foi concluído."
            logger.success(msg)
            status = StatusEncerramentoISS(
                codigo=CodigoEncerramentoISS.SUCCESS,
                message=msg
            )
        except UserStoppedThreadException:
            msg = "Recebi o pedido de cancelamento do operador e parei o processo aqui."
            logger.warning(msg)
            status = StatusEncerramentoISS(
                codigo=CodigoEncerramentoISS.USER_ENDED_PROCESS,
                message=msg
            )
        except FileNotFoundError as fnf_err:
            msg = f"Não foi possível achar o arquivo: {fnf_err.filename}"
            logger.error(msg)
            status = StatusEncerramentoISS(
                codigo=CodigoEncerramentoISS.ERROR_FILE_NOT_FOUND,
                message=msg
            )
        except SpreadsheetIsLockedException as locked:
            logger.error(locked)
            status = StatusEncerramentoISS(
                codigo=CodigoEncerramentoISS.ERROR_SPREADSHEET_LOCKED,
                message=str(locked)
            )
        except Exception as exc:
            logger.error(exc)
            status = StatusEncerramentoISS(
                codigo=CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION,
                message=str(exc)
            )
        finally:
            reader.close_workbook()
            driver.quit_driver()

        result = ResultadoEncerramentoISS(
            status,
            self.processadas,
            self.encerradas,
            self.problemas,
            report_path=Path("")  # Será definido após gerar_relatorio_de_execucao()
        )

        try:
            report_path = self.gerar_relatorio_de_execucao(result)
            result.report_path = report_path
        except SpreadsheetIsLockedException:
            logger.error(
                "Não consegui gerar o relatório de execução porque o arquivo modelo está aberto em outro programa — feche-o e tente novamente."
            )
        except SemEmpresasProcessadasException as sem_empresa:
            logger.warning(sem_empresa)

        return result

    def _escriturar_na_planilha_fiscal(self, reader: XLSXReader, value: date | str, row_number: int):
        self.check_thread_stopped_callback()
        if row_number <= 0:
            raise ValueError(f"Número da coluna não pode ser 0 ou menor. Recebido: `{row_number}`")

        cell = reader.get_cell(f"Y{row_number}")
        if isinstance(cell, MergedCell):
            raise TypeError(f"Não é possível atribuir valor `{value}` à celula mesclada {cell.coordinate}")
        # Apenas escriturar se o campo da planilha estiver vazio
        if cell.value is None or (isinstance(cell.value, str) and cell.value.strip() == ""):
            cell.value = value
            reader.save_spreadsheet(close_workbook=False)

    def _login_iss_fortaleza(self, driver: ChromeDriver):
        logger.info("Estou entrando no portal da ISS Fortaleza com o CPF e a senha informados...")

        # Navegar ao site do ISS
        driver.goto(self.url_iss_fortaleza)

        driver.click(
            driver.find_element().by_attribute(attr_name="href", attr_value="/grpfor/oauth2/login")
        )

        # CNPJ
        driver.type(
            element=driver.find_element().by_id("username"),
            text=self.credenciais.cpf.get_secret_value()
        )

        # Senha
        driver.type(
            element=driver.find_element().by_id("password"),
            text=self.credenciais.senha.get_secret_value()
        )

        # Click no botão de login
        self.check_thread_stopped_callback()
        driver.find_element().by_id("kc-form-login").get_element().submit()

    def _procurar_inscricao_empresa(self, driver: ChromeDriver, reader: XLSXReader, empresa: EmpresaSemMovimentoISSFortaleza):
        logger.info(f"Agora é a vez da empresa {empresa}: vou procurar a inscrição dela no portal.")

        self.check_thread_stopped_callback()

        # Selecionar a opção CNPJ
        driver.click(
            driver.find_element().by_xpath('//table[@id="alteraInscricaoForm:tipoPesquisa"]//input[@value="CNPJ"]'),
        )

        # Após clicar na opção CNPJ, é realizado uma chamada AJAX para aplicar máscara de CPF/CNPJ
        # o que faz com que a caixa de pesquisa re-renderize e o texto digitado seja apagado, se inserido rápido demais.
        driver.explicit_wait(1)

        driver.clear(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"))
        driver.type(driver.find_element().by_id("alteraInscricaoForm:cpfPesquisa"), empresa.cnpj)

        driver.explicit_wait(1)

        # * Nota: antes de apertar o botão "Pesquisar", temos que verificar se a empresa já está na barra de resultados.
        # * Se a empresa já estiver listada, é necessário esperar um pouco mais após apertar o botão "Pesquisar".
        bln_empresa_listada = False

        record_xpath = f'//tbody[@id="alteraInscricaoForm:empresaDataTable:tb"]//a[normalize-space()="{empresa.cnpj_m}"]'

        # Verificar se empresa já está listada
        if driver.find_element().by_xpath(record_xpath).get_all_elements():
            bln_empresa_listada = True

        self.check_thread_stopped_callback()

        # Clicar no botão "Pesquisar"
        driver.click(
            driver.find_element().by_id("alteraInscricaoForm:btnPesquisar")
        )

        try:
            if bln_empresa_listada:
                driver.explicit_wait(5)

            # Procurar a linha da tabela de resultados que tem o CNPJ (com máscara) da empresa
            driver.click(
                driver.find_element().by_xpath(record_xpath),
                max_retries=5
            )
        except (TimeoutException, NoSuchElementException):
            logger.warning(
                f"Não encontrei procuração para a empresa {empresa} no portal — vou marcar como 'sem inscrição' e seguir para a próxima."
            )
            msg = "S/ INSCRIÇÃO"
            self._escriturar_na_planilha_fiscal(reader, value=msg, row_number=empresa.linha_planilha_fiscal)
            empresa.problemas.append(msg)
            self.problemas.append(empresa)
            self.processadas.append(empresa)
            raise ProximaEmpresaException

        # * ------------------- Alerta de Confirmação de Mudança de Empresa --------------------
        self._confirmar_alteracao_de_inscricao(driver)

        logger.info(
            f"Encontrei a procuração da empresa {empresa}. Entrando na página dela agora..."
        )

    def _navegar_tela_escrituracao(self, driver: ChromeDriver):
        # Pra ir até a tela de Escriturações, precisamos clicar nos botões do menu dropdown.
        # Nota: botões não tem nenhum identificador, então vamos procurar pelo texto interno.

        self.check_thread_stopped_callback()

        # Botão "Escriturações"
        driver.click(
            driver.find_element().by_xpath('//ul[contains(@class, "navbar-nav")]//a[normalize-space()="Escrituração"]')
        )

        # Opção "Manter Escrituração"
        driver.click(
            driver.find_element().by_xpath(
                '//ul[contains(@class, "navbar-nav")]//li[@class="dropdown open"]//ul[@class="dropdown-menu"]//a[normalize-space()="Manter Escrituração"]'
            )
        )

    def _preencher_widget_calendario(self, driver: ChromeDriver):
        # Preencher campo de data com mês anterior ao atual
        logger.debug(
            "Abrindo o calendário do portal para selecionar a competência do mês anterior..."
        )

        self.check_thread_stopped_callback()

        # Abrir widget de calendário
        driver.click(
            driver.find_element().by_class("rich-calendar-tool-btn")
        )

        # selecionar o mês no widget
        driver.click(
            driver.find_element().by_id(
                f"manterEscrituracaoForm:dataInicialDateEditorLayoutM{self.competencia.month - 1}"  # Mês (0-based)
            )
        )

        # selecionar o ano no widget
        driver.click(
            driver.find_element().by_xpath(
                f'//div[contains(@id, "manterEscrituracaoForm:dataInicialDateEditorLayoutY") and normalize-space()="{self.competencia.year}"]'
            )
        )

        self.check_thread_stopped_callback()

        # Clicar botão OK do widget
        driver.click(
            driver.find_element().by_id("manterEscrituracaoForm:dataInicialDateEditorButtonOk")
        )

    def _escriturar_empresa(self, driver: ChromeDriver, reader: XLSXReader, empresa: EmpresaSemMovimentoISSFortaleza):
        logger.debug("Consultando no portal se há escriturações em aberto para essa competência...")

        self.check_thread_stopped_callback()

        # Clicar botão "Consultar"
        driver.click(
            driver.find_element().by_id("manterEscrituracaoForm:btnConsultar")
        )

        # Procurar a linha da tabela de resultados que contém o registro relevante
        competencia_str = self.competencia.strftime("%m/%Y")

        def _locate_element():
            element = driver.find_element().by_xpath(
                f'//tbody[@id="manterEscrituracaoForm:dataTable:tb"]/tr/td/span[normalize-space()="{competencia_str}"]/../../self::tr'
            )
            element.get_element()  # Força a busca do elemento no DOM
            return element

        search_result_row = retry_on_exception(
            func=_locate_element,
            max_attempts=60,
            polling_seconds=1,
            exception=NoSuchElementException
        )

        link_escriturar = driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "linkEscriturar")]')

        # Verificar se ISS pode ser escritutado
        # title="Escriturar" => Pode ser escriturado;
        # title="Escrituração Encerrada" => Não pode ser escriturado
        link_escriturar_title: str | None = (
            link_escriturar
            .get_element()
            .get_dom_attribute("title")
        )

        if not link_escriturar_title:
            msg = f"O portal não me deu informação suficiente para saber se a escrituração da empresa {empresa} pode ser encerrada — vou marcar como problema e seguir para a próxima."
            logger.error(msg)

            empresa.problemas.append(msg)
            self.problemas.append(empresa)
            self.processadas.append(empresa)

            self._abrir_modal_de_alteracao_de_inscricao(driver)
            raise ProximaEmpresaException

        match link_escriturar_title.upper().strip():
            case "ESCRITURAÇÃO ENCERRADA":
                logger.warning(
                    f"A escrituração da empresa {empresa} já tinha sido encerrada antes — vou só baixar o certificado dela."
                )

                self.check_thread_stopped_callback()

                # Baixar certificado
                driver.click(
                    driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "certificado")]')
                )

                # Aguardar certificado ficar visível
                driver.click(
                    driver.find_element().by_id("formMenuTopo")
                )

                # Baixar PDF
                self._imprimir_declaracao_fechamento_iss(
                    driver, empresa
                )

                # Voltar 1 página porque o botão de mudar empresa foi escondido
                # para tirar o pdf do certificado
                driver.get_driver().back()

                # Data da 1ª vez que foi encerrado.
                dt_primeiro_encerramento: str = (
                    driver
                    .find_element(search_result_row)
                    .by_xpath('.//*[contains(@id, "dataEncerramento")]')
                    .get_element()
                    .text
                )

                empresa.dt_primeiro_encerramento = datetime.strptime(dt_primeiro_encerramento, FORMATO_DATA).date()

                self._escriturar_na_planilha_fiscal(
                    reader, dt_primeiro_encerramento, empresa.linha_planilha_fiscal
                )

                self.processadas.append(empresa)
                self.encerradas.append(empresa)
            case "ESCRITURAR":
                logger.info(f"A empresa {empresa} ainda está com a escrituração em aberto — vou verificar as pendências dela e encerrar.")

                self.check_thread_stopped_callback()

                driver.click(link_escriturar)

                # * ---------------------- Tela de Escrituração Fiscal -----------------------

                # Verificar se há serviços prestados
                self._verificar_servicos_prestados(driver, empresa)

                # Verificar se há serviços pendentes
                self._verificar_servicos_pendentes(driver, empresa)

                # Clicar no botão "Encerrar"
                btn_encerrar = driver.find_element().by_id("abaEncerramentoForm:btnEncerrarEscrituracao")

                driver.scroll_element_into_view(btn_encerrar)

                driver.click(btn_encerrar)

                # * ---------------------- Tela de Confirmação de Escrituração -----------------------

                logger.debug("Confirmando no portal o encerramento da escrituração...")

                self.check_thread_stopped_callback()

                driver.click(
                    driver.find_element().by_id("formEncerramento:btnSim")
                )

                # * ---------------------- Tela de Certificado de Encerramento -----------------------

                logger.info("Encerramento confirmado! Agora vou baixar o certificado dessa empresa.")

                self.check_thread_stopped_callback()

                # Botão "Certificado de Encerramento da Escrituração"
                driver.click(
                    driver.find_element().by_xpath('//input[contains(@id, "btnCertificadoEscrituracao")]')
                )

                # Aguardar o carregamento do document, senão screenshot virá vazio
                driver.find_element().by_id("docPrincipal")

                self._imprimir_declaracao_fechamento_iss(
                    driver, empresa
                )

                self.processadas.append(empresa)
                self.encerradas.append(empresa)

                empresa.dt_primeiro_encerramento = date.today()

                logger.success(
                    f"Pronto! Encerrei a escrituração da empresa {empresa} para a competência {self.competencia}."
                )

                # Voltar 1 página porque o botão de mudar empresa foi escondido
                # para tirar o pdf do certificado
                driver.get_driver().back()

    def _verificar_servicos_prestados(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza):
        logger.info("Antes de encerrar, vou conferir se essa empresa tem serviços prestados no período...")

        self.check_thread_stopped_callback()

        # Clicar na aba "Serviços Prestados"
        driver.click(
            driver.find_element().by_id("abaServicosPrestados_lbl")
        )

        def _aguardar_elemento_carregar():
            return (
                driver
                .find_element()
                .by_xpath(
                    "//table[@id='abaEncerramentoForm:dataTableServicosPrestados']"
                    "//td[normalize-space(text())='Somatório']/following-sibling::td[1]"
                )
                .get_element()
            )

        quantidade_cell = retry_on_exception(
            func=_aguardar_elemento_carregar,
            max_attempts=60,
            polling_seconds=1
        )

        try:
            quantidade = int(quantidade_cell.text.strip())
        except (TypeError, ValueError):
            msg = f"Não consegui ler direito a quantidade de serviços prestados da empresa {empresa} — o portal retornou um valor que eu não esperava: {quantidade_cell.text!r}"
            logger.error(msg)
            empresa.problemas.append(msg)
            self.processadas.append(empresa)

            self._abrir_modal_de_alteracao_de_inscricao(driver)
            raise ProximaEmpresaException

        if quantidade > 0:  # Empresa possui serviços prestados no período
            logger.warning(
                f"A empresa {empresa} tem serviços prestados no período, então não posso encerrar como sem movimento — vou marcar como problema e seguir."
            )

            empresa.problemas.append("SERVIÇOS PRESTADOS")
            self.processadas.append(empresa)

            self._abrir_modal_de_alteracao_de_inscricao(driver)
            raise ProximaEmpresaException

        self.check_thread_stopped_callback()
        # Voltar para a aba "Encerramento"
        driver.click(
            driver.find_element().by_id("abaEncerramento_lbl")
        )

    def _verificar_servicos_pendentes(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza):
        logger.info("Também vou conferir se há serviços pendentes registrados para essa empresa...")

        self.check_thread_stopped_callback()

        # Clicar na aba "Serviços Pendentes"
        driver.click(
            driver.find_element().by_id("aba_servicos_pendentes_lbl")
        )

        def _aguardar_elemento_carregar():
            return (
                driver
                .find_element()
                .by_id("servicos_pendentes_form:table_servico_tomados_pendente:tb")
                .get_element()
            )

        servicos_pendentes_table = retry_on_exception(
            func=_aguardar_elemento_carregar,
            max_attempts=60,
            polling_seconds=1
        )

        row_count = servicos_pendentes_table.get_property(
            "childElementCount"
        )

        # get_property() pode trazer vários tipos, certificar que é inteiro
        if not isinstance(row_count, int):
            msg = f"Não consegui confirmar se a empresa {empresa} tem serviços pendentes — o portal retornou um tipo de dado inesperado."
            logger.error(msg)
            empresa.problemas.append(msg)
            self.processadas.append(empresa)

            self._abrir_modal_de_alteracao_de_inscricao(driver)
            raise ProximaEmpresaException

        if row_count > 0:  # Empresa possui serviços pendentes
            logger.warning(
                f"A empresa {empresa} tem serviços pendentes, então não posso encerrar agora — vou marcar como problema e seguir para a próxima."
            )

            empresa.problemas.append("SERVIÇOS PENDENTES")
            self.processadas.append(empresa)

            self._abrir_modal_de_alteracao_de_inscricao(driver)
            raise ProximaEmpresaException

        self.check_thread_stopped_callback()
        # Voltar para a aba "Encerramento"
        driver.click(
            driver.find_element().by_id("abaEncerramento_lbl")
        )

    def _abrir_modal_de_alteracao_de_inscricao(self, driver: ChromeDriver) -> None:
        self.check_thread_stopped_callback()
        driver.click(
            driver.find_element().by_attribute(attr_name="title", attr_value="Alterar Inscrição Atual"),
        )

    @staticmethod
    def _confirmar_alteracao_de_inscricao(driver: ChromeDriver) -> None:
        try:
            driver.click(
                driver.find_element().by_id("alteraInscricaoForm:botaoOk"),
                max_retries=5
            )
        except (ElementNotInteractableException, NoSuchElementException, StaleElementReferenceException):
            pass  # Modal não apareceu

    def _dar_ciencia_nas_mensagens_nao_lidas(self, driver: ChromeDriver, emp: EmpresaSemMovimentoISSFortaleza) -> None:
        try:
            def _wait_for_messages_modal_to_popup():
                return (driver
                    .find_element()
                    .by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr")
                    .get_element())

            retry_on_exception(  # Se não tiver mensagens levanta exception, pulando o processo abaixo
                func=_wait_for_messages_modal_to_popup,
                max_attempts=15,
                polling_seconds=1,
                exception=NoSuchElementException
            )

            logger.info("Essa empresa tem mensagens não lidas no portal — vou abrir cada uma e dar ciência antes de continuar.")

            # Mover modal 200 pixels para cima porque o programa quebra se tentar clicar em um
            # elemento fora de vista.
            driver.drag_and_drop_by_offset(
                driver.find_element().by_id("mensagensModalHeader"),
                x_offset=0,
                y_offset=-200,
            )

            # Verificar se o modal de mensagens pendentes apareceu
            # Se tem mensagens, dar ciência em todas
            messages = (driver
                    .find_element()
                    .by_xpath("//*[@id='mensagensForm:mensagemDataTable:tb']//tr")
                    .get_all_elements())
            for msg in messages:
                # ! Não usar "msg" pois a referência DOM não é renovada ao entrar e sair do modal de anexos.
                self.check_thread_stopped_callback()

                driver.click(
                    driver.find_element().by_xpath(f'//a[contains(@id, "mensagensForm:mensagemDataTable:0:linkTitulo")]')
                )

                # * Decidi que seria útil coletar os dados da mensagem, caso no futuro
                # * seja necessário utilizá-los.
                def _collect_message_data() -> MensagemISS:
                    # Coletar dados da mensagem
                    title = driver.find_element().by_id("mensagensForm:titulo").get_element()
                    data = driver.find_element().by_id("mensagensForm:dataRegistro").get_element()
                    conteudo = driver.find_element().by_id("mensagensForm:descricao").get_element()
                    return MensagemISS(
                        title=title.text,
                        date=datetime.strptime(data.text, "%d/%m/%Y").date(),
                        content=conteudo.text
                    )

                m = retry_on_exception(
                    func=_collect_message_data,
                    max_attempts=15,
                    polling_seconds=1,
                )
                emp.mensagens.append(m)

                # Verificar sem tem anexo.
                # Precisa baixar todos os anexos antes de dar ciência.
                try:
                    anexos = (driver.find_element()
                            .by_xpath('//*[@id="mensagensForm:divAnexos"]//a[contains(@id, "linkVisualizarAnexo")]')
                            .get_all_elements())
                    logger.info("Essa mensagem tem anexos — vou baixar todos antes de dar ciência.")
                    for idx, anexo in enumerate(anexos):
                        # ! Não usar "anexo", mesmo motivo de "msg".
                        # * Existem dois tipos de anexos: 1) que abrem o anexo em um modal para leitura, e 2) que só baixa o anexo.
                        self.check_thread_stopped_callback()

                        try:
                            seletor_anexo = f"mensagensForm:anexos:{idx}:linkVisualizarAnexo"
                            # * Tentar clicar no 1° caso.
                            driver.click(
                                driver.find_element().by_id(seletor_anexo),
                                max_retries=5
                            )
                            # Baixar anexo
                            btn_baixar_anexo = driver.find_element().by_id("mensagensForm:botaoBaixarAnexo")
                            driver.click(btn_baixar_anexo)
                            driver.explicit_wait(.5)
                            # Voltar p/ form da mensagem
                            driver.click(
                                driver.find_element().by_id("mensagensForm:botaoVoltarAnexo")
                            )
                        except (ElementClickInterceptedException, NoSuchElementException, StaleElementReferenceException):
                            # * Tentar clicar no 2° caso.
                            seletor_anexo = f"mensagensForm:anexos:{idx}:linkBaixarAnexo"
                            driver.click(
                                driver.find_element().by_id(seletor_anexo),
                                max_retries=5
                            )

                        driver.explicit_wait(1)
                        # * Em mensagems com muitos anexos, o botão de baixar anexo some da tela,
                        # * então precisamos esconder o link do anexo que foi baixado para poder clicar no botão de ciência.
                        driver.hide_elements((f"#{seletor_anexo}",))
                except (TimeoutException, StaleElementReferenceException, NoSuchElementException):
                    pass # Não tem anexo

                self.check_thread_stopped_callback()

                # Dar ciência
                btn_dar_ciencia = driver.find_element().by_id("mensagensForm:botaoDarCiencia")
                driver.click(btn_dar_ciencia)
        except NoSuchElementException:
            pass

    def _imprimir_declaracao_fechamento_iss(self, driver: ChromeDriver, emp: EmpresaSemMovimentoISSFortaleza) -> str:
        self.check_thread_stopped_callback()

        # Renderizar página como PDF e baixar arquivo
        empresa = f"EMP_{emp.codigo}"
        arquivo = f"CERTIFICADO ISS_EMP{emp.codigo}_{timestamp_as_file_name('pdf')}"
        diretorio = f"{self.caminho_saida}/{self.competencia.month:02d}/{empresa}"
        driver.page_to_pdf(
            file_path=diretorio,
            file_name=arquivo,
            hide_elements=(
                "div#top",
                "form#formMenuTopo",
                "div#div-inscricao",
                "div#footer",
                "div.footer",
            ),
            options={
                "scale": 1,
                "fitWindow": True,
                "paperWidth": 8.27,  # A4 dimemsions
                "paperHeight": 11.69,  # A4 dimemsions
            }
        )

        caminho_arquivo = f"{diretorio}/{arquivo}"

        emp.caminho_certificado = Path(caminho_arquivo)

        return caminho_arquivo

    def gerar_relatorio_de_execucao(
        self,
        resultado_encerramento: "ResultadoEncerramentoISS",
    ) -> Path:
        logger.info("Terminei o encerramento das empresas. Agora vou montar o relatório consolidado da execução.")

        if len(resultado_encerramento.processadas) <= 0:
            raise SemEmpresasProcessadasException(
                "Não é possível gerar o relatório de execução: Nenhuma empresa processada."
            )

        # Abrir template
        reader = XLSXReader(
            workbook_path=self.caminho_template_relatorio,
        )

        # Cabeçalho
        now = datetime.now()
        now_str = now.strftime("%d/%m/%Y %H:%M:%S")
        reader.set_cell_value(row=4, col=2, value=now_str)

        # Autores e versão (ver constantes no topo do arquivo)
        reader.set_cell_value(row=5, col=2, value=AUTORES_RELATORIO)
        reader.set_cell_value(row=6, col=2, value=VERSAO_RELATORIO)

        # Competência
        competencia = now - timedelta(days=30)
        reader.set_cell_value(
            row=7,
            col=2,
            value=f"Competência: {competencia.strftime('%m/%Y')}",
        )

        processadas = resultado_encerramento.processadas
        l_processadas = len(processadas)
        encerradas = resultado_encerramento.encerradas
        l_encerradas = len(encerradas)
        problemas = resultado_encerramento.problemas
        l_problemas = len(problemas)

        # Resumo execução
        reader.set_cell_value(row=5, col=3, value=l_processadas)
        reader.set_cell_value(row=5, col=5, value=l_encerradas)
        reader.set_cell_value(row=5, col=7, value=l_problemas)
        reader.set_cell_value(row=5, col=9, value=floor(safe_division(l_encerradas, l_processadas) * 100))

        # Detalhamento por empresa
        linha_inicial = 10
        linha_offset = linha_inicial
        for emp in processadas:
            ws = reader._get_worksheet()
            # Padronizar altura das linhas
            ws.row_dimensions[linha_offset].height = 30

            detalhe = {
                "codigo": emp.codigo,
                "nome": emp.nome,
                "cnpj": int(remover_pontuacao_cnpj(emp.cnpj)),
                "dt_encerramento": (
                    emp.dt_primeiro_encerramento.strftime(FORMATO_DATA)
                    if emp.dt_primeiro_encerramento
                    else None
                ),
                "dt_processamento": now_str,
                "responsavel": emp.responsavel,
                "problemas": " - ".join(emp.problemas),
                "mensagens": len(emp.mensagens),
            }

            if emp.caminho_certificado:
                caminho_cert = self.caminho_saida / emp.caminho_certificado.relative_to(
                    self.caminho_saida
                )
                detalhe["certificado"] = str(caminho_cert)

            coluna_inicial = 1
            coluna_offset = coluna_inicial
            for key, det in detalhe.items():
                if det:
                    cell = reader.get_cell(row=linha_offset, col=coluna_offset)
                    cell.value = det

                    # Estilos
                    # Me deparei com um problema onde os estilos do template estavam sendo resetados
                    # após setar o valor da célula, então resolvi não perder mais tempo e apliquei
                    # os estilos programaticamente abaixo.
                    (reader.style_cell(cell)
                     .font(name="Arial", size=10)
                     .alignment(horizontal=HorizontalAlignment.CENTER, vertical=VerticalAlignment.CENTER))

                    # Formatações personalizadas
                    match key.upper():
                        case "CNPJ":
                            reader.style_cell(cell).number_format(r'00"."000"."000"/"0000"-"00')
                        case _:
                            reader.style_cell(cell).number_format("General")

                coluna_offset += 1

            linha_offset += 1

        # Salvar planilha
        nome_arquivo = f"relatorio_execucao_automacao_{timestamp_as_file_name('xlsx')}"
        caminho_planilha = reader.save_spreadsheet(
            f"{self.caminho_saida}/{self.competencia.month:02d}/{nome_arquivo}"
        )
        logger.success(f"Relatório pronto! Salvei em: {caminho_planilha}")
        return caminho_planilha


class CodigoEncerramentoISS(IntEnum):
    UNEXPECTED_END = -1
    SUCCESS = 0
    USER_ENDED_PROCESS = 1
    ERROR = 2
    ERROR_SPREADSHEET_LOCKED = 3
    ERROR_FILE_NOT_FOUND = 4
    ERROR_UNHANDLED_EXCEPTION = 99  # ! Sempre no fim


class StatusEncerramentoISS:
    def __init__(
            self,
            codigo: CodigoEncerramentoISS = CodigoEncerramentoISS.UNEXPECTED_END,  # Por segurança, assume erro até provar o contrário
            message: str = "Programa terminou inesperadamente"
    ) -> None:
        self.codigo = codigo
        self.message = message

    def __repr__(self) -> str:
        return f"Status: {self.codigo.name} ({self.codigo.value}): {self.message}"


class ResultadoEncerramentoISS:
    def __init__(
            self,
            status: StatusEncerramentoISS,
            processadas: list[EmpresaSemMovimentoISSFortaleza],
            encerradas: list[EmpresaSemMovimentoISSFortaleza],
            problemas: list[EmpresaSemMovimentoISSFortaleza],
            report_path: Path,
    ) -> None:
        self.processadas = processadas
        self.encerradas = encerradas
        self.problemas = problemas
        self.status = status
        self.report_path = report_path


# * ------------------- Exceptions -------------------


class ProximaEmpresaException(Exception):
    """Levantado caso haja algum problema com a inscrituração de uma empresa."""
    pass


class SemEmpresasProcessadasException(Exception):
    pass
