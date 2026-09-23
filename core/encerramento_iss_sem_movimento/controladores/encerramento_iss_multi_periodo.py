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
  como problema daquela competência e o robô segue para a próxima (competência ou empresa),
  refazendo o caminho pelo menu (competência) ou voltando à troca de empresa (empresa) antes de seguir;
- se o Chrome for fechado/travar, ou se o portal não puder ser devolvido à troca de empresa depois de
  um erro, a execução PARA com uma mensagem clara (e gera os relatórios do que já foi feito), em vez de
  tentar cada empresa restante com o navegador morto;
- a mesma informação não é acrescentada duas vezes na célula da planilha ao repetir a execução;
- Pausar/Continuar/Parar nos limites seguros (entre competências e entre empresas);
- competência já encerrada cujo certificado já existe na pasta não baixa o certificado de novo;
- modo "apenas verificar" (não encerra nem baixa nada) e execução restrita ao que foi verificado;
- as empresas podem vir da planilha fiscal (só as sem movimento de Fortaleza, como no original) OU de uma lista
  de CNPJs digitados pelo operador — neste caso não há planilha: nada é gravado na coluna Y, o nome da empresa
  vem da linha de resultado da busca de inscrição no portal e os certificados vão para `CNPJ_<14 dígitos>/`.

ATENÇÃO — duplicação consciente: uma correção futura no fluxo do portal (seletores, navegação)
pode precisar ser feita nos dois módulos (este e `encerramento_iss.py`).
"""

from __future__ import annotations

import re
import time
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
    InvalidSessionIdException,
    NoSuchWindowException,
    StaleElementReferenceException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from urllib3.exceptions import HTTPError as Urllib3HTTPError

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


def cnpj_valido(valor: str) -> bool:
    """CNPJ com 14 dígitos, não todos iguais e com os dois dígitos verificadores corretos (aceita máscara)."""
    cnpj = re.sub(r"\D", "", str(valor or ""))
    if len(cnpj) != 14 or cnpj == cnpj[0] * 14:
        return False

    def digito(base: str, pesos: list[int]) -> int:
        resto = sum(int(d) * p for d, p in zip(base, pesos)) % 11
        return 0 if resto < 2 else 11 - resto

    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    d1 = digito(cnpj[:12], pesos1)
    d2 = digito(cnpj[:12] + str(d1), [6] + pesos1)
    return cnpj[12:] == f"{d1}{d2}"


@dataclass
class EmpresaAlvo:
    """Empresa informada direto pelo operador (sem planilha): só o CNPJ é obrigatório. O nome é preenchido
    com o que o portal mostrar na busca de inscrição; código e responsável não existem nesta origem."""

    cnpj: str  # 14 dígitos
    nome: str = ""
    codigo: int | None = None
    responsavel: str = ""
    linha_planilha_fiscal: int | None = None  # None = sem planilha (nada é gravado na coluna Y)
    cnpj_m: str = ""
    mensagens: list = field(default_factory=list)
    problemas: list = field(default_factory=list)

    def __str__(self) -> str:
        return f"{self.nome} (CNPJ: {self.cnpj})" if self.nome else f"CNPJ {aplicar_mascara_cnpj(self.cnpj)}"


def identificador_empresa(empresa: Any) -> str:
    """Nome da pasta do certificado: `EMP_<código>` (origem planilha) ou `CNPJ_<14 dígitos>` (origem manual)."""
    codigo = getattr(empresa, "codigo", None)
    return f"EMP_{codigo}" if codigo else f"CNPJ_{empresa.cnpj}"


MODO_ENCERRAR = "encerrar"
MODO_VERIFICAR = "verificar"  # percorre e confere tudo, mas NÃO encerra nem baixa nada

# O que aconteceu com uma empresa/competência (`ResultadoMesEmpresa.categoria`).
ACAO_ENCERRADA_AGORA = "encerrada_agora"
ACAO_JA_ENCERRADA_CERT_EXISTENTE = "ja_encerrada_certificado_existente"
ACAO_JA_ENCERRADA_CERT_BAIXADO = "ja_encerrada_certificado_baixado"
ACAO_JA_ENCERRADA_SEM_CERT = "ja_encerrada_sem_certificado"  # só na verificação: falta baixar o certificado
ACAO_APTA = "apta"  # só na verificação: aberta e sem impedimento para encerrar
ACAO_PROBLEMA = "problema"


@dataclass
class ResultadoMesEmpresa:
    empresa: EmpresaSemMovimentoISSFortaleza
    competencia: date
    encerrada: bool
    dt_encerramento: date | None = None
    problemas: list[str] = field(default_factory=list)
    caminho_certificado: Path | None = None
    situacao: str = ""  # texto da coluna "Situação" do portal (ex.: "Aberta - Normal", "Fechada - Normal")
    acao: str = ""  # uma das ACAO_*; vazio = deduzir de `encerrada` (compatível com quem não informa)

    @property
    def categoria(self) -> str:
        if self.acao:
            return self.acao
        return ACAO_ENCERRADA_AGORA if self.encerrada else ACAO_PROBLEMA


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
        """Competências com ISS encerrado ao final da execução (encerradas agora + já estavam encerradas)."""
        return sum(1 for r in self.resultados if r.encerrada)

    @property
    def encerradas_agora(self) -> int:
        return sum(1 for r in self.resultados if r.categoria == ACAO_ENCERRADA_AGORA)

    @property
    def ja_encerradas(self) -> int:
        return sum(1 for r in self.resultados if r.categoria.startswith("ja_encerrada"))

    @property
    def aptas(self) -> int:
        return sum(1 for r in self.resultados if r.categoria == ACAO_APTA)

    @property
    def problemas(self) -> int:
        return sum(1 for r in self.resultados if r.categoria == ACAO_PROBLEMA)


class SemEmpresasProcessadasException(Exception):
    """Nenhuma empresa foi processada para nenhuma competência (nada a gerar de relatório)."""


class NavegadorIndisponivelException(Exception):
    """O Chrome foi fechado (ou parou de responder) durante a execução: continuar só geraria uma
    sequência de falhas, uma por empresa restante."""


class RecuperacaoImpossivelException(Exception):
    """Depois de um erro inesperado não consegui devolver o portal à tela de troca de empresa:
    seguir para a próxima empresa faria todas as seguintes falharem em cascata."""


# Textos que o Selenium/Chrome usam quando a sessão do navegador morreu.
_MARCAS_NAVEGADOR_INDISPONIVEL = (
    "invalid session id",
    "not connected to devtools",
    "disconnected",
    "no such window",
    "target window already closed",
    "web view not found",
    "chrome not reachable",
    "max retries exceeded",
)

# Campos do período da tela "Manter Escrituração". O portal lista as competências de `dataInicial` até
# `dataFinal` (padrão: mês atual), em ordem decrescente e paginadas de 12 em 12 — com só a data inicial
# ajustada, competências mais antigas que a 12ª linha caíam na página 2 e o robô não as achava. Fixando
# as duas datas na competência a tabela devolve exatamente uma linha, sempre na primeira página.
CAMPOS_PERIODO_CONSULTA = (
    "manterEscrituracaoForm:dataInicialInputDate",
    "manterEscrituracaoForm:dataInicialInputCurrentDate",
    "manterEscrituracaoForm:dataFinalInputDate",
    "manterEscrituracaoForm:dataFinalInputCurrentDate",
)
ID_BOTAO_CONSULTAR = "manterEscrituracaoForm:btnConsultar"
ID_CAMPO_PESQUISA_INSCRICAO = "alteraInscricaoForm:cpfPesquisa"
ID_PROGRESSO_PORTAL = "mpProgressoContentTable"

# Textos curtos que vão para a planilha fiscal (coluna Y) e para o relatório — sem nome de exceção do Selenium.
MSG_COMPETENCIA_NAO_ENCONTRADA = "COMPETÊNCIA NÃO ENCONTRADA NO PORTAL"
MSG_ERRO_PORTAL = "ERRO NO PORTAL"


def navegador_indisponivel(exc: BaseException) -> bool:
    """True se `exc` indica que o navegador foi fechado/travou (sessão do Chrome perdida)."""
    if isinstance(exc, (InvalidSessionIdException, NoSuchWindowException, Urllib3HTTPError, ConnectionError)):
        return True
    if isinstance(exc, WebDriverException):
        texto = (getattr(exc, "msg", None) or str(exc)).lower()
        return any(marca in texto for marca in _MARCAS_NAVEGADOR_INDISPONIVEL)
    return False


# ---------------------------------------------------------------------------
# Controlador
# ---------------------------------------------------------------------------


class ControladorEncerramentoISSMultiPeriodo:
    def __init__(
        self,
        caminho_planilha: Path | None,
        caminho_saida: Path,
        caminho_webdriver: Path,
        url_iss_fortaleza: str,
        credenciais: CredenciaisISSFortaleza,
        competencias: list[date],
        caminho_template_relatorio: Path,
        check_thread_stopped_callback: Callable[[], None],
        modo: str = MODO_ENCERRAR,
        callback_pausa: Callable[[], None] | None = None,
        callback_parar: Callable[[], bool] | None = None,
        alvos: dict[str, set[date]] | None = None,
        cnpjs: list[str] | None = None,
    ) -> None:
        """`callback_pausa` bloqueia enquanto o operador estiver com a execução pausada e `callback_parar`
        devolve True quando ele pediu para parar — ambos consultados só nos limites seguros (entre
        competências e entre empresas). `alvos` (CNPJ com 14 dígitos -> competências) restringe a execução
        ao que foi aprovado numa verificação prévia. `cnpjs` (origem manual) substitui a planilha: informe
        `caminho_planilha` OU `cnpjs`, nunca os dois."""
        if not competencias:
            raise ValueError("Informe ao menos uma competência.")
        if modo not in (MODO_ENCERRAR, MODO_VERIFICAR):
            raise ValueError(f"Modo inválido: {modo!r}")
        lista_cnpjs: list[str] | None = None
        if cnpjs:
            lista_cnpjs = []
            for bruto in cnpjs:
                cnpj = re.sub(r"\D", "", str(bruto))
                if len(cnpj) != 14 or cnpj == cnpj[0] * 14:
                    raise ValueError(f"CNPJ inválido: {bruto!r}")
                if cnpj not in lista_cnpjs:  # sem duplicatas, na ordem informada
                    lista_cnpjs.append(cnpj)
        if (caminho_planilha is None) == (lista_cnpjs is None):
            raise ValueError("Informe a planilha fiscal OU a lista de CNPJs (exatamente uma das duas origens).")
        self.cnpjs = lista_cnpjs
        self.caminho_planilha = caminho_planilha
        self.caminho_saida = caminho_saida
        self.caminho_webdriver = caminho_webdriver
        self.url_iss_fortaleza = url_iss_fortaleza
        self.credenciais = credenciais
        # Ordem cronológica e sem duplicatas, independente da ordem em que o operador as adicionou.
        self.competencias: list[date] = sorted(set(competencias))
        self.caminho_template_relatorio = caminho_template_relatorio
        self.check_thread_stopped_callback = check_thread_stopped_callback
        self.modo = modo
        self.callback_pausa = callback_pausa
        self.callback_parar = callback_parar
        self.alvos = alvos
        self.resultados: list[ResultadoMesEmpresa] = []

    def _ponto_seguro(self) -> None:
        """Limite seguro (nenhum encerramento em andamento): espera se estiver pausado e, se o operador pediu
        para parar, interrompe levantando o mesmo sinal do cancelamento (o `executar_processo` gera os
        relatórios do que já foi feito e fecha o navegador)."""
        if self.callback_pausa is not None:
            self.callback_pausa()
        if self.callback_parar is not None and self.callback_parar():
            raise UserStoppedThreadException()

    def _competencias_da_empresa(self, empresa: EmpresaSemMovimentoISSFortaleza) -> list[date]:
        if self.alvos is None:
            return self.competencias
        escolhidas = self.alvos.get(empresa.cnpj, set())
        return [c for c in self.competencias if c in escolhidas]

    # -- fluxo principal -------------------------------------------------

    def executar_processo(self) -> ResultadoEncerramentoISSMultiPeriodo:
        rotulos = ", ".join(c.strftime("%m/%Y") for c in self.competencias)
        if self.cnpjs is not None:
            origem = f"os {len(self.cnpjs)} CNPJ(s) informados"
        else:
            origem = "a planilha em busca das empresas sem movimento"
        if self.modo == MODO_VERIFICAR:
            logger.info(
                f"Vou percorrer {origem} e VERIFICAR (sem encerrar nem baixar nada) "
                f"em {len(self.competencias)} competência(s): {rotulos}."
            )
        else:
            logger.info(
                f"Vou percorrer {origem} para encerrar o ISS delas "
                f"em {len(self.competencias)} competência(s): {rotulos}."
            )

        driver = ChromeDriver(driver_path=self.caminho_webdriver)
        driver.start_driver(options=("--start-maximized",))

        reader = XLSXReader(self.caminho_planilha) if self.cnpjs is None else None

        status = StatusEncerramentoISS()

        try:
            self.check_thread_stopped_callback()
            self._login_iss_fortaleza(driver)

            if self.cnpjs is None:
                for row_number, row in reader.lazy_load_sheet():
                    self.check_thread_stopped_callback()
                    self._ponto_seguro()
                    self._executar_isolado(
                        driver, f"na linha {row_number} da planilha",
                        lambda: self._processar_linha(driver, reader, row_number, row),
                    )
            else:
                for cnpj in self.cnpjs:
                    self.check_thread_stopped_callback()
                    self._ponto_seguro()
                    empresa = EmpresaAlvo(cnpj=cnpj)
                    self._executar_isolado(
                        driver, f"na empresa {aplicar_mascara_cnpj(cnpj)}",
                        lambda: self._processar_empresa(driver, None, empresa),
                    )

            msg = (
                "Terminei de percorrer a planilha para todas as competências solicitadas."
                if self.cnpjs is None
                else "Terminei de percorrer os CNPJs informados para todas as competências solicitadas."
            )
            logger.success(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.SUCCESS, message=msg)
        except UserStoppedThreadException:
            msg = "Recebi o pedido de cancelamento do operador e parei o processo aqui."
            logger.warning(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.USER_ENDED_PROCESS, message=msg)
        except NavegadorIndisponivelException:
            msg = (
                "O navegador (Chrome) foi fechado ou parou de responder durante a execução — parei o processo aqui. "
                "O que já foi processado está na planilha fiscal e nos relatórios."
            )
            logger.error(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION, message=msg)
        except RecuperacaoImpossivelException as impossivel:
            msg = str(impossivel)
            logger.error(msg)
            status = StatusEncerramentoISS(codigo=CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION, message=msg)
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
            if reader is not None:
                reader.close_workbook()
            try:
                driver.quit_driver()
            except Exception as exc:  # noqa: BLE001 — com o Chrome já fechado o quit() pode falhar; não pode impedir o relatório
                logger.debug(f"Não consegui fechar o navegador de forma limpa (provavelmente já estava fechado): {exc}")

        report_paths: list[Path] = []
        if self.modo == MODO_ENCERRAR:  # a verificação não encerra nada: não há relatório de encerramento a gerar
            try:
                report_paths = self._gerar_relatorios_de_execucao()
            except SemEmpresasProcessadasException as sem_empresa:
                logger.warning(sem_empresa)

        return ResultadoEncerramentoISSMultiPeriodo(status, self.resultados, report_paths)

    def _executar_isolado(self, driver: ChromeDriver, descricao: str, executar: Callable[[], None]) -> None:
        """Roda o processamento de UMA empresa isolando erros inesperados: registra, devolve o portal à
        troca de empresa e segue para a próxima (ou para tudo, se for o navegador que morreu)."""
        try:
            executar()
        except (UserStoppedThreadException, NavegadorIndisponivelException, RecuperacaoImpossivelException):
            raise
        except Exception as exc:  # noqa: BLE001 — um erro inesperado numa empresa não pode abortar a execução inteira
            if navegador_indisponivel(exc):
                raise NavegadorIndisponivelException(str(exc)) from exc
            logger.error(f"Erro inesperado {descricao} — vou pular para a próxima. Detalhe: {exc}")
            # O erro pode ter deixado o portal numa tela qualquer; sem voltar à troca de empresa,
            # todas as empresas seguintes falhariam já no primeiro clique.
            self._recuperar_para_proxima_empresa(driver)

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

        self._processar_empresa(driver, reader, empresa)

    def _processar_empresa(self, driver: ChromeDriver, reader: XLSXReader | None, empresa: Any) -> None:
        """Processa todas as competências de uma empresa (vinda da planilha ou digitada pelo operador)."""
        competencias = self._competencias_da_empresa(empresa)
        if not competencias:
            return  # execução restrita a uma lista aprovada e esta empresa não está nela

        self.check_thread_stopped_callback()
        empresa.cnpj_m = aplicar_mascara_cnpj(empresa.cnpj)

        if not self._procurar_inscricao_empresa(driver, empresa):
            logger.warning(
                f"Não encontrei procuração para a empresa {empresa} no portal — vou marcar como 'sem inscrição' "
                "para todas as competências e seguir para a próxima."
            )
            msg = "S/ INSCRIÇÃO"
            for competencia in competencias:
                self.resultados.append(ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg]))
            self._acrescentar_na_planilha_fiscal(reader, competencia=None, valor=msg, row_number=empresa.linha_planilha_fiscal)
            return  # Sem inscrição: não há como trocar de empresa pelo modal (ele não abriu)

        self._dar_ciencia_nas_mensagens_nao_lidas(driver, empresa)
        self._navegar_tela_escrituracao(driver)

        for competencia in competencias:
            self.check_thread_stopped_callback()
            self._ponto_seguro()
            try:
                resultado = self._processar_competencia(driver, reader, empresa, competencia)
            except UserStoppedThreadException:
                raise
            except Exception as exc:  # noqa: BLE001 — um erro numa competência não pode derrubar as demais da empresa
                if navegador_indisponivel(exc):
                    raise
                logger.error(
                    f"{empresa} — {competencia.strftime('%m/%Y')}: não consegui processar essa competência no portal "
                    f"({type(exc).__name__}) — vou seguir para a próxima."
                )
                resultado = ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[MSG_ERRO_PORTAL])
                self._acrescentar_na_planilha_fiscal(reader, competencia, MSG_ERRO_PORTAL, empresa.linha_planilha_fiscal)
                # O erro pode ter acontecido no meio do fluxo (aba de serviços, certificado…): refaz o
                # caminho pelo menu para a próxima competência começar de uma tela conhecida.
                self.resultados.append(resultado)
                self._navegar_tela_escrituracao(driver)
                continue
            self.resultados.append(resultado)

        self._abrir_modal_de_alteracao_de_inscricao(driver)

    def _modal_de_troca_de_inscricao_esta_aberto(self, driver: ChromeDriver) -> bool:
        campos = driver.get_driver().find_elements(By.ID, ID_CAMPO_PESQUISA_INSCRICAO)
        return any(campo.is_displayed() for campo in campos)

    def _aguardar_modal_de_troca(self, driver: ChromeDriver, segundos: int) -> bool:
        for _ in range(max(1, segundos)):
            if self._modal_de_troca_de_inscricao_esta_aberto(driver):
                return True
            self.check_thread_stopped_callback()
            time.sleep(1)
        return self._modal_de_troca_de_inscricao_esta_aberto(driver)

    def _garantir_modal_de_troca(self, driver: ChromeDriver) -> None:
        """Só segue quando a busca de inscrição estiver visível. Depois de uma troca de tela o modal pode
        demorar (ou ter sido fechado pelo recarregamento da página): espera um pouco, tenta abri-lo pelo
        botão "Alterar Inscrição Atual" e só então desiste — antes o robô clicava no CNPJ com o modal
        fechado e perdia 60 s por empresa."""
        if self._aguardar_modal_de_troca(driver, 10):
            return
        logger.debug("A tela de troca de empresa não abriu sozinha — vou abri-la pelo botão do topo.")
        self._abrir_modal_de_alteracao_de_inscricao(driver)
        if not self._aguardar_modal_de_troca(driver, 30):
            raise NoSuchElementException("A tela de troca de empresa não ficou visível.")

    def _recuperar_para_proxima_empresa(self, driver: ChromeDriver) -> None:
        """Depois de um erro inesperado, devolve o portal à tela de troca de empresa (se já não estiver
        nela). Se não conseguir, levanta `RecuperacaoImpossivelException` para parar em vez de deixar
        todas as empresas seguintes falharem em cascata."""
        try:
            if self._modal_de_troca_de_inscricao_esta_aberto(driver):
                return
            self._abrir_modal_de_alteracao_de_inscricao(driver)
        except UserStoppedThreadException:
            raise
        except Exception as exc:  # noqa: BLE001
            if navegador_indisponivel(exc):
                raise NavegadorIndisponivelException(str(exc)) from exc
            raise RecuperacaoImpossivelException(
                "Depois do erro não consegui voltar para a tela de troca de empresa — parei o processo aqui para não "
                "gerar uma sequência de falhas nas próximas empresas. O que já foi processado está na planilha fiscal "
                "e nos relatórios."
            ) from exc

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
        self._garantir_modal_de_troca(driver)

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
            if not getattr(empresa, "nome", "x"):  # origem manual: o nome vem da linha de resultado da busca
                self._ler_nome_da_linha(driver, record_xpath, empresa)
            driver.click(driver.find_element().by_xpath(record_xpath), max_retries=5)
        except (TimeoutException, NoSuchElementException):
            return False

        self._confirmar_alteracao_de_inscricao(driver)
        logger.info(f"Encontrei a procuração da empresa {empresa}. Entrando na página dela agora...")
        return True

    @staticmethod
    def _ler_nome_da_linha(driver: ChromeDriver, record_xpath: str, empresa: Any) -> None:
        """Preenche `empresa.nome` com a razão social da linha de resultado (CNPJ, inscrição, razão social).
        Só informativo: nunca levanta e não espera mais que alguns segundos."""
        try:
            for _ in range(5):
                linhas = driver.get_driver().find_elements(By.XPATH, record_xpath + "/ancestor::tr[1]")
                if linhas:
                    texto = " ".join((linhas[0].text or "").split())
                    texto = texto.removeprefix(empresa.cnpj_m).strip()
                    texto = re.sub(r"^\d{4,8}-\d\s+", "", texto)  # inscrição municipal
                    if texto:
                        empresa.nome = texto
                    return
                time.sleep(1)
        except Exception:  # noqa: BLE001
            return

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
        self._aguardar_tela_de_lista(driver)

    def _tela_de_lista_esta_aberta(self, driver: ChromeDriver) -> bool:
        botoes = driver.get_driver().find_elements(By.ID, ID_BOTAO_CONSULTAR)
        return any(botao.is_displayed() for botao in botoes)

    def _aguardar_tela_de_lista(self, driver: ChromeDriver) -> None:
        """Espera o formulário de Manter Escrituração carregar (o período agora é definido por script, que
        não espera o elemento existir como os cliques fazem)."""

        def _esperar():
            if not self._tela_de_lista_esta_aberta(driver):
                raise NoSuchElementException("A tela de Manter Escrituração ainda não carregou.")
            return True

        retry_on_exception(func=_esperar, max_attempts=60, polling_seconds=1, exception=NoSuchElementException)

    def _garantir_tela_de_lista(self, driver: ChromeDriver) -> None:
        """Se o robô não estiver na lista de escriturações (ex.: ficou na tela de uma competência que
        terminou com problema), refaz o caminho pelo menu antes de consultar a próxima competência."""
        if self._tela_de_lista_esta_aberta(driver):
            return
        logger.debug("O portal não está na lista de escriturações — voltando pelo menu Escrituração > Manter Escrituração.")
        self._navegar_tela_escrituracao(driver)

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

    def _definir_periodo_consulta(self, driver: ChromeDriver, competencia: date) -> None:
        """Fixa data inicial = data final = a competência (ver `CAMPOS_PERIODO_CONSULTA`)."""
        valor = competencia.strftime("%m/%Y")
        logger.debug(f"Definindo o período da consulta do portal como {valor} (data inicial e final)...")

        self.check_thread_stopped_callback()

        for campo in CAMPOS_PERIODO_CONSULTA:
            definido = driver.get_driver().execute_script(
                "var e = document.getElementById(arguments[0]); if (!e) { return false; } e.value = arguments[1]; return true;",
                campo,
                valor,
            )
            if not definido:
                raise NoSuchElementException(f"Campo de período não encontrado na tela: {campo}")

    def _garantir_aba_encerramento(self, driver: ChromeDriver) -> None:
        abas = driver.get_driver().find_elements(By.ID, "abaEncerramento_lbl")
        if abas and "rich-tab-active" not in (abas[0].get_attribute("class") or ""):
            driver.click(driver.find_element().by_id("abaEncerramento_lbl"))

    def _verificar_servicos_prestados(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza) -> str | None:
        """Devolve o motivo do problema (sem gravar nada), ou None se não houver serviços prestados.

        A tabela do "Somatório" (quantidade de serviços prestados) fica DENTRO da aba "Encerramento", que é a
        aba aberta ao entrar na tela. Trocar para a aba "Serviços Prestados" esconde essa tabela e o texto lido
        vem vazio — por isso a leitura é feita sem trocar de aba, esperando o valor aparecer."""
        logger.info("Antes de encerrar, vou conferir se essa empresa tem serviços prestados no período...")

        self.check_thread_stopped_callback()
        self._garantir_aba_encerramento(driver)

        def _ler_quantidade() -> str:
            celula = driver.find_element().by_xpath(
                "//table[@id='abaEncerramentoForm:dataTableServicosPrestados']"
                "//td[normalize-space(text())='Somatório']/following-sibling::td[1]"
            ).get_element()
            texto = (celula.text or "").strip()
            if not texto:  # célula ainda oculta/vazia (texto de elemento oculto vem vazio no Selenium)
                raise NoSuchElementException("O Somatório ainda não tem texto visível.")
            return texto

        motivo: str | None = None
        try:
            texto_quantidade = retry_on_exception(func=_ler_quantidade, max_attempts=60, polling_seconds=1, exception=NoSuchElementException)
        except NoSuchElementException:
            logger.error(
                f"Não consegui ler a quantidade de serviços prestados da empresa {empresa} — "
                "o portal não mostrou o valor a tempo."
            )
            return "ERRO AO LER SERVIÇOS PRESTADOS"

        try:
            quantidade = int(texto_quantidade)
            if quantidade > 0:
                logger.warning(f"A empresa {empresa} tem serviços prestados no período — vou marcar como problema.")
                motivo = "SERVIÇOS PRESTADOS"
        except (TypeError, ValueError):
            logger.error(
                f"Não consegui ler direito a quantidade de serviços prestados da empresa {empresa} — "
                f"o portal retornou um valor que eu não esperava: {texto_quantidade!r}"
            )
            motivo = "ERRO AO LER SERVIÇOS PRESTADOS"

        return motivo

    def _aguardar_progresso_do_portal(self, driver: ChromeDriver, segundos: int = 30) -> None:
        """Espera o indicador "carregando" do portal sumir (nunca levanta): contar linhas de uma tabela que
        ainda está sendo preenchida por AJAX daria "sem pendências" por engano."""
        for _ in range(segundos):
            indicadores = driver.get_driver().find_elements(By.ID, ID_PROGRESSO_PORTAL)
            if not any(i.is_displayed() for i in indicadores):
                return
            self.check_thread_stopped_callback()
            time.sleep(1)

    def _verificar_servicos_pendentes(self, driver: ChromeDriver, empresa: EmpresaSemMovimentoISSFortaleza) -> str | None:
        """Devolve o motivo do problema (sem gravar nada), ou None se não houver serviços pendentes."""
        logger.info("Também vou conferir se há serviços pendentes registrados para essa empresa...")

        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id("aba_servicos_pendentes_lbl"))

        def _aguardar_elemento_carregar():
            return driver.find_element().by_id("servicos_pendentes_form:table_servico_tomados_pendente:tb").get_element()

        servicos_pendentes_table = retry_on_exception(func=_aguardar_elemento_carregar, max_attempts=60, polling_seconds=1)
        self._aguardar_progresso_do_portal(driver)
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
        self._garantir_tela_de_lista(driver)
        self._definir_periodo_consulta(driver, competencia)

        logger.debug("Consultando no portal se há escriturações em aberto para essa competência...")
        self.check_thread_stopped_callback()
        driver.click(driver.find_element().by_id(ID_BOTAO_CONSULTAR))

        competencia_str = competencia.strftime("%m/%Y")

        def _locate_element():
            element = driver.find_element().by_xpath(
                f'//tbody[@id="manterEscrituracaoForm:dataTable:tb"]/tr/td/span[normalize-space()="{competencia_str}"]/../../self::tr'
            )
            element.get_element()
            return element

        try:
            search_result_row = retry_on_exception(func=_locate_element, max_attempts=60, polling_seconds=1, exception=NoSuchElementException)
        except NoSuchElementException:
            logger.error(
                f"{empresa} — {competencia_str}: o portal não devolveu a linha dessa competência depois da consulta — "
                "vou seguir para a próxima."
            )
            self._acrescentar_na_planilha_fiscal(reader, competencia, MSG_COMPETENCIA_NAO_ENCONTRADA, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[MSG_COMPETENCIA_NAO_ENCONTRADA])
        link_escriturar = driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "linkEscriturar")]')
        link_escriturar_title: str | None = link_escriturar.get_element().get_dom_attribute("title")

        situacao = self._ler_situacao(search_result_row, competencia_str)
        if situacao:
            logger.info(f"{empresa} — {competencia_str}: situação no portal: {situacao}.")

        if not link_escriturar_title:
            msg = "O portal não me deu informação suficiente para saber se a escrituração pode ser encerrada."
            logger.error(f"{empresa} — {competencia_str}: {msg}")
            self._acrescentar_na_planilha_fiscal(reader, competencia, msg, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg])

        match link_escriturar_title.upper().strip():
            case "ESCRITURAÇÃO ENCERRADA":
                resultado = self._processar_ja_encerrada(driver, reader, empresa, competencia, search_result_row)
            case "ESCRITURAR":
                resultado = self._processar_a_encerrar(driver, reader, empresa, competencia, link_escriturar)
            case _:
                msg = f"Título inesperado do botão de escriturar: {link_escriturar_title!r}"
                logger.error(f"{empresa} — {competencia_str}: {msg}")
                self._acrescentar_na_planilha_fiscal(reader, competencia, msg, empresa.linha_planilha_fiscal)
                return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[msg], situacao=situacao)

        resultado.situacao = situacao
        return resultado

    @staticmethod
    def _ler_situacao(search_result_row: Any, competencia_str: str) -> str:
        """Texto da linha da competência sem o rótulo "MM/AAAA" (ex.: "Aberta - Normal 01/09/2026"). Nunca levanta."""
        try:
            texto = " ".join(search_result_row.get_element().text.split())
            return texto.removeprefix(competencia_str).strip()
        except Exception:  # noqa: BLE001 — só informativo
            return ""

    def _diretorio_certificados(self, empresa: EmpresaSemMovimentoISSFortaleza, competencia: date) -> Path:
        # <ano>-<mês>, diferente do original (só <mês>): evita colidir competências do mesmo mês em anos diferentes.
        return Path(str(self.caminho_saida)) / f"{competencia.year}-{competencia.month:02d}" / identificador_empresa(empresa)

    def _certificado_existente(self, empresa: EmpresaSemMovimentoISSFortaleza, competencia: date) -> Path | None:
        """Certificado (PDF não vazio) já baixado para essa empresa/competência na pasta de saída, se houver."""
        try:
            candidatos = [
                p for p in self._diretorio_certificados(empresa, competencia).glob("CERTIFICADO ISS_*.pdf")
                if p.is_file() and p.stat().st_size > 0
            ]
        except OSError:
            return None
        return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None

    @staticmethod
    def _ler_data_encerramento(driver: ChromeDriver, search_result_row: Any) -> tuple[str, date]:
        dt_texto: str = driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "dataEncerramento")]').get_element().text
        return dt_texto, datetime.strptime(dt_texto, FORMATO_DATA).date()

    def _processar_ja_encerrada(
        self,
        driver: ChromeDriver,
        reader: XLSXReader,
        empresa: EmpresaSemMovimentoISSFortaleza,
        competencia: date,
        search_result_row: Any,
    ) -> ResultadoMesEmpresa:
        competencia_str = competencia.strftime("%m/%Y")
        existente = self._certificado_existente(empresa, competencia)

        if existente is not None or self.modo == MODO_VERIFICAR:
            # Não precisa (ou, na verificação, não deve) baixar nada: só lê a data de encerramento.
            dt_texto, dt_encerramento = self._ler_data_encerramento(driver, search_result_row)
            if existente is not None:
                logger.info(
                    f"A escrituração da empresa {empresa} já estava encerrada ({competencia_str}) e o certificado "
                    "já existe na pasta de saída — não vou baixar de novo."
                )
                acao = ACAO_JA_ENCERRADA_CERT_EXISTENTE
            else:
                logger.warning(
                    f"A escrituração da empresa {empresa} já estava encerrada ({competencia_str}), mas o certificado "
                    "ainda não está na pasta de saída (a execução vai baixá-lo)."
                )
                acao = ACAO_JA_ENCERRADA_SEM_CERT
            self._acrescentar_na_planilha_fiscal(reader, competencia, dt_texto, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(
                empresa, competencia, encerrada=True, dt_encerramento=dt_encerramento,
                caminho_certificado=existente, acao=acao,
            )

        logger.warning(
            f"A escrituração da empresa {empresa} já tinha sido encerrada antes ({competencia_str}) e o certificado "
            "ainda não está na pasta — vou baixar o certificado."
        )

        self.check_thread_stopped_callback()
        driver.click(driver.find_element(search_result_row).by_xpath('.//*[contains(@id, "certificado")]'))
        driver.click(driver.find_element().by_id("formMenuTopo"))  # aguarda o certificado ficar visível

        caminho_certificado = self._imprimir_declaracao_fechamento_iss(driver, empresa, competencia)

        driver.get_driver().back()  # o botão de mudar empresa fica escondido enquanto o certificado é exibido

        dt_texto, dt_encerramento = self._ler_data_encerramento(driver, search_result_row)

        self._acrescentar_na_planilha_fiscal(reader, competencia, dt_texto, empresa.linha_planilha_fiscal)

        return ResultadoMesEmpresa(
            empresa, competencia, encerrada=True, dt_encerramento=dt_encerramento,
            caminho_certificado=Path(caminho_certificado), acao=ACAO_JA_ENCERRADA_CERT_BAIXADO,
        )

    def _processar_a_encerrar(
        self,
        driver: ChromeDriver,
        reader: XLSXReader,
        empresa: EmpresaSemMovimentoISSFortaleza,
        competencia: date,
        link_escriturar: Any,
    ) -> ResultadoMesEmpresa:
        competencia_str = competencia.strftime("%m/%Y")
        if self.modo == MODO_VERIFICAR:
            logger.info(f"A empresa {empresa} está com a escrituração em aberto ({competencia_str}) — vou só verificar as pendências (não vou encerrar).")
        else:
            logger.info(f"A empresa {empresa} ainda está com a escrituração em aberto ({competencia_str}) — vou verificar as pendências e encerrar.")

        self.check_thread_stopped_callback()
        driver.click(link_escriturar)

        motivo = self._verificar_servicos_prestados(driver, empresa)
        if motivo is None:
            motivo = self._verificar_servicos_pendentes(driver, empresa)

        if motivo is not None:
            self._acrescentar_na_planilha_fiscal(reader, competencia, motivo, empresa.linha_planilha_fiscal)
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, problemas=[motivo])

        if self.modo == MODO_VERIFICAR:
            logger.info(f"Verificação: a competência {competencia_str} da empresa {empresa} está apta a ser encerrada (nada foi encerrado).")
            return ResultadoMesEmpresa(empresa, competencia, encerrada=False, acao=ACAO_APTA)

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

        return ResultadoMesEmpresa(
            empresa, competencia, encerrada=True, dt_encerramento=dt_encerramento,
            caminho_certificado=Path(caminho_certificado), acao=ACAO_ENCERRADA_AGORA,
        )

    def _imprimir_declaracao_fechamento_iss(self, driver: ChromeDriver, emp: EmpresaSemMovimentoISSFortaleza, competencia: date) -> str:
        self.check_thread_stopped_callback()

        sufixo = f"EMP{emp.codigo}" if getattr(emp, "codigo", None) else f"CNPJ{emp.cnpj}"
        arquivo = f"CERTIFICADO ISS_{sufixo}_{timestamp_as_file_name('pdf')}"
        diretorio = self._diretorio_certificados(emp, competencia).as_posix()
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
        if reader is None or row_number is None:
            return  # origem manual: não há planilha onde gravar (o registro fica nos relatórios)
        if self.modo == MODO_VERIFICAR:
            return  # verificar não deixa rastro na planilha: só registra quando o encerramento/baixa realmente acontece
        self.check_thread_stopped_callback()
        if row_number <= 0:
            raise ValueError(f"Número da coluna não pode ser 0 ou menor. Recebido: `{row_number}`")

        cell = reader.get_cell(f"Y{row_number}")
        if isinstance(cell, MergedCell):
            raise TypeError(f"Não é possível atribuir valor `{valor}` à célula mesclada {cell.coordinate}")

        rotulo = f"{competencia.strftime('%m/%Y')}: {valor}" if competencia else valor
        atual = cell.value
        atual_str = atual.strip() if isinstance(atual, str) else ("" if atual is None else str(atual).strip())
        # Repetir a mesma execução não pode empilhar a mesma informação na célula.
        if rotulo in (parte.strip() for parte in atual_str.split(";")):
            return
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
