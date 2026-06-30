"""Automação visível do portal da ISS de Fortaleza.

Este módulo concentra o fluxo do navegador usando Selenium WebDriver para manter
`extrair_nf_pdfs.py` responsável apenas pela extração dos PDFs e pela orquestração
da CLI.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from selenium import webdriver
from selenium.common.exceptions import (
    ElementNotInteractableException,
    JavascriptException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from tratamento_erros import registrar_evento_execucao, configurar_pasta_logs

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

URL_ISS_FORTALEZA = "https://iss.fortaleza.ce.gov.br/grpfor/home.seam"
PORTA_DEBUG_CHROME_PADRAO = 9222
PERFIL_CHROME_FUNCAO2_PADRAO = Path("brain/navegador_funcao2_profile")

# Mapeamentos de mês por nome (normalizado sem acento)
MESES_POR_NOME = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

# Abreviações dos meses usadas no calendário do portal
MESES_ABREVIADOS_PORTAL = {
    1: "jan",
    2: "fev",
    3: "mar",
    4: "abr",
    5: "mai",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "set",
    10: "out",
    11: "nov",
    12: "dez",
}


# ---------------------------------------------------------------------------
# Estruturas de dados
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CompetenciaTrabalho:
    """Representa a competência trabalhada no portal."""

    mes: int
    ano: int
    rotulo: str

    @property
    def nome_mes(self) -> str:
        nomes = [
            "",
            "Janeiro",
            "Fevereiro",
            "Março",
            "Abril",
            "Maio",
            "Junho",
            "Julho",
            "Agosto",
            "Setembro",
            "Outubro",
            "Novembro",
            "Dezembro",
        ]
        return nomes[self.mes]


@dataclass(frozen=True)
class DocumentoPortalISS:
    """Linha da planilha que será transportada para o formulário do portal."""

    arquivo_pdf: str
    prefeitura: str
    cnpj_prestador: str
    numero_nf: str
    id_cnae_final: str
    data_emissao: str
    descricao_servico: str
    uf_local_prestacao: str
    cidade_local_prestacao: str
    natureza_operacao: str
    iss_retido: str
    valor_servico: str
    valor_deducoes: str = ""
    descontos_incondicionados: str = ""
    descontos_condicionados: str = ""
    outras_retencoes: str = ""
    ir: str = ""
    pis_nao_retido: str = ""
    cofins_nao_retido: str = ""
    csrf: str = ""
    inss: str = ""


# ---------------------------------------------------------------------------
# Exceções customizadas
# ---------------------------------------------------------------------------


class PrestadorNaoEncontradoError(RuntimeError):
    """Sinaliza que o portal não encontrou uma razão social para o CNPJ informado."""

    def __init__(self, cnpj: str, detalhe_tela: str = "") -> None:
        self.cnpj = cnpj
        self.detalhe_tela = detalhe_tela.strip()
        mensagem = f"Nenhuma Razão Social Encontrada para o CNPJ {cnpj}."
        if self.detalhe_tela:
            mensagem = f"{mensagem} Detalhe da tela: {self.detalhe_tela}"
        super().__init__(mensagem)


class CompetenciaSemEscriturarDisponivelError(RuntimeError):
    """Sinaliza que a competência consultada não liberou o botão de escriturar."""


# ---------------------------------------------------------------------------
# Utilitários de texto
# ---------------------------------------------------------------------------


def normalizar_texto(texto: str) -> str:
    """Remove acentos e padroniza o texto para facilitar o parse."""

    normalizado = unicodedata.normalize("NFD", texto)
    return "".join(
        caractere for caractere in normalizado if unicodedata.category(caractere) != "Mn"
    ).lower()


def limpar_cnpj_para_digitacao(cnpj: str) -> str:
    """Retorna apenas os dígitos do CNPJ para uma digitação mais estável."""

    return re.sub(r"\D+", "", cnpj or "")


def _formatar_celula_para_string_de_valor(valor: Any) -> str:
    """Normaliza valores numéricos lidos do Excel para formato de string brasileiro."""
    if valor is None:
        return "0,00"
    if isinstance(valor, (int, float)):
        return f"{valor:.2f}".replace(".", ",")
    
    texto = str(valor).strip()
    if not texto:
        return "0,00"
    # Se for string formatada em padrão americano com ponto decimal, ex: "1500.50"
    if "." in texto and "," not in texto:
        try:
            val_f = float(texto)
            return f"{val_f:.2f}".replace(".", ",")
        except ValueError:
            pass
    return texto


def limpar_valor_para_digitacao(valor: str) -> str:
    """Retorna o valor em formato digitável para o campo mascarado do portal."""

    texto = (valor or "").strip()
    texto = texto.replace("R$", "").strip()
    texto = texto.replace(".", "")
    return texto


def limpar_numero_para_digitacao(numero: str) -> str:
    """Mantém apenas dígitos para o campo de número do documento."""

    return re.sub(r"\D+", "", numero or "")


def _somente_digitos(texto: str) -> str:
    """Normaliza qualquer valor para comparação por dígitos."""

    return re.sub(r"\D+", "", texto or "")


def _texto_indica_prestador_nao_encontrado(texto: str) -> bool:
    """Reconhece mensagens visíveis que indicam ausência de razão social no portal."""

    texto_normalizado = normalizar_texto(texto or "")
    marcadores = (
        "nenhuma razao social encontrada",
        "nenhuma razao social",
        "nenhum resultado encontrado",
        "nao foram encontrados",
    )
    return any(marcador in texto_normalizado for marcador in marcadores)


# ---------------------------------------------------------------------------
# Logging da função 2
# ---------------------------------------------------------------------------


def registrar_log_funcao2(caminho_log: Path, mensagem: str) -> None:
    """Registra eventos da função 2 com data e hora para rastrear as linhas ignoradas."""

    caminho_log.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with caminho_log.open("a", encoding="utf-8") as arquivo:
        arquivo.write(f"[{timestamp}] {mensagem}\n")


# ---------------------------------------------------------------------------
# Competência
# ---------------------------------------------------------------------------


def construir_competencia(mes: int, ano: int) -> CompetenciaTrabalho:
    """Valida mês e ano e devolve a competência estruturada."""

    if not 1 <= mes <= 12:
        raise ValueError("O mês deve estar entre 1 e 12.")
    if ano < 1000 or ano > 9999:
        raise ValueError("O ano deve estar no formato AAAA.")
    return CompetenciaTrabalho(mes=mes, ano=ano, rotulo=f"{mes:02d}/{ano}")


def interpretar_competencia(texto: str) -> CompetenciaTrabalho:
    """Converte uma entrada do usuário em mês e ano.

    Aceita entradas como `maio 2026`, `maio/2026` ou `05/2026`.
    """

    texto_limpo = texto.strip()
    if not texto_limpo:
        raise ValueError("Competência vazia.")

    padrao_numerico = re.fullmatch(r"(\d{1,2})\s*(?:/|\s)\s*(\d{4})", texto_limpo)
    if padrao_numerico:
        mes = int(padrao_numerico.group(1))
        ano = int(padrao_numerico.group(2))
        return construir_competencia(mes, ano)

    texto_norm = normalizar_texto(texto_limpo)
    padrao_nome = re.fullmatch(r"([a-z]+)\s+de\s+(\d{4})", texto_norm) or re.fullmatch(
        r"([a-z]+)\s+(\d{4})", texto_norm
    )
    if not padrao_nome:
        raise ValueError(
            "Informe a competência no formato 'maio 2026' ou '05/2026'."
        )

    nome_mes = padrao_nome.group(1)
    ano = int(padrao_nome.group(2))
    mes = MESES_POR_NOME.get(nome_mes)
    if mes is None:
        raise ValueError(
            "Não foi possível identificar o mês informado. Use, por exemplo, 'maio 2026'."
        )
    return construir_competencia(mes, ano)


def solicitar_competencia() -> CompetenciaTrabalho:
    """Pergunta ao usuário a competência antes de abrir o navegador."""

    while True:
        try:
            mes = int(input("Informe o mês de trabalho (1 a 12): ").strip())
            ano_texto = input("Informe o ano de trabalho (AAAA): ").strip()
            if not re.fullmatch(r"\d{4}", ano_texto):
                raise ValueError("O ano deve conter exatamente 4 dígitos.")
            return construir_competencia(mes, int(ano_texto))
        except ValueError as exc:
            print(f"Entrada inválida: {exc}")
            print()


def solicitar_nova_competencia_para_repetir(
    competencia_atual: CompetenciaTrabalho,
) -> CompetenciaTrabalho:
    """Pergunta ao usuário outra competência para reiniciar a seleção De/Até."""

    while True:
        texto = input(
            f"A competência {competencia_atual.rotulo} está sem o botão Escriturar habilitado. "
            "Informe outra competência para voltar aos campos De e Até (ou pressione Enter para cancelar): "
        ).strip()
        if not texto:
            raise SystemExit(
                "A automação foi interrompida porque não foi informada uma nova competência."
            )
        try:
            return interpretar_competencia(texto)
        except ValueError as exc:
            print(f"Competência inválida: {exc}")
            print()


def solicitar_planilha_automacao() -> Path:
    """Pergunta ao usuário qual planilha XLSX será usada na automação."""

    texto = input(
        "Informe o caminho da planilha XLSX de automação (Enter para usar nf_compilado.xlsx): "
    ).strip().strip('"')
    return Path(texto) if texto else Path("nf_compilado.xlsx")


# ---------------------------------------------------------------------------
# Carregamento da planilha
# ---------------------------------------------------------------------------


def carregar_documentos_xlsx(caminho_xlsx: Path) -> list[DocumentoPortalISS]:
    """Lê a planilha `nf_compilado.xlsx` e devolve as linhas prontas para o portal."""

    workbook = load_workbook(caminho_xlsx, data_only=True)
    planilha = workbook.active

    cabecalhos = [celula.value for celula in next(planilha.iter_rows(min_row=1, max_row=1))]
    colunas = {nome: indice for indice, nome in enumerate(cabecalhos)}

    campos_obrigatorios = [
        "ARQUIVO_PDF",
        "PREFEITURA",
        "CNPJ_PRESTADOR",
        "NUMERO_NF",
        "ID_CNAE_FINAL",
        "DATA_EMISSAO",
        "DESCRICAO_SERVICO",
        "UF_LOCAL_PRESTACAO",
        "CIDADE_LOCAL_PRESTACAO",
        "NATUREZA_OPERACAO",
        "ISS_RETIDO",
        "VALOR_SERVICO",
    ]
    faltantes = [campo for campo in campos_obrigatorios if campo not in colunas]
    if faltantes:
        raise ValueError(
            f"A planilha informada não contém as colunas obrigatórias: {', '.join(faltantes)}"
        )

    documentos: list[DocumentoPortalISS] = []
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        if not any(valor is not None and str(valor).strip() for valor in linha):
            continue

        valor_servico = _formatar_celula_para_string_de_valor(linha[colunas["VALOR_SERVICO"]]) if "VALOR_SERVICO" in colunas else "0,00"
        valor_deducoes = _formatar_celula_para_string_de_valor(linha[colunas["VALOR_DEDUCOES"]]) if "VALOR_DEDUCOES" in colunas else "0,00"
        descontos_incondicionados = _formatar_celula_para_string_de_valor(linha[colunas["DESCONTOS_INCONDICIONADOS"]]) if "DESCONTOS_INCONDICIONADOS" in colunas else "0,00"
        descontos_condicionados = _formatar_celula_para_string_de_valor(linha[colunas["DESCONTOS_CONDICIONADOS"]]) if "DESCONTOS_CONDICIONADOS" in colunas else "0,00"
        outras_retencoes = _formatar_celula_para_string_de_valor(linha[colunas["OUTRAS_RETENCOES"]]) if "OUTRAS_RETENCOES" in colunas else "0,00"
        ir = _formatar_celula_para_string_de_valor(linha[colunas["IR"]]) if "IR" in colunas else "0,00"
        pis_nao_retido = _formatar_celula_para_string_de_valor(linha[colunas["PIS_NAO_RETIDO"]]) if "PIS_NAO_RETIDO" in colunas else "0,00"
        cofins_nao_retido = _formatar_celula_para_string_de_valor(linha[colunas["COFINS_NAO_RETIDO"]]) if "COFINS_NAO_RETIDO" in colunas else "0,00"
        csrf = _formatar_celula_para_string_de_valor(linha[colunas["CSRF (CSLL + PIS + Cofins Retidos)"]]) if "CSRF (CSLL + PIS + Cofins Retidos)" in colunas else "0,00"
        inss = _formatar_celula_para_string_de_valor(linha[colunas["INSS"]]) if "INSS" in colunas else "0,00"

        documentos.append(
            DocumentoPortalISS(
                arquivo_pdf=str(linha[colunas["ARQUIVO_PDF"]] or ""),
                prefeitura=str(linha[colunas["PREFEITURA"]] or ""),
                cnpj_prestador=str(linha[colunas["CNPJ_PRESTADOR"]] or ""),
                numero_nf=str(linha[colunas["NUMERO_NF"]] or ""),
                id_cnae_final=str(linha[colunas["ID_CNAE_FINAL"]] or ""),
                data_emissao=str(linha[colunas["DATA_EMISSAO"]] or ""),
                descricao_servico=str(linha[colunas["DESCRICAO_SERVICO"]] or ""),
                uf_local_prestacao=str(linha[colunas["UF_LOCAL_PRESTACAO"]] or ""),
                cidade_local_prestacao=str(linha[colunas["CIDADE_LOCAL_PRESTACAO"]] or ""),
                natureza_operacao=str(linha[colunas["NATUREZA_OPERACAO"]] or ""),
                iss_retido=str(linha[colunas["ISS_RETIDO"]] or ""),
                valor_servico=valor_servico,
                valor_deducoes=valor_deducoes,
                descontos_incondicionados=descontos_incondicionados,
                descontos_condicionados=descontos_condicionados,
                outras_retencoes=outras_retencoes,
                ir=ir,
                pis_nao_retido=pis_nao_retido,
                cofins_nao_retido=cofins_nao_retido,
                csrf=csrf,
                inss=inss,
            )
        )
    return documentos


# ---------------------------------------------------------------------------
# Inicialização do navegador Chrome com Selenium
# ---------------------------------------------------------------------------


def _criar_opcoes_chrome(
    depuracao: bool = False,
    perfil_persistente: Path | None = None,
) -> Options:
    """Monta as opções do Chrome para o Selenium, incluindo proteções anti-detecção."""

    opcoes = Options()
    opcoes.add_argument("--start-maximized")
    opcoes.add_argument("--no-first-run")
    opcoes.add_argument("--no-default-browser-check")
    # Oculta a flag de automação no navegador para evitar bloqueios do portal
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_experimental_option("excludeSwitches", ["enable-automation"])
    opcoes.add_experimental_option("useAutomationExtension", False)
    if perfil_persistente:
        perfil_persistente.mkdir(parents=True, exist_ok=True)
        opcoes.add_argument(f"--user-data-dir={perfil_persistente.resolve()}")
    return opcoes


def abrir_navegador_visivel(depuracao: bool = False) -> WebDriver:
    """Abre um Chrome visível e maximizado, gerenciado pelo webdriver-manager."""

    opcoes = _criar_opcoes_chrome(depuracao=depuracao)
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=opcoes)
    # Remove o atributo webdriver para evitar a detecção pelo portal
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.get(URL_ISS_FORTALEZA)
    return driver


def abrir_navegador_com_perfil_persistente(
    perfil_persistente: Path,
    depuracao: bool = False,
) -> WebDriver:
    """Abre o Chrome com perfil de usuário persistente para manter login entre execuções."""

    opcoes = _criar_opcoes_chrome(depuracao=depuracao, perfil_persistente=perfil_persistente)
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=opcoes)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.get(URL_ISS_FORTALEZA)
    registrar_evento_execucao(
        f"Navegador com perfil persistente aberto: {perfil_persistente.resolve()}",
        "ISS Fortaleza",
    )
    return driver


# ---------------------------------------------------------------------------
# Esperas e cliques com Selenium
# ---------------------------------------------------------------------------


def _aguardar_elemento(
    driver: WebDriver,
    by: str,
    seletor: str,
    timeout: int = 10,
) -> WebElement:
    """Aguarda um elemento ficar visível e retorna o WebElement."""

    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, seletor))
    )


def _aguardar_elemento_clicavel(
    driver: WebDriver,
    by: str,
    seletor: str,
    timeout: int = 10,
) -> WebElement:
    """Aguarda um elemento ficar clicável (visível + habilitado) e o retorna."""

    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, seletor))
    )


def _clicar_com_espera(
    driver: WebDriver,
    by: str,
    seletor: str,
    descricao: str,
    timeout: int = 10,
) -> WebElement:
    """Aguarda o elemento e efetua o clique, com fallback via JavaScript e retry para elementos obsoletos."""

    for tentativa in range(3):
        try:
            elemento = _aguardar_elemento_clicavel(driver, by, seletor, timeout=timeout)
            try:
                elemento.click()
            except (ElementNotInteractableException, WebDriverException) as exc:
                if isinstance(exc, StaleElementReferenceException):
                    raise
                # Fallback via JavaScript para elementos que bloqueiam o clique nativo
                driver.execute_script("arguments[0].click();", elemento)
            return elemento
        except StaleElementReferenceException:
            if tentativa == 2:
                raise
            time.sleep(0.5)


def _clicar_por_id(driver: WebDriver, elemento_id: str, timeout: int = 10) -> None:
    """Clica em um elemento pelo ID com fallback via JavaScript."""

    _clicar_com_espera(driver, By.ID, elemento_id, f"elemento ID={elemento_id}", timeout=timeout)


def _pausar_para_depuracao(
    driver: WebDriver,
    habilitado: bool,
    mensagem: str,
) -> None:
    """Interrompe o fluxo para inspeção manual quando o modo de depuração estiver ativo."""

    if not habilitado:
        return
    print(mensagem)
    if sys.stdin.isatty():
        input(f"{mensagem}\nPressione Enter para continuar com o próximo clique...")


# ---------------------------------------------------------------------------
# Digitação e seleção de campos
# ---------------------------------------------------------------------------


def _digitar_campo(
    driver: WebDriver,
    by: str,
    seletor: str,
    texto: str,
    delay_ms: int = 80,
) -> None:
    """Localiza o campo, limpa e digita o texto tecla por tecla com pequena pausa entre elas.

    Possui lógica de retry para se recuperar caso o elemento mude sob AJAX durante a digitação.
    """

    for tentativa in range(3):
        try:
            campo = _aguardar_elemento_clicavel(driver, by, seletor, timeout=10)
            campo.click()
            campo.send_keys(Keys.CONTROL + "a")
            campo.send_keys(Keys.BACKSPACE)
            time.sleep(0.12)

            # Digita caractere por caractere para disparar os eventos key* do portal
            for caractere in texto:
                campo.send_keys(caractere)
                time.sleep(delay_ms / 1000)

            # Verifica se o valor ficou estável antes de seguir
            esperado_digits = _somente_digitos(texto)
            for _ in range(20):
                try:
                    valor_atual = campo.get_attribute("value") or ""
                    if valor_atual.strip() == texto.strip():
                        return
                    if esperado_digits and _somente_digitos(valor_atual) == esperado_digits:
                        return
                except StaleElementReferenceException:
                    raise
                time.sleep(0.15)

            raise RuntimeError(f"O campo {seletor} não permaneceu com o valor esperado: {texto}")
        except StaleElementReferenceException:
            if tentativa == 2:
                raise
            time.sleep(0.5)


def _definir_valor_instantaneo(
    driver: WebDriver,
    by: str,
    seletor: str,
    texto: str,
) -> None:
    """Define o valor de um campo de texto instantaneamente executando JavaScript no navegador.

    Isso é utilizado para campos de textos muito longos, como a descrição dos serviços,
    evitando o gargalo de tempo da digitação caractere por caractere via Selenium.
    Dispara os eventos 'input' e 'change' para garantir que os ouvintes de eventos da página
    e o estado do JSF reconheçam a alteração do valor.
    """
    for tentativa in range(3):
        try:
            campo = _aguardar_elemento_clicavel(driver, by, seletor, timeout=10)
            driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                campo,
                texto,
            )
            return
        except StaleElementReferenceException:
            if tentativa == 2:
                raise
            time.sleep(0.5)


def _selecionar_opcao_por_texto(
    driver: WebDriver,
    by: str,
    seletor: str,
    textos: list[str],
    timeout: int = 10,
) -> None:
    """Seleciona a primeira opção disponível em um <select> a partir de uma lista de rótulos."""

    for tentativa in range(4):
        try:
            elemento = _aguardar_elemento_clicavel(driver, by, seletor, timeout=timeout)
            select = Select(elemento)

            # 1. Tenta correspondência exata normalizada (ignora caixa e acentos)
            for texto in textos:
                texto_norm = normalizar_texto(texto)
                for opcao in select.options:
                    opcao_texto_norm = normalizar_texto(opcao.text)
                    opcao_valor_norm = normalizar_texto(opcao.get_attribute("value") or "")
                    if (texto_norm == opcao_texto_norm) or (texto_norm == opcao_valor_norm):
                        select.select_by_value(opcao.get_attribute("value"))
                        return

            # 2. Tenta correspondência parcial normalizada (ex: "Fortaleza" contido em "Fortaleza - CE")
            for texto in textos:
                texto_norm = normalizar_texto(texto)
                for opcao in select.options:
                    opcao_texto_norm = normalizar_texto(opcao.text)
                    if texto_norm in opcao_texto_norm:
                        # Ignora placeholders de instruções como "selecione" ou "escolha"
                        if "selecione" in opcao_texto_norm or "escolha" in opcao_texto_norm:
                            continue
                        select.select_by_value(opcao.get_attribute("value"))
                        return

            raise NoSuchElementException(f"Nenhuma das opções {textos} foi localizada no select.")
            
        except (StaleElementReferenceException, NoSuchElementException) as exc:
            if tentativa == 3:
                raise RuntimeError(
                    f"Não foi possível selecionar nenhuma opção entre {textos} após várias tentativas."
                ) from exc
            time.sleep(0.5)


# ---------------------------------------------------------------------------
# Aguardando login manual
# ---------------------------------------------------------------------------


def aguardar_login_manual(driver: WebDriver) -> None:
    """Mostra a instrução e aguarda o usuário concluir o login manual no terminal."""

    print("O navegador foi aberto no portal da ISS de Fortaleza.")
    print("Faça o login manualmente e, quando terminar, volte aqui e pressione Enter.")
    input("Pressione Enter somente após o login estar concluído...")
    time.sleep(1)


def aguardar_login_manual_por_arquivo(driver: WebDriver, caminho_confirmacao: Path) -> None:
    """Espera até que o arquivo de confirmação seja criado pelo operador."""

    caminho_confirmacao.parent.mkdir(parents=True, exist_ok=True)
    if caminho_confirmacao.exists():
        caminho_confirmacao.unlink()

    print("O navegador foi aberto no portal da ISS de Fortaleza.")
    print("Faça o login manualmente e, ao terminar, crie o arquivo de confirmação para continuar.")
    print(f"Arquivo de confirmação esperado: {caminho_confirmacao.resolve()}")

    while not caminho_confirmacao.exists():
        time.sleep(1)

    # Evita que o próximo teste reutilize a mesma confirmação antiga
    try:
        caminho_confirmacao.unlink()
    except Exception:
        pass

    time.sleep(1)


# ---------------------------------------------------------------------------
# Seleção de competência via calendário RichFaces
# ---------------------------------------------------------------------------


def _valor_campo_competencia(driver: WebDriver, base_id: str) -> str:
    """Lê o valor do campo de data do calendário RichFaces."""

    try:
        campo = driver.find_element(By.ID, f"{base_id}InputDate")
        return campo.get_attribute("value") or ""
    except NoSuchElementException:
        return ""


def _abrir_editor_calendario(driver: WebDriver, base_id: str, rotulo: str) -> None:
    """Clica no botão de edição do calendário RichFaces para abrir o editor de mês/ano."""

    # O botão de edição fica no cabeçalho do calendário
    botao = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, f"#{base_id}Header .rich-calendar-tool-btn")
        )
    )
    botao.click()
    # Aguarda o botão OK do editor ficar visível como sinal de abertura do editor
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, f"{base_id}DateEditorButtonOk"))
    )


def _selecionar_mes_no_editor(driver: WebDriver, base_id: str, competencia: CompetenciaTrabalho) -> None:
    """Clica no mês desejado na grade do editor do calendário RichFaces."""

    indice_mes = competencia.mes - 1
    id_mes = f"{base_id}DateEditorLayoutM{indice_mes}"
    _clicar_por_id(driver, id_mes)


def _selecionar_ano_no_editor(driver: WebDriver, base_id: str, competencia: CompetenciaTrabalho) -> None:
    """Navega pelos anos do editor do calendário RichFaces e clica no ano desejado."""

    for _ in range(20):
        anos_visiveis: list[tuple[int, WebElement]] = []
        for indice in range(10):
            id_ano = f"{base_id}DateEditorLayoutY{indice}"
            try:
                el = driver.find_element(By.ID, id_ano)
                if not el.is_displayed():
                    continue
                texto = el.text.strip()
                if texto.isdigit():
                    anos_visiveis.append((int(texto), el))
            except NoSuchElementException:
                continue

        for ano_visivel, el_ano in anos_visiveis:
            if ano_visivel == competencia.ano:
                el_ano.click()
                return

        if not anos_visiveis:
            raise RuntimeError(
                f"Não consegui ler os anos visíveis no calendário de competência ({base_id})."
            )

        menor_ano = min(a for a, _ in anos_visiveis)
        maior_ano = max(a for a, _ in anos_visiveis)

        # Navega para esquerda ou direita conforme o ano-alvo
        if competencia.ano < menor_ano:
            sinal = "<"
        else:
            sinal = ">"

        try:
            botoes = driver.find_elements(By.CSS_SELECTOR, f"#{base_id} .rich-calendar-editor-btn")
            clicou = False
            for btn in botoes:
                if btn.text.strip() == sinal and btn.is_displayed():
                    btn.click()
                    clicou = True
                    break
            if not clicou:
                raise RuntimeError("Botão de navegação do calendário não encontrado.")
        except Exception as exc:
            raise RuntimeError(
                f"Não consegui navegar no calendário para chegar ao ano {competencia.ano}."
            ) from exc

        time.sleep(0.25)

    raise RuntimeError(
        f"Não consegui localizar o ano {competencia.ano} no calendário de competência ({base_id})."
    )


def _confirmar_editor_calendario(driver: WebDriver, base_id: str, esperado: str) -> None:
    """Clica em OK no editor do calendário e confirma que o campo assumiu o valor esperado."""

    _clicar_por_id(driver, f"{base_id}DateEditorButtonOk")
    # Aguarda o campo de data refletir o valor confirmado
    WebDriverWait(driver, 10).until(
        lambda d: (
            lambda el: el is not None and el.get_attribute("value") == esperado
        )(d.find_element(By.ID, f"{base_id}InputDate"))
    )


def _selecionar_campo_competencia(
    driver: WebDriver,
    base_id: str,
    rotulo: str,
    competencia: CompetenciaTrabalho,
) -> None:
    """Seleciona o mês e o ano no calendário RichFaces de um campo de competência."""

    esperado = f"{competencia.mes:02d}/{competencia.ano}"
    valor_atual = _valor_campo_competencia(driver, base_id)

    if valor_atual == esperado:
        registrar_evento_execucao(
            f"Campo {rotulo} da competência já estava em {esperado}.",
            "ISS Fortaleza",
        )
        return

    _abrir_editor_calendario(driver, base_id, rotulo)
    _selecionar_mes_no_editor(driver, base_id, competencia)
    _selecionar_ano_no_editor(driver, base_id, competencia)
    _confirmar_editor_calendario(driver, base_id, esperado)
    registrar_evento_execucao(
        f"Campo {rotulo} da competência ajustado para {esperado}.",
        "ISS Fortaleza",
    )


def selecionar_competencia_na_tela_richfaces(
    driver: WebDriver,
    competencia: CompetenciaTrabalho,
) -> None:
    """Seleciona a competência nos calendários `De` e `Até` usando o editor do RichFaces."""

    # Garante que a tela carregou antes de tentar acessar os calendários
    WebDriverWait(driver, 120).until(
        EC.element_to_be_clickable((By.ID, "manterEscrituracaoForm:btnConsultar"))
    )
    time.sleep(1.2)

    _selecionar_campo_competencia(driver, "manterEscrituracaoForm:dataInicial", "De", competencia)
    _selecionar_campo_competencia(driver, "manterEscrituracaoForm:dataFinal", "Até", competencia)
    registrar_evento_execucao(
        f"Competência selecionada na tela: {competencia.rotulo}",
        "ISS Fortaleza",
    )


# ---------------------------------------------------------------------------
# Botão Escriturar
# ---------------------------------------------------------------------------


def aguardar_botao_escriturar(driver: WebDriver, timeout: int = 120) -> bool:
    """Espera a consulta retornar e verifica se o botão de escriturar está habilitado."""

    id_ativo = "manterEscrituracaoForm:dataTable:0:linkEscriturar"
    id_desabilitado = "manterEscrituracaoForm:dataTable:0:linkEscriturarDesabilitado"

    prazo = time.monotonic() + timeout
    while time.monotonic() < prazo:
        try:
            el = driver.find_element(By.ID, id_ativo)
            if el.is_displayed():
                return True
        except NoSuchElementException:
            pass
        try:
            el = driver.find_element(By.ID, id_desabilitado)
            if el.is_displayed():
                return False
        except NoSuchElementException:
            pass
        time.sleep(0.5)

    raise CompetenciaSemEscriturarDisponivelError(
        "Não consegui identificar se a competência consultada liberou o botão de escriturar."
    )


# ---------------------------------------------------------------------------
# Tela de Escrituração Fiscal
# ---------------------------------------------------------------------------


def aguardar_tela_escrituracao_fiscal(driver: WebDriver, timeout: int = 120) -> None:
    """Espera a tela de Escrituração Fiscal ficar pronta após o clique em Escriturar."""

    prazo = time.monotonic() + timeout
    while time.monotonic() < prazo:
        try:
            el = driver.find_element(By.ID, "tabEncerramentoEscrituracao")
            if el.is_displayed():
                return
        except NoSuchElementException:
            pass
        try:
            el = driver.find_element(By.ID, "aba_tomados_lbl")
            if el.is_displayed():
                return
        except NoSuchElementException:
            pass
        time.sleep(0.25)

    raise RuntimeError(
        "A tela de Escrituração Fiscal não ficou visível após o clique em Escriturar."
    )


# ---------------------------------------------------------------------------
# Aba Serviços Tomados
# ---------------------------------------------------------------------------


def clicar_aba_servicos_tomados(driver: WebDriver) -> None:
    """Clica na aba 'Serviços Tomados' com estratégia específica para o portal RichFaces."""

    # Aguarda a tela de Escrituração Fiscal aparecer primeiro
    WebDriverWait(driver, 120).until(
        EC.visibility_of_element_located((By.ID, "tabEncerramentoEscrituracao"))
    )
    try:
        _clicar_por_id(driver, "aba_tomados_lbl", timeout=120)
        return
    except Exception:
        pass

    # Fallback: busca o elemento pelo onclick
    try:
        el = driver.find_element(By.CSS_SELECTOR, '[onclick*="atualizarDadosDaAbaTomados"]')
        el.click()
    except NoSuchElementException:
        raise RuntimeError("Não foi possível clicar na aba 'Serviços Tomados'.")


# ---------------------------------------------------------------------------
# Tela Digitar Documento
# ---------------------------------------------------------------------------


def aguardar_tela_digitar_documento(driver: WebDriver, timeout: int = 120) -> None:
    """Espera a tela de Digitar Documento carregar por completo."""

    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.ID, "digitarDocumentoForm"))
    )
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.ID, "digitarDocumentoForm:divPesquisaTomador"))
    )
    # Aguarda o rádio de CNPJ aparecer como confirmação final de carregamento
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located(
            (By.ID, "digitarDocumentoForm:tipoPesquisaTomadorRb:1")
        )
    )


# ---------------------------------------------------------------------------
# Seleção do prestador (CNPJ + autocomplete)
# ---------------------------------------------------------------------------


def selecionar_prestador(
    driver: WebDriver,
    cnpj: str,
    depuracao: bool = False,
) -> None:
    """Seleciona o prestador de forma visível, acionando o autocomplete do portal."""

    registrar_evento_execucao(
        f"Iniciando seleção do prestador para o CNPJ {cnpj}",
        "ISS Fortaleza",
    )
    aguardar_tela_digitar_documento(driver)

    # Ativa o modo de pesquisa por CNPJ
    _clicar_por_id(driver, "digitarDocumentoForm:tipoPesquisaTomadorRb:1")
    registrar_evento_execucao("Tipo de pesquisa alterado para CNPJ", "ISS Fortaleza")

    # Pausa para o portal recriar o campo de CNPJ após a alternância
    time.sleep(0.8)

    # Digita o CNPJ para disparar o autocomplete do RichFaces
    _digitar_campo(
        driver,
        By.ID,
        "digitarDocumentoForm:cpfPesquisaTomador",
        limpar_cnpj_para_digitacao(cnpj),
        delay_ms=80,
    )
    registrar_evento_execucao(
        f"CNPJ digitado na pesquisa do prestador: {cnpj}",
        "ISS Fortaleza",
    )
    _pausar_para_depuracao(
        driver,
        depuracao,
        f"O CNPJ {cnpj} foi digitado. Confira o autocomplete antes da seleção.",
    )

    # Aguarda o container do autocomplete aparecer (timeout tolerante)
    container_id = "digitarDocumentoForm:j_id189"
    texto_container = ""
    container = None
    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, container_id))
        )
        container = driver.find_element(By.ID, container_id)
        texto_container = container.text or ""
    except TimeoutException:
        pass

    # Verifica se há sugestões disponíveis na lista do autocomplete
    try:
        if container:
            sugestoes = container.find_elements(By.CSS_SELECTOR, "tr.richfaces_suggestionEntry")
        else:
            sugestoes = []

        if sugestoes and sugestoes[0].is_displayed():
            sugestoes[0].click()
            registrar_evento_execucao(
                f"Prestador selecionado na lista de sugestões para o CNPJ {cnpj}",
                "ISS Fortaleza",
            )
        else:
            if _texto_indica_prestador_nao_encontrado(texto_container):
                raise PrestadorNaoEncontradoError(cnpj, texto_container)
            # Fallback via teclado quando não há sugestão visível clicável
            campo = driver.find_element(By.ID, "digitarDocumentoForm:cpfPesquisaTomador")
            campo.send_keys(Keys.ARROW_DOWN)
            campo.send_keys(Keys.RETURN)
    except PrestadorNaoEncontradoError:
        raise
    except Exception:
        if _texto_indica_prestador_nao_encontrado(texto_container):
            raise PrestadorNaoEncontradoError(cnpj, texto_container)

    # Confirma que o campo do nome do prestador foi preenchido
    for _ in range(20):
        try:
            campo_nome = driver.find_element(By.ID, "digitarDocumentoForm:idNome")
            if campo_nome.get_attribute("value", ).strip():
                registrar_evento_execucao(
                    f"Nome do prestador preenchido com sucesso para o CNPJ {cnpj}",
                    "ISS Fortaleza",
                )
                return
        except (NoSuchElementException, StaleElementReferenceException):
            pass
        time.sleep(0.25)

    if _texto_indica_prestador_nao_encontrado(texto_container):
        raise PrestadorNaoEncontradoError(cnpj, texto_container)

    detalhe = texto_container or "O campo de nome continuou vazio após a busca do CNPJ."
    registrar_evento_execucao(
        f"Portal não retornou um prestador selecionável para o CNPJ {cnpj}",
        "ISS Fortaleza",
    )
    raise PrestadorNaoEncontradoError(cnpj, detalhe)


# ---------------------------------------------------------------------------
# Modal de pesquisa de CNAE
# ---------------------------------------------------------------------------


def _forcar_abertura_modal_cnae(driver: WebDriver) -> None:
    """Chama a API nativa do RichFaces via JavaScript para forçar a abertura do modal."""

    try:
        driver.execute_script("Richfaces.showModalPanel('pesquisarCnaeModal')")
    except JavascriptException:
        # Fallback: chama a função com prefixo alternativo que alguns portais usam
        driver.execute_script("RichFaces.showModalPanel('pesquisarCnaeModal')")


def preencher_modal_pesquisar_cnae(
    driver: WebDriver,
    id_cnae_final: str,
    depuracao: bool = False,
) -> None:
    """Preenche o modal de pesquisa de CNAE com o código final da planilha e dispara a busca.

    Caso o modal não apareça dentro do tempo padrão após o clique, utiliza a API
    nativa do RichFaces via JavaScript para forçá-lo a abrir — estratégia de
    fallback que corrige o timeout observado quando o JSF ainda está reprocessando.
    """

    valor_cnae = limpar_numero_para_digitacao(id_cnae_final)
    registrar_evento_execucao(
        f"Abrindo modal de pesquisa de CNAE para o código {valor_cnae}",
        "ISS Fortaleza",
    )

    # Tenta aguardar o modal ficar visível normalmente
    try:
        WebDriverWait(driver, 8).until(
            EC.visibility_of_element_located(
                (By.ID, "digitarDocumentoForm:pesquisarCnaeModalContainer")
            )
        )
    except TimeoutException:
        # Fallback via JavaScript: força o RichFaces a exibir o modal
        registrar_evento_execucao(
            "Modal de CNAE não abriu sozinho; forçando via JavaScript do RichFaces.",
            "ISS Fortaleza",
        )
        _forcar_abertura_modal_cnae(driver)
        # Aguarda mais tempo após a chamada forçada
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.ID, "digitarDocumentoForm:pesquisarCnaeModalContainer")
            )
        )

    # Preenche o campo de pesquisa dentro do modal
    _digitar_campo(
        driver,
        By.ID,
        "digitarDocumentoForm:idFormularioPesquisaCnae:idCnaePesquisa",
        valor_cnae,
        delay_ms=45,
    )
    _pausar_para_depuracao(
        driver,
        depuracao,
        f"O modal de pesquisa de CNAE está preenchido com {valor_cnae}.",
    )

    # Dispara a pesquisa AJAX do CNAE escolhido
    _clicar_por_id(driver, "digitarDocumentoForm:idFormularioPesquisaCnae:idPesquisar")
    time.sleep(1.2)
    registrar_evento_execucao(
        f"Pesquisa de CNAE enviada para {valor_cnae}",
        "ISS Fortaleza",
    )
    _pausar_para_depuracao(
        driver,
        depuracao,
        f"A pesquisa de CNAE para {valor_cnae} foi enviada. Confirme o resultado na lista.",
    )

    # Seleciona a primeira linha de resultado retornada pelo portal
    # O ID do link de seleção usa um índice dinâmico; buscamos com seletor parcial
    seletor_resultado = (
        "[id^='digitarDocumentoForm:idFormularioPesquisaCnae:idDatatableListaCnae:0:']"
    )
    el_resultado = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_resultado))
    )
    el_resultado.click()
    time.sleep(1.2)
    registrar_evento_execucao(
        f"CNAE selecionado na lista de resultados: {valor_cnae}",
        "ISS Fortaleza",
    )


def _digitar_campo_se_nao_zero(
    driver: WebDriver,
    seletor_id: str,
    valor: str,
) -> None:
    """Preenche o campo apenas se o valor for numérico e diferente de zero."""
    if not valor:
        return

    # Remove tudo que não for dígito para checar se é zero
    digitos = re.sub(r"\D+", "", valor)
    if not digitos or int(digitos) == 0:
        return

    _digitar_campo(
        driver,
        By.ID,
        seletor_id,
        limpar_valor_para_digitacao(valor),
    )


# ---------------------------------------------------------------------------
# Preenchimento da aba Serviço
# ---------------------------------------------------------------------------


def preencher_documento_servico(
    driver: WebDriver,
    documento: DocumentoPortalISS,
    depuracao: bool = False,
) -> None:
    """Preenche a aba de serviço com os dados de uma linha da planilha."""

    # Ativa a aba Serviço real do portal antes de preencher os campos
    _clicar_por_id(driver, "digitarDocumentoForm:abaServico_lbl")
    registrar_evento_execucao("Aba Serviço acionada", "ISS Fortaleza")
    time.sleep(0.5)

    # Aguarda os campos da aba Serviço ficarem disponíveis
    WebDriverWait(driver, 120).until(
        EC.visibility_of_element_located((By.ID, "digitarDocumentoForm"))
    )
    WebDriverWait(driver, 120).until(
        EC.visibility_of_element_located((By.ID, "digitarDocumentoForm:tipoDocumentoDigitado"))
    )
    WebDriverWait(driver, 120).until(
        EC.visibility_of_element_located((By.ID, "digitarDocumentoForm:statusNfse"))
    )
    registrar_evento_execucao(
        "Formulário da aba Serviço ficou visível e pronto para preenchimento",
        "ISS Fortaleza",
    )

    registrar_evento_execucao(
        f"Iniciando preenchimento da aba Serviço para a NF {documento.numero_nf} ({documento.cnpj_prestador})",
        "ISS Fortaleza",
    )

    # Tipo de documento
    _selecionar_opcao_por_texto(
        driver,
        By.ID,
        "digitarDocumentoForm:tipoDocumentoDigitado",
        ["NFS-e de Outro Município", "NFS-e de outro município", "707"],
    )

    # Número da nota fiscal
    _digitar_campo(
        driver,
        By.ID,
        "digitarDocumentoForm:numeroDocumentoDigitado",
        limpar_numero_para_digitacao(documento.numero_nf),
    )

    # Data de emissão
    _digitar_campo(
        driver,
        By.ID,
        "digitarDocumentoForm:dataEmissaoInputDate",
        documento.data_emissao,
    )

    # Status NFSE
    _selecionar_opcao_por_texto(
        driver,
        By.ID,
        "digitarDocumentoForm:statusNfse",
        ["NORMAL", "Normal", "472"],
    )
    registrar_evento_execucao(
        f"Campos iniciais da aba Serviço preenchidos para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )

    # Pausa de estabilização para o JSF reprocessar após selecionar o Status
    time.sleep(1.2)

    # Abre o modal de pesquisa de CNAE
    _clicar_por_id(driver, "digitarDocumentoForm:idLinkPesquisarCnae")
    registrar_evento_execucao(
        f"Botão Pesquisar CNAE acionado para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )

    # Preenche o modal de CNAE (com fallback via JavaScript)
    preencher_modal_pesquisar_cnae(driver, documento.id_cnae_final, depuracao=depuracao)

    # Descrição do serviço preenchida de forma instantânea via JS
    _definir_valor_instantaneo(
        driver,
        By.ID,
        "digitarDocumentoForm:idDescricaoServico",
        documento.descricao_servico,
    )

    # UF de prestação
    _selecionar_opcao_por_texto(
        driver,
        By.ID,
        "digitarDocumentoForm:comboEscolherEstadoLocalPrestacao",
        [documento.uf_local_prestacao],
    )

    # Cidade de prestação
    _selecionar_opcao_por_texto(
        driver,
        By.ID,
        "digitarDocumentoForm:comboEscolherCidadeLocalPrestacao",
        [documento.cidade_local_prestacao],
    )

    # Natureza da operação
    _selecionar_opcao_por_texto(
        driver,
        By.ID,
        "digitarDocumentoForm:comboEscolherLocalPrestacao",
        [documento.natureza_operacao],
    )

    # Checkbox ISS retido
    deve_marcar = normalizar_texto(documento.iss_retido).startswith("sim")
    try:
        checkbox_iss = driver.find_element(By.NAME, "digitarDocumentoForm:j_id361")
        marcado_agora = checkbox_iss.is_selected()
        if deve_marcar and not marcado_agora:
            checkbox_iss.click()
        elif not deve_marcar and marcado_agora:
            checkbox_iss.click()
    except NoSuchElementException:
        pass
    registrar_evento_execucao(
        f"Local de prestação e ISS tratados para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )

    # Valor do serviço e retenções/deduções (somente preenche se não for zero)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idValorServicoPrestado", documento.valor_servico)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idValorDeducoes", documento.valor_deducoes)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idDescontosIncondicionados", documento.descontos_incondicionados)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idDescontosCondicionados", documento.descontos_condicionados)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idOutrasRetencoes", documento.outras_retencoes)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idIR", documento.ir)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idPis", documento.pis_nao_retido)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idConfins", documento.cofins_nao_retido)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idCSLL", documento.csrf)
    _digitar_campo_se_nao_zero(driver, "digitarDocumentoForm:idINSS", documento.inss)

    registrar_evento_execucao(
        f"Aba Serviço preenchida com sucesso para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )


# ---------------------------------------------------------------------------
# Fluxo principal da ISS
# ---------------------------------------------------------------------------


def executar_fluxo_iss(
    documentos: list[DocumentoPortalISS],
    competencia: CompetenciaTrabalho,
    caminho_log: Path,
    depuracao: bool = False,
    aguardar_login_por_arquivo: bool = False,
    caminho_confirmacao_login: Path | None = None,
    reutilizar_navegador: bool = False,
    perfil_navegador: Path = PERFIL_CHROME_FUNCAO2_PADRAO,
    callback_progresso: Any = None,
    executando_em_gui: bool = False,
) -> None:
    """Executa o fluxo visual completo do portal da ISS até o preenchimento do formulário."""

    if not documentos:
        raise RuntimeError("A planilha não contém linhas válidas para a automação.")

    # Abre o Chrome com ou sem perfil persistente, conforme solicitado
    if reutilizar_navegador:
        driver = abrir_navegador_com_perfil_persistente(perfil_navegador, depuracao=depuracao)
    else:
        driver = abrir_navegador_visivel(depuracao=depuracao)

    try:
        registrar_evento_execucao(
            f"Fluxo da ISS iniciado com {len(documentos)} documento(s) e competência {competencia.rotulo}",
            "ISS Fortaleza",
        )

        # Aguarda login manual do operador
        if aguardar_login_por_arquivo:
            caminho = caminho_confirmacao_login or Path("brain/funcao2_login_ok.flag")
            aguardar_login_manual_por_arquivo(driver, caminho)
        else:
            aguardar_login_manual(driver)
        registrar_evento_execucao("Login manual confirmado pelo usuário", "ISS Fortaleza")

        print("Login confirmado. Aguardando o menu do portal ficar disponível...")
        try:
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Escrituração') or contains(text(), 'Escrituracao')]")
                )
            )
        except TimeoutException:
            registrar_evento_execucao(
                "Timeout ao aguardar menu Escrituração visível", "ISS Fortaleza"
            )
            print("AVISO: Timeout ao aguardar menu Escrituração. Prosseguindo mesmo assim...")
        _pausar_para_depuracao(
            driver,
            depuracao,
            "O portal foi carregado e o login já foi confirmado.",
        )

        # Navega para Escrituração > Manter Escrituração
        print("Acessando Escrituração > Manter Escrituração...")
        menus = driver.find_elements(By.CSS_SELECTOR, "a.dropdown-toggle")
        if len(menus) > 4:
            menus[4].click()
            time.sleep(0.5)
        _clicar_por_id(driver, "formMenuTopo:menuEscrituracao:j_id80")
        registrar_evento_execucao("Menu Escrituração acionado", "ISS Fortaleza")

        WebDriverWait(driver, 120).until(
            EC.element_to_be_clickable((By.ID, "manterEscrituracaoForm:btnConsultar"))
        )
        registrar_evento_execucao("Tela Manter Escrituração aberta", "ISS Fortaleza")

        # Seleciona a competência nos calendários De e Até
        print(
            f"Selecionando a competência {competencia.nome_mes} {competencia.ano} na tela de manutenção..."
        )
        selecionar_competencia_na_tela_richfaces(driver, competencia)

        # Aciona a consulta da competência
        print("Consultando a competência selecionada...")
        _clicar_por_id(driver, "manterEscrituracaoForm:btnConsultar")
        registrar_evento_execucao("Botão Consultar acionado", "ISS Fortaleza")

        # Loop para lidar com competências sem botão Escriturar habilitado
        while True:
            botao_escriturar_habilitado = aguardar_botao_escriturar(driver)
            if botao_escriturar_habilitado:
                break

            mensagem_sem_escriturar = (
                f"A competência {competencia.rotulo} está desabilitada para escriturar. "
                "Escolha outro mês para continuar a automação."
            )
            print(mensagem_sem_escriturar)
            registrar_evento_execucao(mensagem_sem_escriturar, "ISS Fortaleza")

            if not sys.stdin.isatty():
                raise CompetenciaSemEscriturarDisponivelError(mensagem_sem_escriturar)

            competencia = solicitar_nova_competencia_para_repetir(competencia)
            print(
                f"Reiniciando a seleção da competência {competencia.nome_mes} {competencia.ano} "
                "a partir dos campos De e Até..."
            )
            registrar_evento_execucao(
                f"Nova competência informada pelo usuário para retry: {competencia.rotulo}",
                "ISS Fortaleza",
            )
            selecionar_competencia_na_tela_richfaces(driver, competencia)
            _clicar_por_id(driver, "manterEscrituracaoForm:btnConsultar")
            registrar_evento_execucao("Botão Consultar acionado", "ISS Fortaleza")

        # Abre o formulário de escrituração
        print("Abrindo a rotina de escrituração...")
        _clicar_por_id(driver, "manterEscrituracaoForm:dataTable:0:linkEscriturar")
        registrar_evento_execucao("Botão Escriturar acionado", "ISS Fortaleza")
        aguardar_tela_escrituracao_fiscal(driver)
        registrar_evento_execucao("Tela Escrituração Fiscal aberta", "ISS Fortaleza")

        # Clica na aba Serviços Tomados
        print("Selecionando a aba Serviços Tomados...")
        clicar_aba_servicos_tomados(driver)
        registrar_evento_execucao("Aba Serviços Tomados acionada", "ISS Fortaleza")

        # Aguarda e clica no botão Digitar Documento
        WebDriverWait(driver, 120).until(
            EC.visibility_of_element_located((By.ID, "servico_tomado_form:seamj_id849"))
        )
        print("Abrindo Digitar Documento...")
        _clicar_por_id(driver, "servico_tomado_form:seamj_id849")
        registrar_evento_execucao("Tela Digitar Documento aberta", "ISS Fortaleza")
        aguardar_tela_digitar_documento(driver)
        time.sleep(0.5)
        _pausar_para_depuracao(
            driver,
            depuracao,
            "A tela Digitar Documento está aberta e pronta para o próximo clique.",
        )

        print(
            f"A planilha possui {len(documentos)} linha(s) válida(s). O fluxo iniciará o processamento em lote."
        )

        pasta_log = caminho_log.parent
        if pasta_log.name != "log":
            pasta_log = pasta_log / "log"
        pasta_log.mkdir(parents=True, exist_ok=True)
        
        # Limpa arquivos de logs anteriores para esta execução
        for nome in ["log_escrituradas_sucesso.txt", "log_nao_cadastrados.txt", "log_prefeitura_fortaleza.txt"]:
            caminho_antigo = pasta_log / nome
            if caminho_antigo.exists():
                try:
                    caminho_antigo.unlink()
                except Exception:
                    pass

        configurar_pasta_logs(pasta_log)
        notas_sucesso: list[str] = []
        notas_nao_cadastradas: list[str] = []

        for indice, candidato in enumerate(documentos, start=1):
            if callback_progresso:
                try:
                    callback_progresso(indice, len(documentos))
                except Exception:
                    pass
            prefeitura_normalizada = normalizar_texto(candidato.prefeitura).upper()
            if prefeitura_normalizada == "PREFEITURA MUNICIPAL DE FORTALEZA":
                mensagem_log = (
                    f"ARQUIVO_PDF={candidato.arquivo_pdf} | "
                    f"CNPJ_PRESTADOR={candidato.cnpj_prestador} | "
                    f"IGNORADO: Nota emitida pela Prefeitura de Fortaleza"
                )
                registrar_log_funcao2(caminho_log, mensagem_log)
                print(f"Linha {indice}/{len(documentos)} ignorada: Nota da Prefeitura de Fortaleza.")
                
                # Grava no log exclusivo de Fortaleza em tempo real
                caminho_fortaleza = pasta_log / "log_prefeitura_fortaleza.txt"
                try:
                    if not caminho_fortaleza.exists():
                        with open(caminho_fortaleza, "w", encoding="utf-8") as f:
                            f.write("As seguintes notas fiscais foram ignoradas por terem sido emitidas pela Prefeitura Municipal de Fortaleza:\n\n")
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    with open(caminho_fortaleza, "a", encoding="utf-8") as f:
                        f.write(f"[{timestamp}] {candidato.arquivo_pdf} - CNPJ: {candidato.cnpj_prestador} - NF: {candidato.numero_nf}\n")
                except Exception as exc_fort:
                    print(f"Erro ao registrar log de Fortaleza: {exc_fort}")
                continue

            print(
                f"Processando linha {indice}/{len(documentos)}: "
                f"{candidato.cnpj_prestador} - NF {candidato.numero_nf}"
            )
            try:
                selecionar_prestador(driver, candidato.cnpj_prestador, depuracao=depuracao)
            except PrestadorNaoEncontradoError as exc:
                mensagem_log = (
                    f"ARQUIVO_PDF={candidato.arquivo_pdf} | "
                    f"CNPJ_PRESTADOR={candidato.cnpj_prestador} | "
                    f"{exc}"
                )
                registrar_log_funcao2(caminho_log, mensagem_log)
                nf_info = f"{candidato.arquivo_pdf} - {candidato.cnpj_prestador} - {candidato.numero_nf}"
                notas_nao_cadastradas.append(nf_info)
                print(f"Prestador não encontrado para {candidato.cnpj_prestador}; avançando.")
                
                # Grava no log de prestadores não cadastrados em tempo real
                caminho_nao_cadastrados = pasta_log / "log_nao_cadastrados.txt"
                try:
                    if not caminho_nao_cadastrados.exists():
                        with open(caminho_nao_cadastrados, "w", encoding="utf-8") as f:
                            f.write("As seguintes notas fiscais não puderam ser escrituradas pois os prestadores correspondentes não estavam cadastrados no portal da ISS Fortaleza:\n\n")
                    with open(caminho_nao_cadastrados, "a", encoding="utf-8") as f:
                        f.write(f"- {nf_info}\n")
                except Exception as exc_nc:
                    print(f"Erro ao salvar log de não cadastrados em tempo real: {exc_nc}")
                continue

            try:
                preencher_documento_servico(driver, candidato, depuracao=depuracao)
                
                # Gravar documento
                print("Gravando documento no portal...")
                registrar_evento_execucao(f"Gravando NF {candidato.numero_nf}", "ISS Fortaleza")
                try:
                    _clicar_por_id(driver, "digitarDocumentoForm:j_id477", timeout=15)
                except Exception:
                    el = driver.find_element(By.XPATH, "//*[@id='digitarDocumentoForm:j_id477'] | //input[@value='Gravar' or @value='Gravar Documento']")
                    driver.execute_script("arguments[0].click();", el)

                # Aguarda feedback do portal
                WebDriverWait(driver, 15).until(
                    EC.text_to_be_present_in_element((By.XPATH, '//*[@id="content"]/legend/h2'), "Documento digitado com Sucesso")
                )
                registrar_evento_execucao(f"Sucesso na gravação da NF {candidato.numero_nf}", "ISS Fortaleza")
                nf_info = f"{candidato.arquivo_pdf} - {candidato.cnpj_prestador} - {candidato.numero_nf}"
                notas_sucesso.append(nf_info)
                
                # Grava no log de sucesso em tempo real
                caminho_sucesso = pasta_log / "log_escrituradas_sucesso.txt"
                try:
                    if not caminho_sucesso.exists():
                        with open(caminho_sucesso, "w", encoding="utf-8") as f:
                            f.write("As seguintes notas fiscais foram escrituradas e gravadas com sucesso no portal:\n\n")
                    with open(caminho_sucesso, "a", encoding="utf-8") as f:
                        f.write(f"- {nf_info}\n")
                except Exception as exc_suc:
                    print(f"Erro ao salvar log de sucesso em tempo real: {exc_suc}")

            except Exception as exc:
                registrar_evento_execucao(f"Falha ao processar NF {candidato.numero_nf}: {exc}", "ISS Fortaleza")
                print(f"Erro ao processar NF {candidato.numero_nf}: {exc}")
                registrar_log_funcao2(caminho_log, f"ERRO ao processar {candidato.arquivo_pdf}: {exc}")

            finally:
                # Clica em Digitar novo documento para limpar o formulário para a próxima linha
                try:
                    _clicar_com_espera(driver, By.XPATH, '//*[@id="j_id165:novo"]', "botão Novo Documento", timeout=15)
                    time.sleep(2)
                except Exception as exc:
                    print(f"Não foi possível clicar em 'Novo documento': {exc}. O fluxo pode falhar na próxima iteração.")

        print()
        print("Gravação de documentos em lote concluída.")
        if not executando_em_gui:
            print("Revise a tela no navegador e, se quiser encerrar, volte ao terminal.")
            input("Pressione Enter para encerrar esta sessão automatizada e manter a tela aberta...")
        else:
            print("Execução da GUI concluída. O navegador permanecerá aberto.")

    finally:
        # Mantém o navegador aberto em modo de sessão persistente; fecha nos demais casos
        if not reutilizar_navegador:
            try:
                driver.quit()
            except Exception:
                pass
        else:
            registrar_evento_execucao(
                "Sessão persistente da função 2 preservada para reutilização em nova execução",
                "ISS Fortaleza",
            )


# ---------------------------------------------------------------------------
# Ponto de entrada síncrono exposto para a CLI
# ---------------------------------------------------------------------------


def executar_automacao_iss(
    caminho_xlsx: Path,
    competencia: CompetenciaTrabalho,
    depuracao: bool = False,
    usar_inspector: bool = False,
    aguardar_login_por_arquivo: bool = False,
    caminho_confirmacao_login: Path | None = None,
    reutilizar_navegador: bool = False,
    perfil_navegador: Path = PERFIL_CHROME_FUNCAO2_PADRAO,
    porta_debug_navegador: int = PORTA_DEBUG_CHROME_PADRAO,
    callback_progresso: Any = None,
    executando_em_gui: bool = False,
) -> None:
    """Ponto de entrada síncrono para a opção 2 da CLI.

    O parâmetro `usar_inspector` não é mais aplicável com Selenium e é mantido
    apenas para compatibilidade de assinatura com o código da CLI.
    O parâmetro `porta_debug_navegador` foi removido do fluxo ativo pois o
    Selenium gerencia o driver de forma nativa; é mantido para compatibilidade.
    """

    registrar_evento_execucao(
        f"Automação da ISS solicitada com planilha {caminho_xlsx.resolve()} e competência {competencia.rotulo}",
        "ISS Fortaleza",
    )
    documentos = carregar_documentos_xlsx(caminho_xlsx)
    caminho_log = caminho_xlsx.parent / "log" / f"{caminho_xlsx.stem}_log_funcao2.txt"
    executar_fluxo_iss(
        documentos,
        competencia,
        caminho_log,
        depuracao=depuracao,
        aguardar_login_por_arquivo=aguardar_login_por_arquivo,
        caminho_confirmacao_login=caminho_confirmacao_login,
        reutilizar_navegador=reutilizar_navegador,
        perfil_navegador=perfil_navegador,
        callback_progresso=callback_progresso,
        executando_em_gui=executando_em_gui,
    )


# ---------------------------------------------------------------------------
# Extração avulsa do portal (funcionalidade auxiliar mantida)
# ---------------------------------------------------------------------------


def _limpar_texto_exibido(texto: str | None) -> str:
    """Normaliza espaços para deixar a extração mais legível no XLSX."""

    return " ".join((texto or "").split()).strip()


def extrair_dados_da_pagina_iss(driver: WebDriver) -> dict[str, object]:
    """Coleta o conteúdo visível da tela logada para exportação e análise posterior."""

    titulo = driver.title
    url = driver.current_url

    try:
        texto_bruto = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        texto_bruto = ""

    dados: dict[str, object] = {
        "titulo": titulo,
        "url": url,
        "extraido_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "texto_visivel": [linha.strip() for linha in texto_bruto.splitlines() if linha.strip()],
        "elementos": _coletar_elementos_visiveis(driver),
    }
    return dados


def _coletar_elementos_visiveis(driver: WebDriver) -> list[dict[str, str]]:
    """Captura campos e botões visíveis para documentar a tela atual do portal."""

    elementos: list[dict[str, str]] = []
    for tag in ("input", "select", "textarea", "button"):
        for el in driver.find_elements(By.TAG_NAME, tag):
            try:
                if not el.is_displayed():
                    continue
                elementos.append({
                    "tag": tag,
                    "id": el.get_attribute("id") or "",
                    "name": el.get_attribute("name") or "",
                    "type": el.get_attribute("type") or "",
                    "value": el.get_attribute("value") or "",
                    "texto": el.text or "",
                })
            except StaleElementReferenceException:
                continue
    return elementos


def _aplicar_cabecalho(aba) -> None:
    """Padroniza o estilo visual das planilhas geradas pela extração."""

    for celula in aba[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")


def ajustar_largura_colunas(aba) -> None:
    """Ajusta a largura das colunas da aba para facilitar a leitura."""

    for coluna in aba.columns:
        largura_max = max(
            (len(str(celula.value)) if celula.value else 0) for celula in coluna
        )
        aba.column_dimensions[coluna[0].column_letter].width = min(largura_max + 4, 80)


def exportar_extracao_iss(dados: dict[str, object], destino: Path) -> None:
    """Exporta a página extraída para XLSX."""

    wb = Workbook()

    aba_resumo = wb.active
    aba_resumo.title = "Resumo"
    aba_resumo.append(["Campo", "Valor"])
    aba_resumo.append(["Título", dados.get("titulo", "")])
    aba_resumo.append(["URL", dados.get("url", "")])
    aba_resumo.append(["Extraído em", dados.get("extraido_em", "")])
    aba_resumo.append(["Linhas de texto", len(dados.get("texto_visivel", []))])
    aba_resumo.append(["Elementos", len(dados.get("elementos", []))])
    _aplicar_cabecalho(aba_resumo)

    aba_texto = wb.create_sheet("Texto")
    aba_texto.append(["Linha", "Conteúdo"])
    for indice, linha in enumerate(dados.get("texto_visivel", []), start=1):
        aba_texto.append([indice, linha])
    _aplicar_cabecalho(aba_texto)

    aba_elementos = wb.create_sheet("Elementos")
    aba_elementos.append(["Tag", "ID", "Name", "Type", "Value", "Texto"])
    for el in dados.get("elementos", []):
        aba_elementos.append([
            el.get("tag", ""),
            el.get("id", ""),
            el.get("name", ""),
            el.get("type", ""),
            el.get("value", ""),
            el.get("texto", ""),
        ])
    _aplicar_cabecalho(aba_elementos)

    for aba in wb.worksheets:
        ajustar_largura_colunas(aba)

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)


def executar_extracao_portal_iss(destino_xlsx: Path = Path("iss_extracao.xlsx")) -> None:
    """Ponto de entrada síncrono para extrair o conteúdo visível do portal logado."""

    registrar_evento_execucao(
        f"Extração avulsa do portal ISS iniciada com destino {destino_xlsx.resolve()}",
        "ISS Fortaleza",
    )
    driver = abrir_navegador_visivel()
    try:
        aguardar_login_manual(driver)

        print()
        print("Quando estiver na tela exata da qual deseja extrair os dados, volte ao terminal.")
        input("Pressione Enter para capturar a página atual...")

        time.sleep(1)
        dados = extrair_dados_da_pagina_iss(driver)
        exportar_extracao_iss(dados, destino_xlsx)
        registrar_evento_execucao(
            f"Extração avulsa concluída com sucesso em {destino_xlsx.resolve()}",
            "ISS Fortaleza",
        )

        print(f"Extração concluída com sucesso: {destino_xlsx.resolve()}")
        input("Pressione Enter para encerrar esta sessão automatizada e fechar o navegador...")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
