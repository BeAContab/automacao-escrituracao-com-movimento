"""Módulo de automação para exportação em lote de XMLs de Serviços Prestados do portal da ISS Fortaleza."""

from __future__ import annotations

import time
from pathlib import Path

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from iss_fortaleza_automacao import (
    abrir_navegador_visivel,
    interpretar_competencia,
    CompetenciaTrabalho,
    _selecionar_campo_competencia,
)
from tratamento_erros import registrar_evento_execucao

# Seletores XPath para a tela de Consulta de NFS-e
XPATH_LINK_CONSULTAR_NFSE    = "//*[@id='homeForm:divHotLinks']/div[4]/a/h4"
XPATH_ABA_COMPETENCIA        = "//*[@id='consultarnfseForm:competencia_prestador_tab_lbl']"
XPATH_BTN_CONSULTAR          = "//*[@id='consultarnfseForm:j_id237']"
XPATH_BTN_SELECIONAR_PAGINA  = "//*[@id='consultarnfseForm:j_id324']"
XPATH_BTN_EXPORTAR_XML       = "//*[@id='consultarnfseForm:j_id321']/div[1]/input[3]"

# Base ID do calendário de competência na tela de Consulta
BASE_ID_COMPETENCIA = "consultarnfseForm:competenciaHeader"

# Número máximo de páginas por lote antes de exportar
MAX_PAGINAS_POR_LOTE = 10


def _aguardar_downloads_concluirem(pasta: Path, timeout: int = 120) -> None:
    """Aguarda até que não haja mais arquivos temporários de download do Chrome (.crdownload) na pasta."""
    prazo = time.monotonic() + timeout
    while time.monotonic() < prazo:
        temporarios = list(pasta.glob("*.crdownload"))
        if not temporarios:
            return
        time.sleep(1.5)
    raise TimeoutError(f"Download não concluiu em {timeout} segundos na pasta: {pasta}")


def _tentar_avancar_pagina(driver) -> bool:
    """Tenta clicar no botão '>' de avançar de página. Retorna False se a última página foi atingida."""
    try:
        botao = driver.find_element(By.XPATH, "//td[contains(@class,'rich-datascr-button')]/a[normalize-space(text())='>']")
        if botao.is_displayed() and botao.is_enabled():
            botao.click()
            time.sleep(2.5)
            return True
    except NoSuchElementException:
        pass
    return False


def _selecionar_competencia_consulta(driver, competencia: CompetenciaTrabalho) -> None:
    """Seleciona o mês/ano no campo de competência da tela de Consulta de NFS-e."""
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, f"//*[@id='{BASE_ID_COMPETENCIA}']"))
    )
    time.sleep(1.0)
    _selecionar_campo_competencia(driver, BASE_ID_COMPETENCIA, "Competência", competencia)


def executar_exportacao_xml_prestados(
    pasta_destino: str,
    competencia_str: str,
    callback_log=None,
    callback_progresso=None,
    caminho_confirmacao_login: Path | None = None,
) -> None:
    """
    Executa a automação de exportação de XMLs de Serviços Prestados do portal da ISS Fortaleza.

    Fluxo:
    1. Abre o Chrome com pasta de download configurada para pasta_destino.
    2. Aguarda o login manual do operador via arquivo de flag.
    3. Navega até a tela de Consulta de NFS-e, aba Competência/Tomador.
    4. Seleciona a competência informada pelo usuário na GUI.
    5. Consulta e itera pelas páginas em lotes de até 10, exportando o XML a cada lote.
    """

    def log(msg: str, is_error: bool = False) -> None:
        if callback_log:
            callback_log(msg, is_error)
        else:
            print(f"{'[ERRO] ' if is_error else ''}{msg}", flush=True)

    def progresso(pct: int, status: str) -> None:
        if callback_progresso:
            callback_progresso(pct, status)

    competencia = interpretar_competencia(competencia_str)
    pasta = Path(pasta_destino)

    log(f"Iniciando exportação de XMLs para competência {competencia.mes:02d}/{competencia.ano}...")
    log(f"Pasta de destino dos downloads: {pasta}")
    progresso(5, "Abrindo navegador...")

    driver = abrir_navegador_visivel(pasta_downloads=pasta)

    try:
        # Aguarda login manual do operador por arquivo de flag
        if caminho_confirmacao_login:
            log("Aguardando login manual no portal. Confirme na interface após fazer o login...")
            prazo_login = time.monotonic() + 300
            while not caminho_confirmacao_login.exists():
                if time.monotonic() > prazo_login:
                    raise TimeoutError("Tempo limite de login (5 min) esgotado.")
                time.sleep(1.0)
            log("Login confirmado pelo operador. Iniciando navegação...")

        progresso(10, "Acessando tela de Consulta de NFS-e...")

        # Passo 1: Clicar em "Consultar NFS-e"
        log("Clicando em 'Consultar NFS-e'...")
        WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_LINK_CONSULTAR_NFSE))
        ).click()
        time.sleep(2.0)

        # Passo 2: Clicar na aba "Competência/Tomador"
        log("Selecionando aba 'Competência/Tomador'...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_ABA_COMPETENCIA))
        ).click()
        time.sleep(1.5)

        # Passo 3: Selecionar mês/ano no calendário
        log(f"Selecionando competência {competencia.mes:02d}/{competencia.ano}...")
        _selecionar_competencia_consulta(driver, competencia)
        progresso(20, "Competência selecionada. Consultando notas...")

        # Passo 4: Clicar em "Consultar"
        log("Clicando em 'Consultar'...")
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_BTN_CONSULTAR))
        ).click()
        time.sleep(3.0)

        progresso(25, "Consultando resultados...")

        # Passo 5: Loop de varredura de páginas com exportação em lote
        pagina_atual = 1
        lote_atual = 0
        total_exportados = 0

        while True:
            log(f"Selecionando todas as notas da página {pagina_atual}...")

            try:
                btn_selecionar = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_BTN_SELECIONAR_PAGINA))
                )
                btn_selecionar.click()
                time.sleep(1.5)
            except TimeoutException:
                log("Botão de seleção de página não encontrado. Encerrando loop.", is_error=True)
                break

            lote_atual += 1
            pagina_atual += 1

            # Tenta avançar para a próxima página antes de decidir se exporta
            ultima_pagina = not _tentar_avancar_pagina(driver)

            # Exporta se atingiu o lote máximo ou chegou na última página
            if lote_atual >= MAX_PAGINAS_POR_LOTE or ultima_pagina:
                log(f"Exportando lote de {lote_atual} página(s)...")
                progresso(min(30 + (total_exportados * 5), 90), f"Exportando lote ({total_exportados + 1})...")

                try:
                    btn_exportar = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_BTN_EXPORTAR_XML))
                    )
                    btn_exportar.click()
                    time.sleep(2.0)

                    log("Aguardando conclusão do download...")
                    _aguardar_downloads_concluirem(pasta)
                    total_exportados += 1
                    log(f"Download do lote {total_exportados} concluído com sucesso!")

                except TimeoutException:
                    log("Botão de exportação não encontrado ou não clicável.", is_error=True)

                lote_atual = 0

                if ultima_pagina:
                    log("Todas as páginas foram processadas. Exportação concluída!")
                    break

        registrar_evento_execucao(
            f"Exportação de XMLs concluída: {total_exportados} lote(s) exportados para {pasta}.",
            "ISS Fortaleza",
        )
        log(f"Exportação finalizada! {total_exportados} lote(s) de XML exportados para: {pasta}")
        progresso(100, "Exportação concluída!")

    except Exception as e:
        log(f"Erro crítico durante a exportação de XMLs: {e}", is_error=True)
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass
