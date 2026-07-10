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
)
from tratamento_erros import registrar_evento_execucao

# Seletores XPath para a tela de Consulta de NFS-e
XPATH_LINK_CONSULTAR_NFSE    = "//*[@id='homeForm:divHotLinks']/div[4]/a/h4"
XPATH_ABA_COMPETENCIA        = "//*[@id='consultarnfseForm:competencia_prestador_tab_lbl']"
XPATH_BTN_CONSULTAR          = "//*[@id='consultarnfseForm:j_id237']"
XPATH_BTN_SELECIONAR_PAGINA  = "//*[@id='consultarnfseForm:j_id324']"
XPATH_BTN_EXPORTAR_XML       = "//*[@id='consultarnfseForm:j_id321']/div[1]/input[3]"

# Base ID do calendário de competência na tela de Consulta
# ATENÇÃO: NÃO incluir o sufixo 'Header' — _abrir_editor_calendario o concatena internamente
BASE_ID_COMPETENCIA = "consultarnfseForm:competencia"

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
    """Tenta clicar no botão de avançar página. Retorna False se a última página foi atingida."""
    try:
        # Encontra o botão de próxima página (o texto usa o caractere › em vez de > e está direto no td)
        botoes = driver.find_elements(By.XPATH, "//td[contains(@class, 'rich-datascr-button') and contains(text(), '›')]")
        
        for botao in botoes:
            if botao.is_displayed():
                # Se o botão estiver inativo, significa que chegamos na última página
                classes = botao.get_attribute("class") or ""
                if "rich-datascr-button-inactive" in classes:
                    return False
                
                # Se estiver ativo, clica para avançar
                botao.click()
                time.sleep(2.5)
                return True
    except Exception as e:
        print(f"Erro ao tentar avançar página: {e}")
    return False


def _selecionar_competencia_consulta(driver, mes: int, ano: int) -> None:
    """
    Seleciona mês e ano no calendário inline RichFaces da tela Consulta de NFS-e.

    Este calendário é sempre visível (não é popup), portanto NÃO usa PopupButton.
    O fluxo correto é clicar diretamente no botão de edição (.rich-calendar-tool-btn)
    dentro do elemento pai do calendário para abrir a grade de meses/anos.
    """
    # Seletor do botão de cabeçalho que abre a grade de meses/anos
    SELETOR_EDITOR_BTN = f"[id='{BASE_ID_COMPETENCIA}'] .rich-calendar-tool-btn"

    # Passo 1: Aguarda e clica no botão de edição para abrir a grade de meses/anos
    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, SELETOR_EDITOR_BTN))
    ).click()
    time.sleep(0.5)

    # Passo 2: Aguarda o editor carregar (verifica pelo botão OK como indicador de prontidão)
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located(
            (By.ID, f"{BASE_ID_COMPETENCIA}DateEditorButtonOk")
        )
    )

    # Passo 3: Seleciona o mês pelo índice 0-based (Janeiro=M0, ..., Dezembro=M11)
    id_mes = f"{BASE_ID_COMPETENCIA}DateEditorLayoutM{mes - 1}"
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, id_mes))
    ).click()

    # Passo 4: Localiza o ano-alvo navegando entre décadas se necessário
    for _ in range(20):  # limite de 20 iterações para evitar loop infinito
        anos_visiveis = {}
        for i in range(10):
            try:
                el = driver.find_element(
                    By.ID, f"{BASE_ID_COMPETENCIA}DateEditorLayoutY{i}"
                )
                texto = el.text.strip()
                if texto.isdigit():
                    anos_visiveis[int(texto)] = el
            except NoSuchElementException:
                continue

        # Clica no ano se ele estiver visível na grade atual
        if ano in anos_visiveis:
            anos_visiveis[ano].click()
            break

        if not anos_visiveis:
            raise RuntimeError("Grade de anos do calendário vazia ou não renderizada.")

        # Define a direção de navegação entre décadas
        menor = min(anos_visiveis)
        sinal = "<" if ano < menor else ">"
        botoes = driver.find_elements(
            By.CSS_SELECTOR, f"[id='{BASE_ID_COMPETENCIA}'] .rich-calendar-editor-btn"
        )
        clicou = False
        for btn in botoes:
            if btn.text.strip() == sinal and btn.is_displayed():
                btn.click()
                clicou = True
                time.sleep(0.3)
                break
        if not clicou:
            raise RuntimeError(
                f"Botão de navegação '{sinal}' não encontrado no calendário de anos."
            )
    else:
        raise RuntimeError(
            f"Não foi possível localizar o ano {ano} no calendário após 20 tentativas."
        )

    # Passo 5: Confirma a seleção clicando em OK
    driver.find_element(By.ID, f"{BASE_ID_COMPETENCIA}DateEditorButtonOk").click()
    time.sleep(0.5)


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

        # Passo 3: Selecionar mês/ano no calendário com a função dedicada a esta tela
        log(f"Selecionando competência {competencia.mes:02d}/{competencia.ano}...")
        _selecionar_competencia_consulta(driver, competencia.mes, competencia.ano)
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
