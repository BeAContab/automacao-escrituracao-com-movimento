"""Automação visível do portal da ISS de Fortaleza.

Este módulo concentra o fluxo do navegador para manter `extrair_nf_pdfs.py`
responsável apenas pela extração dos PDFs e pela orquestração da CLI.
"""

from __future__ import annotations

import asyncio
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from playwright.async_api import Page, TimeoutError as PlaywrightTimeoutError, async_playwright
from tratamento_erros import registrar_evento_execucao

URL_ISS_FORTALEZA = "https://iss.fortaleza.ce.gov.br/grpfor/home.seam"

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


class PrestadorNaoEncontradoError(RuntimeError):
    """Sinaliza que o portal não encontrou uma razão social para o CNPJ informado."""

    def __init__(self, cnpj: str, detalhe_tela: str = "") -> None:
        self.cnpj = cnpj
        self.detalhe_tela = detalhe_tela.strip()
        mensagem = f"Nenhuma Razão Social Encontrada para o CNPJ {cnpj}."
        if self.detalhe_tela:
            mensagem = f"{mensagem} Detalhe da tela: {self.detalhe_tela}"
        super().__init__(mensagem)


def normalizar_texto(texto: str) -> str:
    """Remove acentos e padroniza o texto para facilitar o parse."""

    normalizado = unicodedata.normalize("NFD", texto)
    return "".join(
        caractere for caractere in normalizado if unicodedata.category(caractere) != "Mn"
    ).lower()


def limpar_cnpj_para_digitacao(cnpj: str) -> str:
    """Retorna apenas os dígitos do CNPJ para uma digitação mais estável."""

    return re.sub(r"\D+", "", cnpj or "")


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


def registrar_log_funcao2(caminho_log: Path, mensagem: str) -> None:
    """Registra eventos da função 2 com data e hora para rastrear as linhas ignoradas."""

    caminho_log.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with caminho_log.open("a", encoding="utf-8") as arquivo:
        arquivo.write(f"[{timestamp}] {mensagem}\n")


async def pausar_para_depuracao(habilitado: bool, mensagem: str) -> None:
    """Interrompe o fluxo para inspeção manual quando o modo de teste estiver ativo."""

    if not habilitado:
        return
    if not sys.stdin.isatty():
        return

    await asyncio.to_thread(
        input,
        f"{mensagem}\nPressione Enter para continuar com o próximo clique...",
    )


async def clicar_por_id_com_fallback(page: Page, elemento_id: str) -> None:
    """Clica em um elemento pelo ID e usa clique via JavaScript se o portal bloquear o evento."""

    locator = page.locator(f"#{elemento_id.replace(':', '\\:')}")
    try:
        await locator.click(force=True)
    except Exception:
        await page.evaluate(
            """(id) => {
                const elemento = document.getElementById(id);
                if (!elemento) {
                    throw new Error('Elemento não encontrado: ' + id);
                }
                elemento.click();
            }""",
            elemento_id,
        )


async def preencher_modal_pesquisar_cnae(
    page: Page,
    id_cnae_final: str,
    depuracao: bool = False,
) -> None:
    """Preenche o modal de pesquisa de CNAE com o código final da planilha e dispara a busca."""

    registrar_evento_execucao(
        f"Abrindo modal de pesquisa de CNAE para o código {limpar_numero_para_digitacao(id_cnae_final)}",
        "ISS Fortaleza",
    )
    await page.locator("#digitarDocumentoForm\\:pesquisarCnaeModalContainer").wait_for(
        state="visible",
        timeout=10000,
    )
    valor_cnae = limpar_numero_para_digitacao(id_cnae_final)
    campo_cnae = page.locator("#digitarDocumentoForm\\:idFormularioPesquisaCnae\\:idCnaePesquisa")
    await campo_cnae.click(force=True)
    await page.keyboard.press("Control+A")
    await page.keyboard.press("Backspace")
    await page.keyboard.type(valor_cnae, delay=45)
    await pausar_para_depuracao(
        depuracao,
        f"O modal de pesquisa de CNAE está preenchido com {valor_cnae}.",
    )
    # O botão interno do modal dispara a pesquisa AJAX do CNAE escolhido.
    await page.locator("#digitarDocumentoForm\\:idFormularioPesquisaCnae\\:idPesquisar").click(
        force=True
    )
    await page.wait_for_timeout(1200)
    registrar_evento_execucao(
        f"Pesquisa de CNAE enviada para {valor_cnae}",
        "ISS Fortaleza",
    )
    await pausar_para_depuracao(
        depuracao,
        f"A pesquisa de CNAE para {valor_cnae} foi enviada. Confirme o resultado na lista.",
    )
    # Depois da pesquisa, o portal exibe a linha retornada e precisamos selecioná-la.
    await page.locator(
        "#digitarDocumentoForm\\:idFormularioPesquisaCnae\\:idDatatableListaCnae\\:0\\:j_id453"
    ).click(force=True)
    await page.wait_for_timeout(1200)
    registrar_evento_execucao(
        f"CNAE selecionado na lista de resultados: {valor_cnae}",
        "ISS Fortaleza",
    )


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


def carregar_documentos_xlsx(caminho_xlsx: Path) -> list[DocumentoPortalISS]:
    """Lê a planilha `nf_compilado.xlsx` e devolve as linhas prontas para o portal."""

    workbook = load_workbook(caminho_xlsx, data_only=True)
    planilha = workbook.active

    cabecalhos = [celula.value for celula in next(planilha.iter_rows(min_row=1, max_row=1))]
    colunas = {nome: indice for indice, nome in enumerate(cabecalhos)}

    campos_obrigatorios = [
        "ARQUIVO_PDF",
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

        documentos.append(
            DocumentoPortalISS(
                arquivo_pdf=str(linha[colunas["ARQUIVO_PDF"]] or ""),
                cnpj_prestador=str(linha[colunas["CNPJ_PRESTADOR"]] or ""),
                numero_nf=str(linha[colunas["NUMERO_NF"]] or ""),
                id_cnae_final=str(linha[colunas["ID_CNAE_FINAL"]] or ""),
                data_emissao=str(linha[colunas["DATA_EMISSAO"]] or ""),
                descricao_servico=str(linha[colunas["DESCRICAO_SERVICO"]] or ""),
                uf_local_prestacao=str(linha[colunas["UF_LOCAL_PRESTACAO"]] or ""),
                cidade_local_prestacao=str(linha[colunas["CIDADE_LOCAL_PRESTACAO"]] or ""),
                natureza_operacao=str(linha[colunas["NATUREZA_OPERACAO"]] or ""),
                iss_retido=str(linha[colunas["ISS_RETIDO"]] or ""),
                valor_servico=str(linha[colunas["VALOR_SERVICO"]] or ""),
            )
        )

    return documentos


async def abrir_navegador_visivel(url_inicial: str | None = None, depuracao: bool = False):
    """Abre um navegador visível e maximizado para o fluxo manual/assistido."""

    playwright = await async_playwright().start()
    try:
        browser = await playwright.chromium.launch(
            channel="chrome",
            headless=False,
            slow_mo=200 if depuracao else 0,
            args=["--start-maximized"],
        )
    except Exception:
        browser = await playwright.chromium.launch(
            headless=False,
            slow_mo=200 if depuracao else 0,
            args=["--start-maximized"],
        )

    context = await browser.new_context(viewport=None)
    page = await context.new_page()
    if url_inicial:
        await page.goto(
            url_inicial,
            wait_until="domcontentloaded",
            timeout=120000,
        )
        await page.bring_to_front()
    return playwright, browser, context, page


async def clicar_visivelmente(page: Page, texto: str) -> None:
    """Clica em um elemento visível procurando por texto, botão, link ou título."""

    candidatos = [
        page.get_by_role("button", name=texto),
        page.get_by_role("link", name=texto),
        page.locator(f'[title="{texto}"]'),
        page.locator(f'[aria-label="{texto}"]'),
        page.get_by_text(texto, exact=False),
    ]
    for candidato in candidatos:
        try:
            await candidato.first.click(timeout=5000)
            return
        except Exception:
            continue
    raise RuntimeError(f"Não foi possível clicar no item visível: {texto}")


async def clicar_aba_por_texto(page: Page, texto: str) -> None:
    """Clica em uma aba/guia visível usando diferentes estratégias de seleção."""

    candidatos = [
        page.get_by_role("tab", name=texto),
        page.get_by_role("link", name=texto),
        page.get_by_role("button", name=texto),
        page.locator(f'a:has-text("{texto}")'),
        page.locator(f'li:has-text("{texto}")'),
        page.locator(f'span:has-text("{texto}")'),
        page.get_by_text(texto, exact=True),
        page.get_by_text(texto, exact=False),
    ]

    for candidato in candidatos:
        try:
            if await candidato.first.is_visible():
                await candidato.first.click(timeout=5000, force=True)
                return
        except Exception:
            continue

    raise RuntimeError(f"Não foi possível clicar na aba visível: {texto}")


async def digitar_visivelmente(page: Page, seletor: str, texto: str) -> None:
    """Digita como um usuário, com pequenas pausas entre teclas e validação estável.

    Em campos com máscara ou autocomplete, a digitação humana reduz a chance de o
    portal limpar o valor por reprocessamento rápido de eventos. A confirmação usa
    dígitos normalizados para aceitar valores formatados pelo próprio frontend.
    """

    campo = page.locator(seletor)
    await campo.wait_for(state="visible", timeout=10000)
    try:
        await campo.click(force=True)
        await page.keyboard.press("Control+A")
        await page.keyboard.press("Backspace")
        await page.wait_for_timeout(120)
        await page.keyboard.type(texto, delay=85)
    except Exception:
        # Se o campo não aceitar a sequência padrão, tentamos novamente com uma pausa maior.
        await campo.click(force=True)
        await page.wait_for_timeout(150)
        await page.keyboard.press("Control+A")
        await page.keyboard.press("Backspace")
        await page.wait_for_timeout(150)
        await page.keyboard.type(texto, delay=110)

    # Confirma que o valor ficou estável no campo antes de avançar para o próximo clique.
    esperado_normalizado = _somente_digitos(texto)
    for _ in range(20):
        try:
            valor_atual = await campo.input_value()
            if valor_atual.strip() == texto.strip():
                return
            if esperado_normalizado and _somente_digitos(valor_atual) == esperado_normalizado:
                return
        except Exception:
            pass
        await page.wait_for_timeout(150)

    raise RuntimeError(f"O campo {seletor} não permaneceu com o valor esperado: {texto}")


async def selecionar_opcao_por_texto(locator, textos: list[str]) -> None:
    """Seleciona a primeira opção disponível a partir de uma lista de rótulos."""

    ultimo_erro: Exception | None = None
    for texto in textos:
        try:
            await locator.select_option(label=texto)
            return
        except Exception as exc:
            ultimo_erro = exc
        try:
            await locator.select_option(value=texto)
            return
        except Exception as exc:
            ultimo_erro = exc
            continue
    raise RuntimeError(f"Não foi possível selecionar nenhuma opção entre: {textos}") from ultimo_erro


async def aguardar_login_manual(page: Page) -> None:
    """Mostra a instrução e aguarda o usuário concluir o login manual."""

    print("O navegador foi aberto no portal da ISS de Fortaleza.")
    print("Faça o login manualmente e, quando terminar, volte aqui e pressione Enter.")
    input("Pressione Enter somente após o login estar concluído...")
    await page.wait_for_timeout(1000)


async def selecionar_competencia_na_tela(page: Page, competencia: CompetenciaTrabalho) -> None:
    """Seleciona o período na tela de manutenção, priorizando os campos visíveis."""

    # A tela do portal carrega em etapas; esperamos a área de competência ficar disponível antes de ler os selects.
    await page.locator("#manterEscrituracaoForm\\:btnConsultar").wait_for(
        state="visible",
        timeout=120000,
    )
    try:
        texto_pagina = await page.locator("body").inner_text(timeout=5000)
    except Exception:
        texto_pagina = ""
    if competencia.nome_mes in texto_pagina and str(competencia.ano) in texto_pagina:
        registrar_evento_execucao(
            f"Competência {competencia.rotulo} já estava visível na tela de manutenção; seguindo sem alterar selects.",
            "ISS Fortaleza",
        )
        return
    await page.wait_for_timeout(1200)

    mes_textos = [competencia.nome_mes, competencia.nome_mes.upper(), competencia.nome_mes.lower()]
    ano_textos = [str(competencia.ano)]
    combinados = [f"{competencia.nome_mes} {competencia.ano}", f"{competencia.nome_mes}/{competencia.ano}"]
    selecionado = False
    select_mes = None
    select_ano = None
    select_composto = None

    async def _texto_selecionado(select) -> str:
        return await select.evaluate(
            """(el) => {
                const selecionado = el.selectedOptions && el.selectedOptions[0];
                return selecionado ? (selecionado.textContent || '').trim() : '';
            }"""
        )

    async def _coletar_selects_por_fonte() -> list[tuple[str, object]]:
        """Varre a página principal e todos os frames visíveis para encontrar selects."""

        fontes: list[tuple[str, object]] = [("page", page)]
        for indice, frame in enumerate(page.frames, start=1):
            fontes.append((f"frame:{indice}", frame))

        encontrados: list[tuple[str, object]] = []
        for nome_fonte, fonte in fontes:
            try:
                total = await fonte.locator("select").count()
            except Exception:
                continue

            for indice in range(total):
                encontrados.append((nome_fonte, fonte.locator("select").nth(indice)))

        return encontrados

    encontrados: list[tuple[str, object]] = []
    for _ in range(24):
        encontrados = await _coletar_selects_por_fonte()
        if encontrados:
            break
        await page.wait_for_timeout(500)

    if not encontrados:
        registrar_evento_execucao(
            "Nenhum select de competência foi encontrado na página principal nem nos frames.",
            "ISS Fortaleza",
        )
        raise RuntimeError(
            "Não encontrei selects de competência na tela de Manter Escrituração."
        )

    # Primeiro tentamos encontrar selects já compostos com mês e ano no mesmo campo.
    for nome_fonte, select in encontrados:
        try:
            opcoes = [texto.strip() for texto in await select.locator("option").all_text_contents()]
        except Exception:
            continue

        opcoes_norm = [normalizar_texto(texto) for texto in opcoes]
        if any(normalizar_texto(item) in opcoes_norm for item in combinados):
            select_composto = select
            try:
                valor_atual = normalizar_texto(await _texto_selecionado(select))
                if any(normalizar_texto(item) == valor_atual for item in combinados):
                    registrar_evento_execucao(
                        f"Competência {competencia.rotulo} já estava selecionada em um select composto encontrado em {nome_fonte}.",
                        "ISS Fortaleza",
                    )
                    return
            except Exception:
                pass
            for item in combinados:
                try:
                    await select.select_option(label=item)
                    selecionado = True
                    registrar_evento_execucao(
                        f"Competência {competencia.rotulo} selecionada em um select composto encontrado em {nome_fonte}.",
                        "ISS Fortaleza",
                    )
                    break
                except Exception:
                    continue

    # Se não houver um seletor composto, procuramos selects de mês e de ano.
    for nome_fonte, select in encontrados:
        try:
            opcoes = [texto.strip() for texto in await select.locator("option").all_text_contents()]
        except Exception:
            continue

        opcoes_norm = [normalizar_texto(texto) for texto in opcoes]
        if select_mes is None and any(normalizar_texto(mes) in opcoes_norm for mes in mes_textos):
            select_mes = select
        if select_ano is None and any(normalizar_texto(ano) in opcoes_norm for ano in ano_textos):
            select_ano = select

    if select_mes is not None:
        try:
            mes_atual = normalizar_texto(await _texto_selecionado(select_mes))
            if any(normalizar_texto(mes) == mes_atual for mes in mes_textos):
                if select_ano is None:
                    registrar_evento_execucao(
                        f"Mês da competência {competencia.rotulo} já estava selecionado.",
                        "ISS Fortaleza",
                    )
                    return
        except Exception:
            pass
        await selecionar_opcao_por_texto(select_mes, mes_textos)
        selecionado = True

    if select_ano is not None:
        try:
            ano_atual = normalizar_texto(await _texto_selecionado(select_ano))
            if any(normalizar_texto(ano) == ano_atual for ano in ano_textos):
                if selecionado:
                    registrar_evento_execucao(
                        f"Ano da competência {competencia.rotulo} já estava selecionado.",
                        "ISS Fortaleza",
                    )
                    return
        except Exception:
            pass
        await selecionar_opcao_por_texto(select_ano, ano_textos)
        selecionado = True

    if not selecionado:
        diagnostico = []
        for nome_fonte, select in encontrados[:12]:
            try:
                seletor_id = await select.get_attribute("id")
                seletor_name = await select.get_attribute("name")
                opcoes = [texto.strip() for texto in await select.locator("option").all_text_contents()]
                diagnostico.append(
                    f"{nome_fonte} | id={seletor_id or ''} | name={seletor_name or ''} | opcoes={', '.join(opcoes[:8])}"
                )
            except Exception:
                continue
        if diagnostico:
            registrar_evento_execucao(
                "Diagnóstico dos selects de competência encontrados: " + " || ".join(diagnostico),
                "ISS Fortaleza",
            )
        raise RuntimeError(
            "Não consegui reconhecer os campos de competência automaticamente. "
            "Será necessário ajustar os seletores da tela de Manter Escrituração."
        )


async def selecionar_prestador(page: Page, cnpj: str, depuracao: bool = False) -> None:
    """Seleciona o prestador de forma visível, abrindo a sugestão do autocomplete."""

    registrar_evento_execucao(
        f"Iniciando seleção do prestador para o CNPJ {cnpj}",
        "ISS Fortaleza",
    )
    # O rádio de CNPJ precisa estar ativo antes da busca.
    await page.locator("#digitarDocumentoForm\\:tipoPesquisaTomadorRb\\:1").click()
    registrar_evento_execucao("Tipo de pesquisa alterado para CNPJ", "ISS Fortaleza")

    # O portal faz um reload parcial ao alternar para CNPJ; esperamos o campo ser recriado
    # e ficar estável antes de começar a digitar.
    await page.wait_for_timeout(800)
    campo_busca = page.locator("#digitarDocumentoForm\\:cpfPesquisaTomador")
    await campo_busca.wait_for(state="visible", timeout=10000)
    await page.wait_for_timeout(300)

    await digitar_visivelmente(page, "#digitarDocumentoForm\\:cpfPesquisaTomador", limpar_cnpj_para_digitacao(cnpj))
    registrar_evento_execucao(
        f"CNPJ digitado na pesquisa do prestador: {cnpj}",
        "ISS Fortaleza",
    )
    await pausar_para_depuracao(
        depuracao,
        f"O CNPJ {cnpj} foi digitado. Confira o autocomplete antes da seleção.",
    )

    # O portal RichFaces exibe o autocomplete em uma lista visível; clicamos na opção
    # apresentada para que o preenchimento do nome aconteça como no fluxo manual.
    # O DOM expõe o mesmo id em mais de um nó, então restringimos a busca ao bloco
    # da pesquisa do tomador para evitar o strict mode violation do Playwright.
    container = page.locator(
        "#digitarDocumentoForm\\:divPesquisaTomador [id='digitarDocumentoForm:j_id189']"
    ).first
    try:
        await container.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeoutError:
        pass

    texto_container = ""
    try:
        texto_container = await container.inner_text(timeout=2000)
    except Exception:
        try:
            texto_container = await container.text_content(timeout=2000) or ""
        except Exception:
            texto_container = ""

    sugeridos = container.locator("tr.richfaces_suggestionEntry")
    if await sugeridos.count() > 0:
        await sugeridos.first.click()
        registrar_evento_execucao(
            f"Prestador selecionado na lista de sugestões para o CNPJ {cnpj}",
            "ISS Fortaleza",
        )
    else:
        if _texto_indica_prestador_nao_encontrado(texto_container):
            registrar_evento_execucao(
                f"Portal retornou nenhuma razão social para o CNPJ {cnpj}",
                "ISS Fortaleza",
            )
            raise PrestadorNaoEncontradoError(cnpj, texto_container)
        # Fallback visível: navega na lista com o teclado caso a linha não tenha sido localizada.
        await campo_busca.focus()
        await page.keyboard.press("ArrowDown")
        await page.keyboard.press("Enter")

    # Garante que o campo do nome foi preenchido antes de seguir.
    campo_nome = page.locator("#digitarDocumentoForm\\:idNome")
    for _ in range(20):
        try:
            valor = await campo_nome.input_value()
            if valor.strip():
                registrar_evento_execucao(
                    f"Nome do prestador preenchido com sucesso para o CNPJ {cnpj}",
                    "ISS Fortaleza",
                )
                return
        except Exception:
            pass
        await page.wait_for_timeout(250)

    if _texto_indica_prestador_nao_encontrado(texto_container):
        registrar_evento_execucao(
            f"Prestador não encontrado ao finalizar a leitura do campo de nome para o CNPJ {cnpj}",
            "ISS Fortaleza",
        )
        raise PrestadorNaoEncontradoError(cnpj, texto_container)

    raise RuntimeError("O prestador não foi selecionado corretamente; o campo de nome continuou vazio.")


async def preencher_documento_servico(
    page: Page,
    documento: DocumentoPortalISS,
    depuracao: bool = False,
) -> None:
    """Preenche a aba de serviço com os dados de uma linha da planilha."""

    registrar_evento_execucao(
        f"Iniciando preenchimento da aba Serviço para a NF {documento.numero_nf} ({documento.cnpj_prestador})",
        "ISS Fortaleza",
    )
    await selecionar_opcao_por_texto(
        page.locator("#digitarDocumentoForm\\:tipoDocumentoDigitado"),
        ["NFS-e de Outro Município", "NFS-e de outro município", "707"],
    )
    await digitar_visivelmente(
        page,
        "#digitarDocumentoForm\\:numeroDocumentoDigitado",
        limpar_numero_para_digitacao(documento.numero_nf),
    )
    await digitar_visivelmente(
        page,
        "#digitarDocumentoForm\\:dataEmissaoInputDate",
        documento.data_emissao,
    )
    await selecionar_opcao_por_texto(
        page.locator("#digitarDocumentoForm\\:statusNfse"),
        ["NORMAL", "Normal", "472"],
    )
    registrar_evento_execucao(
        f"Campos iniciais da aba Serviço preenchidos para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )
    await page.locator("#digitarDocumentoForm\\:idLinkPesquisarCnae").click(force=True)
    registrar_evento_execucao(
        f"Botão Pesquisar CNAE acionado para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )
    await preencher_modal_pesquisar_cnae(page, documento.id_cnae_final, depuracao=depuracao)
    await digitar_visivelmente(
        page,
        "#digitarDocumentoForm\\:idDescricaoServico",
        documento.descricao_servico,
    )
    await selecionar_opcao_por_texto(
        page.locator("#digitarDocumentoForm\\:comboEscolherEstadoLocalPrestacao"),
        [documento.uf_local_prestacao],
    )
    await selecionar_opcao_por_texto(
        page.locator("#digitarDocumentoForm\\:comboEscolherCidadeLocalPrestacao"),
        [documento.cidade_local_prestacao],
    )
    await selecionar_opcao_por_texto(
        page.locator("#digitarDocumentoForm\\:comboEscolherLocalPrestacao"),
        [documento.natureza_operacao],
    )

    checkbox_iss = page.locator('input[name="digitarDocumentoForm:j_id361"]')
    deve_marcar = normalizar_texto(documento.iss_retido).startswith("sim")
    if deve_marcar and not await checkbox_iss.is_checked():
        await checkbox_iss.click()
    elif not deve_marcar and await checkbox_iss.is_checked():
        await checkbox_iss.click()
    registrar_evento_execucao(
        f"Local de prestação e ISS tratados para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )

    await digitar_visivelmente(
        page,
        "#digitarDocumentoForm\\:idValorServicoPrestado",
        limpar_valor_para_digitacao(documento.valor_servico),
    )
    registrar_evento_execucao(
        f"Aba Serviço preenchida com sucesso para a NF {documento.numero_nf}",
        "ISS Fortaleza",
    )


async def executar_fluxo_iss(
    documentos: list[DocumentoPortalISS],
    competencia: CompetenciaTrabalho,
    caminho_log: Path,
    depuracao: bool = False,
) -> None:
    """Executa o fluxo visual até a revisão final do primeiro prestador localizado."""

    if not documentos:
        raise RuntimeError("A planilha não contém linhas válidas para a automação.")

    playwright, browser, context, page = await abrir_navegador_visivel(
        URL_ISS_FORTALEZA,
        depuracao=depuracao,
    )
    try:
        registrar_evento_execucao(
            f"Fluxo da ISS iniciado com {len(documentos)} documento(s) e competência {competencia.rotulo}",
            "ISS Fortaleza",
        )
        await aguardar_login_manual(page)
        registrar_evento_execucao("Login manual confirmado pelo usuário", "ISS Fortaleza")

        print("Login confirmado. Aguardando o menu do portal ficar disponível...")
        await page.get_by_text("Escrituração", exact=False).first.wait_for(timeout=120000)
        await pausar_para_depuracao(
            depuracao,
            "O portal foi carregado e o login já foi confirmado.",
        )

        print("Acessando Escrituração > Manter Escrituração...")
        await page.locator("a.dropdown-toggle").nth(4).click(force=True)
        await page.wait_for_timeout(500)
        await clicar_por_id_com_fallback(page, "formMenuTopo:menuEscrituracao:j_id80")
        registrar_evento_execucao("Menu Escrituração acionado", "ISS Fortaleza")
        await page.locator("#manterEscrituracaoForm\\:btnConsultar").wait_for(
            state="visible",
            timeout=120000,
        )
        registrar_evento_execucao("Tela Manter Escrituração aberta", "ISS Fortaleza")

        print(
            f"Selecionando a competência {competencia.nome_mes} {competencia.ano} na tela de manutenção..."
        )
        await selecionar_competencia_na_tela(page, competencia)
        registrar_evento_execucao(
            f"Competência selecionada na tela: {competencia.rotulo}",
            "ISS Fortaleza",
        )

        print("Consultando a competência selecionada...")
        await clicar_por_id_com_fallback(page, "manterEscrituracaoForm:btnConsultar")
        registrar_evento_execucao("Botão Consultar acionado", "ISS Fortaleza")
        await page.locator("#manterEscrituracaoForm\\:dataTable\\:0\\:linkEscriturar").wait_for(
            state="visible",
            timeout=120000,
        )

        print("Abrindo a rotina de escrituração...")
        await clicar_por_id_com_fallback(page, "manterEscrituracaoForm:dataTable:0:linkEscriturar")
        registrar_evento_execucao("Botão Escriturar acionado", "ISS Fortaleza")
        await page.wait_for_timeout(1000)

        print("Selecionando a aba Serviços Tomados...")
        await clicar_aba_por_texto(page, "Serviços Tomados")
        registrar_evento_execucao("Aba Serviços Tomados acionada", "ISS Fortaleza")

        await page.locator("#servico_tomado_form\\:seamj_id849").wait_for(
            state="visible",
            timeout=120000,
        )

        print("Abrindo Digitar Documento...")
        await clicar_por_id_com_fallback(page, "servico_tomado_form:seamj_id849")
        registrar_evento_execucao("Tela Digitar Documento aberta", "ISS Fortaleza")
        await page.locator("#digitarDocumentoForm\\:tipoPesquisaTomadorRb\\:1").wait_for(
            state="visible",
            timeout=120000,
        )
        await page.wait_for_timeout(500)
        await pausar_para_depuracao(
            depuracao,
            "A tela Digitar Documento está aberta e pronta para o próximo clique.",
        )

        print(
            f"A planilha possui {len(documentos)} linha(s) válida(s). O fluxo vai tentar o primeiro "
            "prestador encontrado e registrar em log os CNPJs sem razão social."
        )
        documento: DocumentoPortalISS | None = None
        for indice, candidato in enumerate(documentos, start=1):
            print(
                f"Tentando localizar o prestador da linha {indice}/{len(documentos)}: "
                f"{candidato.cnpj_prestador}"
            )
            try:
                await selecionar_prestador(page, candidato.cnpj_prestador, depuracao=depuracao)
            except PrestadorNaoEncontradoError as exc:
                mensagem_log = (
                    f"ARQUIVO_PDF={candidato.arquivo_pdf} | "
                    f"CNPJ_PRESTADOR={candidato.cnpj_prestador} | "
                    f"{exc}"
                )
                registrar_log_funcao2(caminho_log, mensagem_log)
                print(
                    f"Prestador não encontrado para {candidato.cnpj_prestador}; "
                    "avançando para a próxima linha."
                )
                continue

            documento = candidato
            registrar_evento_execucao(
                f"Prestador localizado para a linha {indice}/{len(documentos)}: {candidato.cnpj_prestador}",
                "ISS Fortaleza",
            )
            break

        if documento is None:
            raise RuntimeError(
                "Nenhum prestador da planilha foi localizado no portal. "
                f"Consulte o log em {caminho_log.name}."
            )

        print(
            "Selecionando o prestador e preenchendo a aba Serviço com a linha encontrada na planilha..."
        )
        await clicar_visivelmente(page, "Serviço")
        registrar_evento_execucao("Aba Serviço acionada", "ISS Fortaleza")
        await page.wait_for_timeout(1000)
        await pausar_para_depuracao(
            depuracao,
            "A aba Serviço será preenchida agora com a linha já localizada.",
        )
        await preencher_documento_servico(page, documento, depuracao=depuracao)

        print()
        print("A aba Serviço foi preenchida visivelmente no navegador.")
        print("A gravação do documento ainda não será automatizada, conforme sua orientação.")
        print("Revise a tela no navegador e, se quiser encerrar, volte ao terminal.")
        input("Pressione Enter para encerrar esta sessão automatizada e manter a tela aberta...")
    finally:
        await context.close()
        await browser.close()
        await playwright.stop()


def executar_automacao_iss(
    caminho_xlsx: Path,
    competencia: CompetenciaTrabalho,
    depuracao: bool = False,
) -> None:
    """Ponto de entrada síncrono para a opção 2 da CLI."""

    registrar_evento_execucao(
        f"Automação da ISS solicitada com planilha {caminho_xlsx.resolve()} e competência {competencia.rotulo}",
        "ISS Fortaleza",
    )
    documentos = carregar_documentos_xlsx(caminho_xlsx)
    caminho_log = caminho_xlsx.with_name(f"{caminho_xlsx.stem}_log_funcao2.txt")
    asyncio.run(
        executar_fluxo_iss(
            documentos,
            competencia,
            caminho_log,
            depuracao=depuracao,
        )
    )


def _limpar_texto_exibido(texto: str | None) -> str:
    """Normaliza espaços para deixar a extração mais legível no XLSX."""

    return " ".join((texto or "").split()).strip()


async def _coletar_textos_de_frames(page: Page) -> list[dict[str, str]]:
    """Captura o texto visível dos frames para não perder conteúdo carregado em iframe."""

    frames: list[dict[str, str]] = []
    for indice, frame in enumerate(page.frames, start=1):
        try:
            texto = await frame.locator("body").inner_text(timeout=5000)
        except Exception:
            continue

        texto_limpo = _limpar_texto_exibido(texto)
        if not texto_limpo:
            continue

        frames.append(
            {
                "indice": str(indice),
                "url": frame.url,
                "texto": texto_limpo,
            }
        )

    return frames


async def _coletar_links_visiveis(page: Page) -> list[dict[str, str]]:
    """Lista os links visíveis da página para apoiar o diagnóstico do portal."""

    links: list[dict[str, str]] = []
    total = await page.locator("a").count()
    for indice in range(total):
        elemento = page.locator("a").nth(indice)
        try:
            if not await elemento.is_visible():
                continue
            texto = _limpar_texto_exibido(await elemento.inner_text(timeout=2000))
            href = await elemento.get_attribute("href") or ""
        except Exception:
            continue

        if not texto and not href:
            continue

        links.append(
            {
                "indice": str(len(links) + 1),
                "texto": texto,
                "href": href,
            }
        )

    return links


async def _coletar_elementos_visiveis(page: Page) -> list[dict[str, str]]:
    """Captura campos e botões visíveis para documentar a tela atual do portal."""

    seletor = "input, select, textarea, button"
    elementos: list[dict[str, str]] = []
    total = await page.locator(seletor).count()

    for indice in range(total):
        elemento = page.locator(seletor).nth(indice)
        try:
            if not await elemento.is_visible():
                continue
        except Exception:
            continue

        try:
            dados = await elemento.evaluate(
                """
                (el) => {
                  const tag = el.tagName.toLowerCase();
                  const id = el.id || '';
                  const name = el.name || '';
                  const type = el.getAttribute('type') || '';
                  const placeholder = el.getAttribute('placeholder') || '';
                  const aria = el.getAttribute('aria-label') || '';
                  const disabled = !!el.disabled;
                  const checked = !!el.checked;
                  const texto = tag === 'button' ? (el.innerText || '').trim() : '';
                  const valor = typeof el.value === 'string' ? el.value : '';
                  let selecionado = '';
                  if (tag === 'select') {
                    selecionado = Array.from(el.selectedOptions || [])
                      .map((opcao) => (opcao.textContent || '').trim())
                      .filter(Boolean)
                      .join(' | ');
                  }

                  let rotulo = '';
                  if (id) {
                    const labelFor = document.querySelector(`label[for="${CSS.escape(id)}"]`);
                    if (labelFor) {
                      rotulo = (labelFor.innerText || '').trim();
                    }
                  }
                  if (!rotulo) {
                    const labelPai = el.closest('label');
                    if (labelPai) {
                      rotulo = (labelPai.innerText || '').trim();
                    }
                  }

                  return { tag, type, id, name, placeholder, aria, disabled, checked, texto, valor, selecionado, rotulo };
                }
                """
            )
        except Exception:
            continue

        elementos.append(
            {
                "indice": str(len(elementos) + 1),
                "tag": str(dados.get("tag", "")),
                "tipo": str(dados.get("type", "")),
                "id": str(dados.get("id", "")),
                "name": str(dados.get("name", "")),
                "rotulo": str(dados.get("rotulo", "")),
                "texto": str(dados.get("texto", "")),
                "valor": str(dados.get("valor", "")),
                "selecionado": str(dados.get("selecionado", "")),
                "placeholder": str(dados.get("placeholder", "")),
                "aria_label": str(dados.get("aria", "")),
                "checked": "sim" if dados.get("checked") else "nao",
                "disabled": "sim" if dados.get("disabled") else "nao",
            }
        )

    return elementos


async def _coletar_tabelas_visiveis(page: Page) -> list[dict[str, object]]:
    """Lê as tabelas visíveis da página para exportar a grade apresentada ao usuário."""

    return await page.evaluate(
        """
        () => {
          const visivel = (el) => !!(el && el.getClientRects && el.getClientRects().length);

          return Array.from(document.querySelectorAll('table'))
            .filter(visivel)
            .map((table, indice) => {
              const legenda = (table.querySelector('caption')?.innerText || '').trim();
              const linhas = Array.from(table.querySelectorAll('tr'))
                .map((linha) => Array.from(linha.querySelectorAll('th,td')).map((celula) => (celula.innerText || '').replace(/\\s+/g, ' ').trim()))
                .filter((linha) => linha.some(Boolean));

              return {
                indice: indice + 1,
                legenda,
                linhas,
              };
            });
        }
        """
    )


async def extrair_dados_da_pagina_iss(page: Page) -> dict[str, object]:
    """Coleta o conteúdo visível da tela logada para exportação e análise posterior."""

    titulo = await page.title()
    url = page.url

    try:
        texto_bruto = await page.locator("body").inner_text(timeout=10000)
    except Exception:
        texto_bruto = ""

    dados = {
        "titulo": titulo,
        "url": url,
        "extraido_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "texto_visivel": [linha.strip() for linha in texto_bruto.splitlines() if linha.strip()],
        "frames": await _coletar_textos_de_frames(page),
        "tabelas": await _coletar_tabelas_visiveis(page),
        "elementos": await _coletar_elementos_visiveis(page),
        "links": await _coletar_links_visiveis(page),
    }
    return dados


def _aplicar_cabecalho(aba) -> None:
    """Padroniza o estilo visual das planilhas geradas pela extração."""

    for celula in aba[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")


def exportar_extracao_iss(dados: dict[str, object], destino: Path) -> None:
    """Exporta a página extraída para XLSX com abas auxiliares de diagnóstico."""

    wb = Workbook()

    aba_resumo = wb.active
    aba_resumo.title = "Resumo"
    aba_resumo.append(["Campo", "Valor"])
    aba_resumo.append(["Título", dados.get("titulo", "")])
    aba_resumo.append(["URL", dados.get("url", "")])
    aba_resumo.append(["Extraído em", dados.get("extraido_em", "")])
    aba_resumo.append(["Quantidade de linhas de texto", len(dados.get("texto_visivel", []))])
    aba_resumo.append(["Quantidade de frames", len(dados.get("frames", []))])
    aba_resumo.append(["Quantidade de tabelas", len(dados.get("tabelas", []))])
    aba_resumo.append(["Quantidade de elementos", len(dados.get("elementos", []))])
    aba_resumo.append(["Quantidade de links", len(dados.get("links", []))])
    _aplicar_cabecalho(aba_resumo)

    aba_texto = wb.create_sheet("Texto")
    aba_texto.append(["Linha", "Conteúdo"])
    for indice, linha in enumerate(dados.get("texto_visivel", []), start=1):
        aba_texto.append([indice, linha])
    _aplicar_cabecalho(aba_texto)

    aba_frames = wb.create_sheet("Frames")
    aba_frames.append(["Frame", "URL", "Conteúdo"])
    for frame in dados.get("frames", []):
        aba_frames.append([
            frame.get("indice", ""),
            frame.get("url", ""),
            frame.get("texto", ""),
        ])
    _aplicar_cabecalho(aba_frames)

    aba_tabelas = wb.create_sheet("Tabelas")
    aba_tabelas.append(["Tabela", "Linha", "Coluna", "Valor"])
    for tabela in dados.get("tabelas", []):
        indice_tabela = tabela.get("indice", "")
        legenda = tabela.get("legenda", "")
        linhas = tabela.get("linhas", [])
        if legenda:
            aba_tabelas.append([indice_tabela, 0, 0, f"Legenda: {legenda}"])
        for indice_linha, linha in enumerate(linhas, start=1):
            for indice_coluna, valor in enumerate(linha, start=1):
                aba_tabelas.append([indice_tabela, indice_linha, indice_coluna, valor])
    _aplicar_cabecalho(aba_tabelas)

    aba_elementos = wb.create_sheet("Elementos")
    aba_elementos.append(
        [
            "Indice",
            "Tag",
            "Tipo",
            "ID",
            "Name",
            "Rotulo",
            "Texto",
            "Valor",
            "Selecionado",
            "Placeholder",
            "Aria label",
            "Checked",
            "Disabled",
        ]
    )
    for elemento in dados.get("elementos", []):
        aba_elementos.append(
            [
                elemento.get("indice", ""),
                elemento.get("tag", ""),
                elemento.get("tipo", ""),
                elemento.get("id", ""),
                elemento.get("name", ""),
                elemento.get("rotulo", ""),
                elemento.get("texto", ""),
                elemento.get("valor", ""),
                elemento.get("selecionado", ""),
                elemento.get("placeholder", ""),
                elemento.get("aria_label", ""),
                elemento.get("checked", ""),
                elemento.get("disabled", ""),
            ]
        )
    _aplicar_cabecalho(aba_elementos)

    aba_links = wb.create_sheet("Links")
    aba_links.append(["Indice", "Texto", "Href"])
    for link in dados.get("links", []):
        aba_links.append([link.get("indice", ""), link.get("texto", ""), link.get("href", "")])
    _aplicar_cabecalho(aba_links)

    for aba in wb.worksheets:
        ajustar_largura_colunas(aba)

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)


def executar_extracao_portal_iss(destino_xlsx: Path = Path("iss_extracao.xlsx")) -> None:
    """Ponto de entrada síncrono para extrair o conteúdo visível do portal logado."""

    async def _executar() -> None:
        registrar_evento_execucao(
            f"Extração avulsa do portal ISS iniciada com destino {destino_xlsx.resolve()}",
            "ISS Fortaleza",
        )
        playwright, browser, context, page = await abrir_navegador_visivel(URL_ISS_FORTALEZA)
        try:
            await aguardar_login_manual(page)

            print()
            print("Quando estiver na tela exata da qual deseja extrair os dados, volte ao terminal.")
            input("Pressione Enter para capturar a página atual...")

            await page.wait_for_timeout(1000)
            dados = await extrair_dados_da_pagina_iss(page)
            exportar_extracao_iss(dados, destino_xlsx)
            registrar_evento_execucao(
                f"Extração avulsa concluída com sucesso em {destino_xlsx.resolve()}",
                "ISS Fortaleza",
            )

            print(f"Extração concluída com sucesso: {destino_xlsx.resolve()}")
            print("O XLSX contém resumo, texto visível, frames, tabelas, elementos e links.")
            input("Pressione Enter para encerrar esta sessão automatizada e fechar o navegador...")
        finally:
            await context.close()
            await browser.close()
            await playwright.stop()

    asyncio.run(_executar())
