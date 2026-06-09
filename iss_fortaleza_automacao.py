"""Automação visível do portal da ISS de Fortaleza.

Este módulo concentra o fluxo do navegador para manter `extrair_nf_pdfs.py`
responsável apenas pela extração dos PDFs e pela orquestração da CLI.
"""

from __future__ import annotations

import asyncio
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from playwright.async_api import Page, TimeoutError as PlaywrightTimeoutError, async_playwright

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
    data_emissao: str
    descricao_servico: str
    uf_local_prestacao: str
    cidade_local_prestacao: str
    natureza_operacao: str
    iss_retido: str
    valor_servico: str


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


async def abrir_navegador_visivel():
    """Abre um navegador visível e maximizado para o fluxo manual/assistido."""

    playwright = await async_playwright().start()
    try:
        browser = await playwright.chromium.launch(
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
        )
    except Exception:
        browser = await playwright.chromium.launch(
            headless=False,
            args=["--start-maximized"],
        )

    context = await browser.new_context(viewport=None)
    page = await context.new_page()
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


async def digitar_visivelmente(page: Page, seletor: str, texto: str) -> None:
    """Limpa o campo e digita o conteúdo como um usuário faria."""

    campo = page.locator(seletor)
    await campo.click()
    await page.keyboard.press("Control+A")
    await page.keyboard.press("Backspace")
    await page.keyboard.type(texto, delay=20)


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

    await page.wait_for_timeout(1000)

    selects = page.locator("select")
    total_selects = await selects.count()
    if total_selects == 0:
        raise RuntimeError(
            "Não encontrei selects de competência na tela de Manter Escrituração."
        )

    mes_textos = [competencia.nome_mes, competencia.nome_mes.upper(), competencia.nome_mes.lower()]
    ano_textos = [str(competencia.ano)]
    combinados = [f"{competencia.nome_mes} {competencia.ano}", f"{competencia.nome_mes}/{competencia.ano}"]
    selecionado = False

    # Primeiro tentamos encontrar selects já compostos com mês e ano no mesmo campo.
    for indice in range(total_selects):
        select = selects.nth(indice)
        try:
            opcoes = [texto.strip() for texto in await select.locator("option").all_text_contents()]
        except Exception:
            continue

        opcoes_norm = [normalizar_texto(texto) for texto in opcoes]
        if any(normalizar_texto(item) in opcoes_norm for item in combinados):
            for item in combinados:
                try:
                    await select.select_option(label=item)
                    selecionado = True
                    break
                except Exception:
                    continue

    # Se não houver um seletor composto, procuramos selects de mês e de ano.
    for indice in range(total_selects):
        select = selects.nth(indice)
        try:
            opcoes = [texto.strip() for texto in await select.locator("option").all_text_contents()]
        except Exception:
            continue

        opcoes_norm = [normalizar_texto(texto) for texto in opcoes]
        if any(normalizar_texto(mes) in opcoes_norm for mes in mes_textos):
            await selecionar_opcao_por_texto(select, mes_textos)
            selecionado = True
            continue
        if any(normalizar_texto(ano) in opcoes_norm for ano in ano_textos):
            await selecionar_opcao_por_texto(select, ano_textos)
            selecionado = True

    if not selecionado:
        raise RuntimeError(
            "Não consegui reconhecer os campos de competência automaticamente. "
            "Será necessário ajustar os seletores da tela de Manter Escrituração."
        )


async def selecionar_prestador(page: Page, cnpj: str) -> None:
    """Seleciona o prestador de forma visível, abrindo a sugestão do autocomplete."""

    # O rádio de CNPJ precisa estar ativo antes da busca.
    await page.locator("#digitarDocumentoForm\\:tipoPesquisaTomadorRb\\:1").click()

    campo_busca = page.locator("#digitarDocumentoForm\\:cpfPesquisaTomador")
    await digitar_visivelmente(page, "#digitarDocumentoForm\\:cpfPesquisaTomador", limpar_cnpj_para_digitacao(cnpj))

    # O portal RichFaces exibe o autocomplete em uma lista visível; clicamos na opção
    # apresentada para que o preenchimento do nome aconteça como no fluxo manual.
    container = page.locator("#digitarDocumentoForm\\:j_id189")
    try:
        await container.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeoutError:
        pass

    sugeridos = container.locator("tr.richfaces_suggestionEntry")
    if await sugeridos.count() > 0:
        await sugeridos.first.click()
    else:
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
                return
        except Exception:
            pass
        await page.wait_for_timeout(250)

    raise RuntimeError("O prestador não foi selecionado corretamente; o campo de nome continuou vazio.")


async def preencher_documento_servico(page: Page, documento: DocumentoPortalISS) -> None:
    """Preenche a aba de serviço com os dados de uma linha da planilha."""

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

    await digitar_visivelmente(
        page,
        "#digitarDocumentoForm\\:idValorServicoPrestado",
        limpar_valor_para_digitacao(documento.valor_servico),
    )


async def executar_fluxo_iss(
    documentos: list[DocumentoPortalISS],
    competencia: CompetenciaTrabalho,
) -> None:
    """Executa o fluxo visual até a revisão final do primeiro documento."""

    if not documentos:
        raise RuntimeError("A planilha não contém linhas válidas para a automação.")

    playwright, browser, context, page = await abrir_navegador_visivel()
    try:
        await page.goto(
            URL_ISS_FORTALEZA,
            wait_until="domcontentloaded",
            timeout=120000,
        )
        await page.bring_to_front()
        await aguardar_login_manual(page)

        print("Login confirmado. Aguardando o menu do portal ficar disponível...")
        await page.get_by_text("Escrituração", exact=False).first.wait_for(timeout=120000)

        print("Acessando Escrituração > Manter Escrituração...")
        await clicar_visivelmente(page, "Escrituração")
        await page.wait_for_timeout(800)
        await clicar_visivelmente(page, "Manter Escrituração")

        print(
            f"Selecionando a competência {competencia.nome_mes} {competencia.ano} na tela de manutenção..."
        )
        await selecionar_competencia_na_tela(page, competencia)

        print("Consultando a competência selecionada...")
        await clicar_visivelmente(page, "Consultar")
        await page.wait_for_timeout(1000)

        print("Abrindo a rotina de escrituração...")
        await clicar_visivelmente(page, "Escriturar")
        await page.wait_for_timeout(1000)

        print("Entrando em Serviços Tomados...")
        await clicar_visivelmente(page, "Serviços Tomados")
        await page.wait_for_timeout(1000)

        print("Abrindo Digitar Documento...")
        await clicar_visivelmente(page, "Digitar Documento")
        await page.wait_for_timeout(1000)

        documento = documentos[0]
        print(
            f"A planilha possui {len(documentos)} linha(s) válida(s), mas esta versão prepara apenas a primeira "
            "até a etapa anterior a GRAVAR DOCUMENTO."
        )
        print(
            "Selecionando o prestador e preenchendo a aba Serviço com a primeira linha da planilha..."
        )
        await selecionar_prestador(page, documento.cnpj_prestador)
        await clicar_visivelmente(page, "Serviço")
        await page.wait_for_timeout(1000)
        await preencher_documento_servico(page, documento)

        print()
        print("A aba Serviço foi preenchida visivelmente no navegador.")
        print("A gravação do documento ainda não será automatizada, conforme sua orientação.")
        print("Revise a tela no navegador e, se quiser encerrar, volte ao terminal.")
        input("Pressione Enter para encerrar esta sessão automatizada e manter a tela aberta...")
    finally:
        await context.close()
        await browser.close()
        await playwright.stop()


def executar_automacao_iss(caminho_xlsx: Path, competencia: CompetenciaTrabalho) -> None:
    """Ponto de entrada síncrono para a opção 2 da CLI."""

    documentos = carregar_documentos_xlsx(caminho_xlsx)
    asyncio.run(executar_fluxo_iss(documentos, competencia))


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
        playwright, browser, context, page = await abrir_navegador_visivel()
        try:
            await page.goto(
                URL_ISS_FORTALEZA,
                wait_until="domcontentloaded",
                timeout=120000,
            )
            await page.bring_to_front()
            await aguardar_login_manual(page)

            print()
            print("Quando estiver na tela exata da qual deseja extrair os dados, volte ao terminal.")
            input("Pressione Enter para capturar a página atual...")

            await page.wait_for_timeout(1000)
            dados = await extrair_dados_da_pagina_iss(page)
            exportar_extracao_iss(dados, destino_xlsx)

            print(f"Extração concluída com sucesso: {destino_xlsx.resolve()}")
            print("O XLSX contém resumo, texto visível, frames, tabelas, elementos e links.")
            input("Pressione Enter para encerrar esta sessão automatizada e fechar o navegador...")
        finally:
            await context.close()
            await browser.close()
            await playwright.stop()

    asyncio.run(_executar())
