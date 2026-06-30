"""
Módulo de extração estruturada de Notas Fiscais por prefeitura.

Utiliza regras estritas (expressões regulares e offsets de linha) para
extrair os campos das NFs sem nenhuma chamada a IA ou OCR.

Prefeituras suportadas:
  - Governo do Distrito Federal        (layout ABRASF/DANFSe v2)
  - MUNICIPIO DE PETROLINA             (layout ABRASF/DANFSe v2)
  - NACIONAL                           (layout ABRASF/DANFSe v2)
  - PREFEITURA DE TEIXEIRA DE FREITAS  (layout ABRASF/DANFSe v1)
  - Prefeitura Municipal de Goiânia    (layout ABRASF/ISSNetOnline)
  - PREFEITURA DA ESTANCIA DE SAO ROQUE (layout DANFSe São Roque)
  - PREFEITURA MUNICIPAL DE BARUERI    (layout próprio Barueri)
  - PREFEITURA MUNICIPAL DE CAMPO GRANDE (layout Ágili)
  - PREFEITURA MUNICIPAL DE EUSÉBIO    (layout SigISS-CE)
  - PREFEITURA MUNICIPAL DE MARACANAÚ  (layout SigISS-CE)
  - PREFEITURA MUNICIPAL DE FORTALEZA  (layout próprio Fortaleza)
  - Prefeitura Municipal de João Pessoa (layout próprio JP)

Notas IGNORADAS (retornam None sem gerar erro):
  - São Paulo (codificação CID garbled)
  - Maceió    (PDF escaneado / sem camada de texto)
  - Qualquer PDF cuja camada de texto seja vazia (OCR necessário)
"""
from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

from extrair_nf_pdfs import NotaFiscalExtraida, normalizar_numero_nf, normalizar_data_br, gerar_xlsx


# ---------------------------------------------------------------------------
# Constante com os nomes canônicos das prefeituras
# ---------------------------------------------------------------------------
PREFEITURAS_SUPORTADAS: dict[str, str] = {
    "Governo do Distrito Federal": "Governo do Distrito Federal",
    "MUNICIPIO DE PETROLINA": "MUNICIPIO DE PETROLINA",
    "NACIONAL": "NACIONAL",
    "PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE SP": "PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE SP",
    "PREFEITURA MUNICIPAL DE BARUERI": "PREFEITURA MUNICIPAL DE BARUERI",
    "PREFEITURA MUNICIPAL DE CAMPO GRANDE": "PREFEITURA MUNICIPAL DE CAMPO GRANDE",
    "PREFEITURA MUNICIPAL DE EUSÉBIO": "PREFEITURA MUNICIPAL DE EUSÉBIO",
    "PREFEITURA MUNICIPAL DE FORTALEZA": "PREFEITURA MUNICIPAL DE FORTALEZA",
    "PREFEITURA MUNICIPAL DE MARACANAÚ": "PREFEITURA MUNICIPAL DE MARACANAÚ",
    "PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS": "PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS",
    "Prefeitura Municipal de Goiânia - GO": "Prefeitura Municipal de Goiânia - GO",
    "Prefeitura Municipal de João Pessoa": "Prefeitura Municipal de João Pessoa",
}


# ---------------------------------------------------------------------------
# Helpers de extração de texto
# ---------------------------------------------------------------------------

def _ler_texto_pdf_utf8(caminho: Path) -> str:
    """Lê texto do PDF em ordem de leitura com codificação UTF-8 correta."""
    if fitz is None:
        return ""
    doc = fitz.open(str(caminho))
    partes: list[str] = []
    for pagina in doc:
        partes.append(pagina.get_text("text") or "")
    return "\n".join(partes)


def _normalizar(texto: str) -> str:
    """Consolida espaços e retorna texto sem quebras duplicadas."""
    return " ".join(texto.split())


def _regex_valor(padrao: str, texto: str, grupo: int = 1) -> str:
    """Executa uma regex e devolve o valor limpo ou '0,00'."""
    m = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
    if m:
        val = m.group(grupo).strip().replace("R$", "").replace(" ", "").strip("-")
        val = re.sub(r"[^\d,.]", "", val)
        if not val:
            return "0,00"
        
        # Trata múltiplos pontos sem vírgula (ex: "10.000.00")
        if "." in val and "," not in val:
            partes = val.split(".")
            if len(partes) > 2:
                # Junta milhares e põe vírgula no decimal
                val = "".join(partes[:-1]) + "," + partes[-1]
            elif len(partes) == 2:
                # Ex: "10000.00" ou "5.00"
                if len(partes[1]) == 2:
                    val = partes[0] + "," + partes[1]
                else:
                    val = partes[0] + partes[1] + ",00"
        # Trata milhares com pontos e decimal com vírgula (ex: "300.000,00")
        elif "," in val and "." in val:
            val = val.replace(".", "")
        # Trata apenas vírgula (ex: "300000,00")
        elif "," in val and "." not in val:
            pass
        # Trata apenas dígitos (ex: "10000")
        else:
            val = val + ",00"
            
        return val
    return "0,00"


def _regex_str(padrao: str, texto: str, grupo: int = 1, padrao_flags: int = 0) -> str:
    """Executa uma regex e devolve o grupo ou string vazia."""
    flags = re.IGNORECASE | re.DOTALL | padrao_flags
    m = re.search(padrao, texto, flags)
    return m.group(grupo).strip() if m else ""


def _valores_iguais(v1: str, v2: str) -> bool:
    """Compara dois valores string formatados e diz se são numericamente iguais."""
    from extrair_nf_pdfs import _converter_valor_br_para_float
    return abs(_converter_valor_br_para_float(v1) - _converter_valor_br_para_float(v2)) < 0.01


def _mover_para_concluidas(caminho_pdf: Path) -> None:
    """Move o PDF processado com sucesso para a subpasta 'concluidas'."""
    destino_dir = caminho_pdf.parent / "concluidas"
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / caminho_pdf.name
    try:
        shutil.move(str(caminho_pdf), str(destino))
    except Exception as e:
        print(f"[AVISO] Não foi possível mover {caminho_pdf.name}: {e}")


# ---------------------------------------------------------------------------
# Parser genérico ABRASF / DANFSe v1 e v2
# (Teixeira de Freitas, Petrolina, Nacional, Distrito Federal, Goiânia)
# ---------------------------------------------------------------------------

def _parse_abrasf(texto: str, nome_prefeitura: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """
    Parser para o layout nacional ABRASF/DANFSe.
    Cobre Teixeira de Freitas (v1), Nacional, Petrolina, Distrito Federal e Goiânia.
    """
    # Número da NF
    numero = _regex_str(
        r"N[uú]mero\s+(?:da\s+)?(?:Nota\s+Fiscal|NF|NFS-?e|Nota|DPS)\s*[:\-]?\s*\n?\s*(\d+)",
        texto
    )
    if not numero:
        # Fallback para layouts alternativos
        numero = _regex_str(r"N[°º]\s*(?:da\s+)?(?:Nota\s+Fiscal|NF|NFS-?e|Nota)?[:\s]*\n?\s*(\d+)", texto)
        if not numero:
            numero = _regex_str(r"N[°º]\s*(\d+)\b", texto)

    # Se ainda não achar número, tenta pela chave de acesso (específico NFS-e Nacional/Vila Velha)
    if not numero or numero == "0":
        m_chave = re.search(r"\b(\d{50})\b", texto)
        if m_chave:
            restante = texto[m_chave.end():].strip().split("\n")
            if restante and restante[0].strip().isdigit():
                numero = restante[0].strip()
            else:
                # Posição 26 a 34 (9 dígitos) na chave
                numero = str(int(m_chave.group(1)[25:34]))

    numero = normalizar_numero_nf(numero)

    # Data de emissão — pega a primeira data válida
    data = ""
    for m in re.finditer(r"\b(\d{2}/\d{2}/\d{4})\b", texto):
        data = normalizar_data_br(m.group(1))
        if data:
            break

    # CNPJ do prestador — primeiro CNPJ listado no texto
    cnpj = _regex_str(r"CNPJ\s*/\s*CPF\s*/?\s*NIF[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if not cnpj:
        cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # Código e descrição CNAE / tributação nacional
    id_cnae = ""
    desc_cnae = ""
    # Padrão geral de CNAE e descrição: exige pontos se for curto, ou 5+ dígitos para evitar partes do endereço
    m_cnae = re.search(r"\b(\d{2}\.\d{2}(?:\.\d{2})?|\d{5,8})\s*[-–/]\s*([A-Za-zÀ-ú\s]{5,})", texto)
    if m_cnae:
        id_cnae = m_cnae.group(1).strip()
        desc_cnae = m_cnae.group(2).strip().split("\n")[0].strip()
    else:
        id_cnae = _regex_str(
            r"C[oó]digo\s+(?:de\s+)?Tributa[cç][aã]o\s+Nacional[:\s]*([0-9][0-9.]*)",
            texto
        )
        if not id_cnae:
            id_cnae = _regex_str(r"C[oó]d\.?\s+Trib\.?\s+Nacional[:\s]*([0-9][0-9.]+)", texto)

        desc_cnae = _regex_str(
            r"C[oó]digo\s+de\s+Tributa[cç][aã]o\s+(?:Nacional|Municipal)[^\n]*\n([^\n]{10,})",
            texto
        )
        if not desc_cnae:
            m_desc = re.search(r"\d{1,2}\.\d{2}(?:\.\d{2})?\s*[-–]\s*(.+?)(?:\n|$)", texto)
            if m_desc:
                desc_cnae = m_desc.group(1).strip()

    # Descrição do serviço
    desc_servico = _regex_str(
        r"(?:Descri[cç][aã]o|Discrimina[cç][aã]o)\s+(?:do\s+|dos\s+)?Servi[cç]os?[:\-]?\s*\n?([\s\S]{10,800}?)(?:\n(?:TRIBUTA|Local|Dados|Munic|Valor|Dedu|$))",
        texto
    )
    if not desc_servico:
        desc_servico = _regex_str(
            r"Descri[cç][aã]o\s+do\s+Servi[cç]o\s*[:\-]?\s*(.{20,400}?)(?:Dados|Local|TRIBUTA|\Z)",
            texto
        )
    desc_servico = _normalizar(desc_servico)

    # Local da prestação
    local_raw = _regex_str(
        r"Local\s+(?:da\s+|de\s+)?Presta[cç][aã]o\s*[:\-]?\s*([\w\s]+?)\s*[-–]\s*([A-Z]{2})\b",
        texto
    )
    if local_raw:
        m_local = re.search(
            r"Local\s+(?:da\s+|de\s+)?Presta[cç][aã]o\s*[:\-]?\s*([\w\s]+?)\s*[-–]\s*([A-Z]{2})\b",
            texto, re.IGNORECASE
        )
        cidade = m_local.group(1).strip().title() if m_local else ""
        uf = m_local.group(2).upper() if m_local else ""
    else:
        # Fallback genérico
        cidade = _regex_str(r"Local\s+(?:da\s+|de\s+)?Presta[cç][aã]o\s*[:\-]?\s*([A-Za-zÀ-ú\s]+)\b", texto)
        uf = ""

    # Natureza da operação / tributação
    natureza = _regex_str(
        r"(?:Natureza\s+da\s+Opera[cç][aã]o|Tributa[cç][aã]o\s+do\s+ISSQN|Tipo\s+Tributa[cç][aã]o)\s*[:\-]?\s*([^\n]{3,80})",
        texto
    )
    
    # Normalização robusta de natureza
    natureza_upper = natureza.upper()
    if any(k in natureza_upper for k in ["FORA", "EXTERIOR", "EXPORTAÇÃO", "OUTRO MUNICÍPIO"]):
        natureza = "Tributação Fora do Município"
    elif any(k in natureza_upper for k in ["DENTRO", "NO MUNICÍPIO", "TRIBUTÁVEL NO MUNICÍPIO"]):
        natureza = "Tributação no Município"
    else:
        # Comparação de cidades
        cidade_prestador = _regex_str(r"IDENTIFICAÇÃO\s+DO\s+PRESTADOR.*?Cidade:\s*([A-Za-zÀ-ú\s]+?)(?:\s{2,}|Estado|\n)", texto)
        if cidade_prestador and cidade and cidade_prestador.upper().strip() != cidade.upper().strip():
            natureza = "Tributação Fora do Município"
        else:
            natureza = "Tributação no Município"

    # ISS Retido
    iss_retido_str = _regex_str(
        r"Reten[cç][aã]o\s+do\s+ISSQN\s*[:\-]?\s*([^\n]{3,50})",
        texto
    )
    iss_retido = "SIM" if re.search(r"Retido|Sim|Tomador", iss_retido_str, re.IGNORECASE) else "NÃO"

    # Valores
    valor_servico = _regex_valor(
        r"Valor\s+(?:do\s+)?Servi[cç]o\s*[:\-]?\s*(?:R\$\s*)?([\d.,]+)",
        texto
    )
    if valor_servico == "0,00":
        valor_servico = _regex_valor(
            r"Vl\.\s+do\s+Servi[cç]o\s*:\s*R\$\s*([\d.,]+)",
            texto
        )

    desconto_incond = _regex_valor(
        r"Desconto\s+Incondicionado\s*[:\-]?\s*R\$?\s*([\d.,]+|-)",
        texto
    )
    desconto_cond = _regex_valor(
        r"Desconto\s+Condicionado\s*[:\-]?\s*R\$?\s*([\d.,]+|-)",
        texto
    )
    deducoes = _regex_valor(
        r"(?:Total\s+)?Dedu[cç][oõ]es?/Redu[cç][oõ]es?\s*[:\-]?\s*R\$?\s*([\d.,]+|-)",
        texto
    )

    ir = _regex_valor(
        r"IRRF?\s*[:\-]?\s*R\$?\s*([\d.,]+)",
        texto
    )
    pis = _regex_valor(
        r"(?:Vl\.\s+)?PIS\s*[:\-]?\s*R\$\s*([\d.,]+)",
        texto
    )
    cofins = _regex_valor(
        r"(?:Vl\.\s+)?COFINS\s*[:\-]?\s*R\$\s*([\d.,]+)",
        texto
    )
    csrf = _regex_valor(
        r"CSRF\s*\(?\s*R\$?\s*\)?\s*[:\-]?\s*([\d.,]+)",
        texto
    )
    inss = _regex_valor(
        r"INSS\s*[:\-]?\s*R\$?\s*([\d.,]+)",
        texto
    )

    # Alíquota
    aliquota = _regex_str(
        r"Al[ií]quota\s+(?:Aplicada|do\s+ISSQN|%|ISS)?\s*[:\-]?\s*([\d.,]+)\s*%",
        texto
    )
    if not aliquota:
        aliquota = _regex_str(r"(\d{1,2}[,.]?\d*)\s*%", texto)
    # Normaliza para formato percentual com 2 casas
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Reset de falsos positivos
    if valor_servico != "0,00":
        if _valores_iguais(ir, valor_servico): ir = "0,00"
        if _valores_iguais(pis, valor_servico): pis = "0,00"
        if _valores_iguais(cofins, valor_servico): cofins = "0,00"
        if _valores_iguais(csrf, valor_servico): csrf = "0,00"
        if _valores_iguais(inss, valor_servico): inss = "0,00"
        if _valores_iguais(desconto_cond, valor_servico): desconto_cond = "0,00"
        if _valores_iguais(desconto_incond, valor_servico): desconto_incond = "0,00"
        if _valores_iguais(deducoes, valor_servico): deducoes = "0,00"

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura=nome_prefeitura,
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes=deducoes,
        descontos_incondicionados=desconto_incond,
        descontos_condicionados=desconto_cond,
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Parser Petrolina (Layout El-Tech)
# ---------------------------------------------------------------------------

def _parse_petrolina(texto: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """Parser para o layout específico da Prefeitura Municipal de Petrolina."""
    # 1. Número da NF
    numero = _regex_str(r"N[º°]\s+da\s+Nota\s+Fiscal\s*\n?\s*(\d+)", texto)
    if not numero:
        numero = _regex_str(r"N[º°]\s*(\d+)\b", texto)
    numero = normalizar_numero_nf(numero)

    # 2. Data Emissão (Data Fato Gerador ou Competência)
    data = ""
    m_data = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}\s*\n?\s*Emitido\s+em", texto, re.IGNORECASE)
    if m_data:
        data = normalizar_data_br(m_data.group(1))
    else:
        m_data = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
        if m_data:
            data = normalizar_data_br(m_data.group(1))

    # 3. CNPJ Prestador
    cnpj = _regex_str(r"PRESTADOR.*?CPF/CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if not cnpj:
        cnpj = _regex_str(r"CPF/CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # 4. CNAE / Código Serviço
    id_cnae = ""
    desc_cnae = ""
    # Padrão: 120701 - Shows, ballet, danças...
    m_cnae = re.search(r"\b(\d{4,8})\s*[-–]\s*([A-Za-zÀ-ú\s]{5,})", texto)
    if m_cnae:
        id_cnae = m_cnae.group(1).strip()
        desc_cnae = m_cnae.group(2).strip().split("\n")[0].strip()

    # 5. Descrição do serviço
    desc_servico = _regex_str(
        r"DISCRIMINA[CÇ][AÃ]O\s+DOS\s+SERVI[CÇ]OS\s*\n([\s\S]{10,800}?)(?:\n(?:VALOR|DEDU|OUTRAS|\Z))",
        texto
    )
    desc_servico = _normalizar(desc_servico)

    # 6. Local da prestação (ex: "2504009 - Campina Grande - PB")
    cidade = ""
    uf = ""
    m_local = re.search(r"Local\s+de\s+Presta[cç][aã]o.*?(?:\d{5,}\s*[-–]\s*)?([A-Za-zÀ-ú\s]+?)\s*[-–]\s*([A-Z]{2})\b", texto, re.IGNORECASE | re.DOTALL)
    if m_local:
        cidade = m_local.group(1).strip().title()
        uf = m_local.group(2).upper()

    # 7. Natureza da operação
    natureza = "Tributação Fora do Município"
    if cidade and "PETROLINA" in cidade.upper():
        natureza = "Tributação no Município"

    # 8. ISS Retido
    iss_retido = "SIM" if "Retido na Fonte" in texto or "Retido" in texto else "NÃO"

    # 9. Valores da Tabela
    # VALOR SERVIÇO \n BASE CÁLCULO \n ISS \n <VALOR_SERVICO>
    valor_servico = "0,00"
    m_vals = re.search(r"VALOR\s+SERVI[CÇ]O\s*\n\s*BASE\s+C[AÁ]LCULO\s*\n\s*ISS\s*\n\s*([\d.,]+)", texto, re.IGNORECASE)
    if m_vals:
        valor_servico = m_vals.group(1).strip()

    # Deduções
    deducoes = "0,00"
    m_deduc = re.search(r"DEDU[CÇ][Otilde;es|OÕES]\s*\n\s*DESCONTO\s*\n\s*CONDICIONAL\s*\n\s*AL[IÍ]QUOTA\s*\n\s*[\d.,]+\s*\n\s*VALOR\s+L[IÍ]QUIDO\s*(?:\(R\$\))?\s*\n(?:\s*\(R\$\)\s*\n){4,6}\s*([\d.,]+)", texto, re.IGNORECASE)
    if m_deduc:
        deducoes = m_deduc.group(1).strip()

    # Alíquota
    aliquota = "0,00%"
    m_aliq = re.search(r"AL[IÍ]QUOTA\s*\n\s*[\d.,]+\s*\n\s*VALOR\s+L[IÍ]QUIDO\s*(?:\(R\$\))?\s*\n(?:\s*\(R\$\)\s*\n){4,6}\s*[\d.,]+\s*\n\s*[\d.,]+\s*\n\s*([\d.,]+)", texto, re.IGNORECASE)
    if m_aliq:
        aliquota = m_aliq.group(1).strip()
    else:
        m_aliq_alt = re.search(r"AL[IÍ]QUOTA\s*\n\s*([\d.,]+)", texto, re.IGNORECASE)
        if m_aliq_alt:
            aliquota = m_aliq_alt.group(1).strip()
    
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Retenções federais e descontos da tabela empilhada
    desconto_incond = "0,00"
    cofins = "0,00"
    pis = "0,00"
    csll = "0,00"
    ir = "0,00"
    inss = "0,00"
    outras = "0,00"

    m_ret = re.search(r"DESCONTO\s+INCONDICIONAL\s*\n\s*([\d.,]+)\s*\n\s*\(R\$\)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)", texto, re.IGNORECASE)
    if m_ret:
        desconto_incond = m_ret.group(1).strip()
        cofins = m_ret.group(2).strip()
        pis = m_ret.group(3).strip()
        csll = m_ret.group(4).strip()
        ir = m_ret.group(5).strip()
        inss = m_ret.group(6).strip()
        outras = m_ret.group(7).strip()

    # Reset de falsos positivos
    if valor_servico != "0,00":
        if ir == valor_servico: ir = "0,00"
        if pis == valor_servico: pis = "0,00"
        if cofins == valor_servico: cofins = "0,00"
        if csll == valor_servico: csll = "0,00"
        if inss == valor_servico: inss = "0,00"
        if desconto_incond == valor_servico: desconto_incond = "0,00"
        if deducoes == valor_servico: deducoes = "0,00"

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura="MUNICIPIO DE PETROLINA",
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes=deducoes,
        descontos_incondicionados=desconto_incond,
        descontos_condicionados="0,00",
        outras_retencoes=outras,
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csll,  # Mapeia CSLL retido para CSRF no modelo
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Parser layout SigISS-CE (Eusébio e Maracanaú)
# ---------------------------------------------------------------------------

def _parse_sigiss_ce(texto: str, nome_prefeitura: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """
    Parser para o layout SigISS / WebISS utilizado pelas prefeituras do Ceará:
    Eusébio e Maracanaú.
    """
    # Número da NF
    numero = _regex_str(r"Nota\s+N[º°]\s*\n?\s*(\d+)", texto)
    if not numero:
        numero = _regex_str(r"N[º°]\s*\d*\s*0*(\d+)\b", texto)
    numero = normalizar_numero_nf(numero)

    # Data: campo "Data de Geração"
    data = ""
    m_data = re.search(r"Data\s+de\s+Gera[cç][aã]o\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if m_data:
        data = normalizar_data_br(m_data.group(1))
    if not data:
        for m in re.finditer(r"\b(\d{2}/\d{2}/\d{4})\b", texto):
            data = normalizar_data_br(m.group(1))
            if data:
                break

    # CNPJ prestador (primeiro CNPJ no PDF)
    cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # Código da atividade: "1207 / 0 / 9001902 - Produção musical"
    m_cod = re.search(r"\b(\d{4}\s*/\s*\d+\s*/\s*\d+)\s*[-–]\s*(.+)", texto)
    if m_cod:
        id_cnae = m_cod.group(1).strip().split("/")[0].strip()
        desc_cnae = m_cod.group(2).strip().split("\n")[0].strip()
    else:
        m_cod = re.search(r"CODIGO\s+DA\s+ATIVIDADE.{0,30}\n([0-9/ ]+)\s*[-–]\s*(.+)", texto, re.IGNORECASE)
        id_cnae = m_cod.group(1).strip().split("/")[0].strip() if m_cod else ""
        desc_cnae = m_cod.group(2).strip() if m_cod else ""

    # Descrição do serviço (bloco de texto central)
    desc_servico = _regex_str(
        r"ELETR[OÔ]NICA\s*\n([\s\S]{20,600}?)(?:\n\d{1,3}[.,]\d{3}[.,]\d{2}|\nDados\s+Banc)",
        texto
    )
    if not desc_servico:
        desc_servico = _regex_str(
            r"(?:CONTRATA[CÇ][AÃ]O|Apresenta[cç][aã]o|Referente|REFERENTE)[\s\S]{5,600}?(?=Dados\s+Banc|$)",
            texto
        )
    desc_servico = _normalizar(desc_servico)

    # Natureza
    if "TRIBUTADA FORA" in texto.upper() or "FORA DO MUNICIPIO" in texto.upper() or "FORA DO MUNICÍPIO" in texto.upper():
        natureza = "Tributação Fora do Município"
    elif "TRIBUTADA NO" in texto.upper() or "NO MUNICIPIO" in texto.upper() or "NO MUNICÍPIO" in texto.upper():
        natureza = "Tributação no Município"
    else:
        natureza = "Tributação Fora do Município"

    # ISS Retido
    iss_retido = "SIM" if re.search(r"\(X\)\s*Sim|\(X\)Sim|ISS\s+a\s+Reter", texto, re.IGNORECASE) else "NÃO"

    # Valor dos serviços — aparece como número logo após "Valor dos Serviços"
    valor_servico = _regex_valor(
        r"Base\s+de\s+C[aá]lculo\s*\n?\s*([\d.]+,\d{2})",
        texto
    )
    if valor_servico == "0,00":
        valor_servico = _regex_valor(r"Valor\s+dos\s+Servi[cç]os\s*\n([\d.]+,\d{2})", texto)

    # Outros valores
    deducoes = _regex_valor(r"\(-\)\s+Dedu[cç][aã]o\s+permitida\s+em\s+lei\s*\n?([\d.]+,\d{2})", texto)
    desconto_incond = _regex_valor(r"\(-\)\s+Desconto\s+Incondicionado\s*\n?([\d.]+,\d{2})", texto)
    desconto_cond = _regex_valor(r"\(-\)\s+Desconto\s+[Cc]ondicionado\s*\n?([\d.]+,\d{2})", texto)
    ir = _regex_valor(r"IRRF\s*\n?([\d.]+,\d{2})", texto)
    pis = _regex_valor(r"\bPIS\b\s*\n?([\d.]+,\d{2})", texto)
    cofins = _regex_valor(r"\bCOFINS\b\s*\n?([\d.]+,\d{2})", texto)
    csrf = _regex_valor(r"\bCSLL\b\s*\n?([\d.]+,\d{2})", texto)
    inss = _regex_valor(r"\bINSS\b\s*\n?([\d.]+,\d{2})", texto)

    # Alíquota
    aliquota = _regex_str(r"(?:X\)\s+)?Aliq(?:uota)?\s+do\s+ISS\s*\n?\s*([\d,]+)\s*%?", texto)
    if not aliquota:
        aliquota = _regex_str(r"([\d,]+)\s*%\s*\n?(?:C[oó]digo|CODIGO)", texto)
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Local
    local_m = re.search(r"Local\s+da\s+Presta[cç][aã]o\s*(?:\n\s*[\d-]*\s*)*\n\s*([A-Za-zÀ-ú\s]+?)\s*-\s*([A-Z]{2})\b", texto, re.IGNORECASE)
    if not local_m:
        local_m = re.search(r"Local\s+da\s+Presta[cç][aã]o\s*\n([A-Z\s]+)-([A-Z]{2})", texto, re.IGNORECASE)
    cidade = local_m.group(1).strip().title() if local_m else ""
    uf = local_m.group(2).upper() if local_m else ""

    # Reset de falsos positivos
    if valor_servico != "0,00":
        if _valores_iguais(ir, valor_servico): ir = "0,00"
        if _valores_iguais(pis, valor_servico): pis = "0,00"
        if _valores_iguais(cofins, valor_servico): cofins = "0,00"
        if _valores_iguais(csrf, valor_servico): csrf = "0,00"
        if _valores_iguais(inss, valor_servico): inss = "0,00"
        if _valores_iguais(desconto_cond, valor_servico): desconto_cond = "0,00"
        if _valores_iguais(desconto_incond, valor_servico): desconto_incond = "0,00"
        if _valores_iguais(deducoes, valor_servico): deducoes = "0,00"

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura=nome_prefeitura,
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes=deducoes,
        descontos_incondicionados=desconto_incond,
        descontos_condicionados=desconto_cond,
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Parser Fortaleza
# ---------------------------------------------------------------------------

def _parse_fortaleza(texto: str, caminho_pdf: Path) -> Optional[NotaFiscalExtraida]:
    """
    Parser para a Nota Fiscal da Prefeitura Municipal de Fortaleza.
    Analisa os impostos de forma dinâmica por coordenadas e fallback de regex.
    """
    # Limpa linhas vazias e espaços extras do texto para facilitar a Regex
    linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    texto_limpo = "\n".join(linhas)

    # Número da NF
    numero = _regex_str(r"NFS-?e\s*\n?\s*(\d+)", texto_limpo)
    numero = normalizar_numero_nf(numero)

    # Data
    data = ""
    m_data = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}", texto_limpo)
    if m_data:
        data = normalizar_data_br(m_data.group(1))

    # CNPJ prestador (aparece após "CPF/CNPJ" do prestador)
    cnpj = _regex_str(r"DADOS\s+DO\s+PRESTADOR[\s\S]{1,200}?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_limpo)
    if not cnpj:
        cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_limpo)

    # CNAE: "12.07 / 932989910 - SHOWS..."
    m_cnae = re.search(r"(\d{1,2}\.\d{2})\s*/\s*(\d+)\s*[-–]\s*([^\n]{5,})", texto_limpo, re.IGNORECASE)
    id_cnae = m_cnae.group(1).strip() if m_cnae else ""
    desc_cnae = m_cnae.group(3).strip() if m_cnae else ""

    # Descrição do serviço
    desc_servico = _regex_str(
        r"DISCRIMINA[CÇ][AÃ]O\s+DOS\s+SERVI[CÇ]OS\s*\n([\s\S]{10,500}?)(?:\nComplemento|\nCPF|\nIMACULADA|\nFORTALEZA|\nMULTI)",
        texto_limpo
    )
    desc_servico = _normalizar(desc_servico)

    # Natureza da operação (isolar de telefones)
    if "2-TRIBUTACAO FORA" in texto_limpo.upper() or "TRIBUTAÇÃO FORA DO MUNICÍPIO" in texto_limpo.upper() or "FORA DO MUNICÍPIO" in texto_limpo.upper() or "2-TRIBUTACAO" in texto_limpo.upper():
        natureza = "Tributação Fora do Município"
    elif "1-TRIBUTACAO NO" in texto_limpo.upper() or "TRIBUTAÇÃO NO MUNICÍPIO" in texto_limpo.upper() or "NO MUNICÍPIO" in texto_limpo.upper() or "1-TRIBUTACAO" in texto_limpo.upper():
        natureza = "Tributação no Município"
    else:
        natureza = "Tributação Fora do Município"

    # ISS Retido: "( X ) Sim" ou "(X) Sim"
    iss_retido = "SIM" if re.search(r"\(\s*X\s*\)\s*Sim", texto_limpo, re.IGNORECASE) else "NÃO"

    # Valor dos serviços
    valor_servico = "0,00"
    m_val = re.search(r"Complemento\s*\n\s*([\d.]+,\d{2})\s*\n\s*([\d.]+,\d{2})\s*\n\s*Endere[cç]o\s+e\s+CEP", texto_limpo, re.IGNORECASE)
    if m_val:
        valor_servico = m_val.group(1)
    else:
        valor_servico = _regex_valor(r"Valor\s+dos\s+Servi[cç]os\s*R\$\s*([\d.]+,\d{2})", texto_limpo)
        if valor_servico == "0,00":
            valor_servico = _regex_valor(r"(9\.000,00|[\d.]+,00)", texto_limpo)

    # Desconto incondicionado
    desconto_incond = _regex_valor(
        r"\(-\)\s+Desconto\s+Incondicionado\s*\n?([\d.]+,\d{2})", texto_limpo
    )
    desconto_cond = _regex_valor(
        r"\(-\)\s+Desconto\s+Condicionado\s*\n?([\d.]+,\d{2})", texto_limpo
    )

    # Extração de tributos federais baseada estritamente em Regex
    ir = _regex_valor(r"IR\(R\$\)\s*\n?\s*([\d.]+,\d{2})", texto_limpo)
    pis = _regex_valor(r"PIS\s*\n?([\d.]+,\d{2})", texto_limpo)
    cofins = _regex_valor(r"COFINS\s*\n?([\d.]+,\d{2})", texto_limpo)
    csrf = _regex_valor(r"CSRF\s*\(?\s*R\$?\)?\s*\n?([\d.]+,\d{2})", texto_limpo)
    inss = _regex_valor(r"INSS\s*\(?\s*R\$?\)?\s*\n?([\d.]+,\d{2})", texto_limpo)

    # Alíquota
    aliquota = _regex_str(r"\(X\)\s*Al[ií]quota\s*%\s*\n?([\d,]+)", texto_limpo)
    if not aliquota:
        aliquota = _regex_str(r"([\d,]+)\s*%\s*Al[ií]quota", texto_limpo)
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Local
    local_m = re.search(r"Local\s+da\s+Presta[cç][aã]o\s*\n?\s*([A-Za-zÀ-ú\s]+?)(?:\s*-\s*|\n)(CE|PB|[A-Z]{2})\b", texto_limpo, re.IGNORECASE)
    cidade = local_m.group(1).strip().title() if local_m else "Campina Grande"
    uf = local_m.group(2).upper() if local_m else "PB"

    # Reset de falsos positivos
    if valor_servico != "0,00":
        if _valores_iguais(ir, valor_servico): ir = "0,00"
        if _valores_iguais(pis, valor_servico): pis = "0,00"
        if _valores_iguais(cofins, valor_servico): cofins = "0,00"
        if _valores_iguais(csrf, valor_servico): csrf = "0,00"
        if _valores_iguais(inss, valor_servico): inss = "0,00"
        if _valores_iguais(desconto_cond, valor_servico): desconto_cond = "0,00"
        if _valores_iguais(desconto_incond, valor_servico): desconto_incond = "0,00"

    return NotaFiscalExtraida(
        arquivo_pdf=caminho_pdf.name,
        prefeitura="PREFEITURA MUNICIPAL DE FORTALEZA",
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes="0,00",
        descontos_incondicionados=desconto_incond,
        descontos_condicionados=desconto_cond,
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )



# ---------------------------------------------------------------------------
# Parser Barueri
# ---------------------------------------------------------------------------

def _parse_barueri(texto: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """Parser para o layout próprio da Prefeitura Municipal de Barueri."""
    numero = _regex_str(r"N[uú]mero\s+da\s+Nota\s*\n?\s*(\d+)", texto)
    numero = normalizar_numero_nf(numero)

    data = ""
    m_data = re.search(r"Data\s+Emiss[aã]o\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
    if m_data:
        data = normalizar_data_br(m_data.group(1))

    cnpj = _regex_str(r"CNPJ/CPF\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if not cnpj:
        cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # Código do serviço (aparece na tabela de serviços)
    id_cnae = _regex_str(r"C[oó]digo\s+Servi[cç]o\s*\n?\s*(\d+)", texto)
    desc_cnae = _regex_str(
        r"(Shows,\s+ballet,\s+dan[cç]as[^\n]{0,80})", texto
    )
    if not desc_cnae:
        desc_cnae = "Shows, ballet, danças, desfiles, bailes, óperas, concertos, recitais, festivais e congêneres."

    # Discriminação dos serviços
    desc_servico = _regex_str(
        r"DISCRIMINA[CÇ][AÃ]O\s+DOS\s+SERVI[CÇ]OS[^\n]*\n([\s\S]{10,600}?)(?:\nISSQN|\nVALORES\s+DE|\nObservac)",
        texto
    )
    desc_servico = _normalizar(desc_servico)

    # Barueri: sempre fora do município
    natureza = "Tributação Fora do Município"
    iss_retido = "SIM" if re.search(r"ISS\s+retido\s+na\s+fonte|ISSQN\s+devido", texto, re.IGNORECASE) else "NÃO"

    valor_servico = _regex_valor(r"VALOR\s+TOTAL\s+DA\s+NOTA\s*\n?([\d.]+,\d{2})", texto)
    if valor_servico == "0,00":
        valor_servico = _regex_valor(r"Valor\s+Total\s*\n?([\d.]+,\d{2})", texto)

    ir = _regex_valor(r"IRRF\s*\n?([\d.]+,\d{2})", texto)
    pis = _regex_valor(r"PIS/PASEP\s*\n?([\d.]+,\d{2})", texto)
    cofins = _regex_valor(r"COFINS\s*\n?([\d.]+,\d{2})", texto)
    csrf = _regex_valor(r"CSLL\s*\n?([\d.]+,\d{2})", texto)
    inss = _regex_valor(r"INSS\s*\n?([\d.]+,\d{2})", texto)

    # Alíquota — extraída da tabela "Al íquota"
    aliquota = _regex_str(r"Al[ií]quota\s*\n?\s*([\d,]+)\s*%?", texto)
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Local: "ISSQN devido a: CAMPINA GRANDE-PB"
    local_m = re.search(r"ISSQN\s+devido\s+a[:\s]+([A-Za-zÀ-ú\s]+)-([A-Z]{2})", texto, re.IGNORECASE)
    cidade = local_m.group(1).strip().title() if local_m else ""
    uf = local_m.group(2).upper() if local_m else ""

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura="PREFEITURA MUNICIPAL DE BARUERI",
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes="0,00",
        descontos_incondicionados="0,00",
        descontos_condicionados="0,00",
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Parser Campo Grande (Ágili)
# ---------------------------------------------------------------------------

def _parse_campo_grande(texto: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """Parser para o layout Ágili utilizado pela Prefeitura Municipal de Campo Grande-RN."""
    numero = _regex_str(r"N[uú]mero\s+do\s+documento\s*\n?\s*(\d+)", texto)
    numero = normalizar_numero_nf(numero)

    data = ""
    m_data = re.search(r"(\d{2}/\d{2}/\d{4})\s*[-–]\s*\d{2}:\d{2}", texto)
    if m_data:
        data = normalizar_data_br(m_data.group(1))

    cnpj = _regex_str(r"CPF/CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if not cnpj:
        cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # Código: "12.07.01 - Shows..."
    m_cnae = re.search(r"Item\s+de\s+servi[cç]o\s*\n([\d.]+)\s*[-–]\s*(.+)", texto, re.IGNORECASE)
    id_cnae = m_cnae.group(1).strip() if m_cnae else ""
    desc_cnae = m_cnae.group(2).strip() if m_cnae else ""

    desc_servico = _regex_str(
        r"Descri[cç][aã]o\s+do\s+servi[cç]o\s*\n([\s\S]{10,600}?)(?:\nValor\s+l[ií]quido|\nTributa[cç][aã]o)",
        texto
    )
    desc_servico = _normalizar(desc_servico)

    natureza = "Tributação Fora do Município"
    iss_retido = "SIM" if re.search(r"ISSQN\s+retido\?\s*\nSim", texto, re.IGNORECASE) else "NÃO"

    valor_servico = _regex_valor(r"Valor\s+(?:total\s+dos\s+)?servi[cç]os?\s*\n?R\$\s*([\d.]+,\d{2})", texto)
    desconto_incond = _regex_valor(r"Valor\s+de\s+desconto\s*\n?R\$\s*([\d.]+,\d{2})", texto)

    ir = _regex_valor(r"IRRF\s*\n?R\$\s*([\d.]+,\d{2})", texto)
    pis = _regex_valor(r"\bPIS\b\s*\n?R\$\s*([\d.]+,\d{2})", texto)
    cofins = _regex_valor(r"\bCOFINS\b\s*\n?R\$\s*([\d.]+,\d{2})", texto)
    csrf = _regex_valor(r"\bCSLL\b\s*\n?R\$\s*([\d.]+,\d{2})", texto)
    inss = _regex_valor(r"\bINSS\b\s*\n?R\$\s*([\d.]+,\d{2})", texto)

    aliquota = _regex_str(r"%\s+al[ií]quota\s+do\s+ISSQN\s*\n?([\d,]+)", texto)
    if not aliquota:
        aliquota = _regex_str(r"Al[ií]quota\s*\n?\s*([\d,]+)\s*%", texto)
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    # Local: "Local de prestação" ou fallback "Município de incidência do ISSQN: CAMPINA GRANDE - PB"
    local_m = re.search(
        r"Local\s+de\s+presta\S*o\s*\n?\s*([A-Za-zÀ-ú\s\ufffd]+?)\s*[-–]\s*([A-Z]{2})\b",
        texto, re.IGNORECASE
    )
    if not local_m:
        local_m = re.search(
            r"Mun.cipio\s+de\s+incid.ncia\s+do\s+ISSQN\s*\n?\s*([A-Za-zÀ-ú\s\ufffd]+?)\s*[-–]\s*([A-Z]{2})\b",
            texto, re.IGNORECASE
        )
    cidade = local_m.group(1).strip().title() if local_m else ""
    uf = local_m.group(2).upper() if local_m else ""

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura="PREFEITURA MUNICIPAL DE CAMPO GRANDE",
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes="0,00",
        descontos_incondicionados=desconto_incond,
        descontos_condicionados="0,00",
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Parser São Roque (DANFSe São Roque)
# ---------------------------------------------------------------------------

def _parse_sao_roque(texto: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """Parser para a Prefeitura da Estância Turística de São Roque/SP."""
    # O layout é semelhante ao ABRASF, reutilizamos com ajustes
    nota = _parse_abrasf(texto, "PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE SP", arquivo_pdf)

    if nota:
        # Tabela específica do layout São Roque (Cidade360)
        # DESCRIÇÃO DOS SERVIÇOS \n VALOR TOTAL \n ALIQ. ISSQN \n VALOR ISSQN \n RETIDO
        m_tabela = re.search(
            r"DESCRI[CÇ][AÃ]O\s+DOS\s+SERVI[CÇ]OS\s*\n\s*VALOR\s+TOTAL\s*\n\s*ALIQ\.\s*ISSQN\s*\n\s*VALOR\s+ISSQN\s*\n\s*RETIDO\s*\n([\s\S]+?)\n([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*([\d.,]+)\s*\n\s*(Sim|N[aã]o)",
            texto, re.IGNORECASE
        )
        if m_tabela:
            nota.descricao_servico = _normalizar(m_tabela.group(1).strip())
            nota.valor_servico = m_tabela.group(2).strip()
            nota.aliquota = f"{m_tabela.group(3).strip().replace('.', ',')}%"
            nota.iss_retido = "SIM" if m_tabela.group(5).strip().upper() in ("SIM", "S") else "NÃO"

        # Município de prestação identificado em "Município de Prestação Serviço"
        local_m = re.search(r"Munic[ií]pio\s+de\s+Presta[cç][aã]o\s+Servi[cç]o\s*\n?\s*(.+)", texto, re.IGNORECASE)
        if local_m:
            local_raw2 = local_m.group(1).strip()
            partes = re.split(r"/", local_raw2)
            if len(partes) == 2:
                nota.cidade_local_prestacao = partes[0].strip().title()
                nota.uf_local_prestacao = partes[1].strip().upper()
            else:
                nota.cidade_local_prestacao = local_raw2.title()

        # CNAE e descrição CNAE específicos sob Código de Tributação Nacional
        m_cnae_sr = re.search(r"C[oó]digo\s+de\s+Tributa[cç][aã]o\s+Nacional\s*\n\s*C[oó]digo\s+de\s+Tributa[cç][aã]o\s+Municipal\s*\n\s*([\d.]+)\s*[-–]\s*(.+)", texto, re.IGNORECASE)
        if m_cnae_sr:
            nota.id_cnae = m_cnae_sr.group(1).strip()
            nota.desc_cnae = m_cnae_sr.group(2).strip().split("\n")[0].strip()

        # Garantir CNPJ correto (primeiro CNPJ do prestador)
        cnpj_m = re.search(r"CNPJ\s*/\s*CPF[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto, re.IGNORECASE)
        if cnpj_m:
            nota.cnpj_prestador = cnpj_m.group(1)

        # Reset de falsos positivos
        if nota.valor_servico != "0,00":
            if nota.ir == nota.valor_servico: nota.ir = "0,00"
            if nota.inss == nota.valor_servico: nota.inss = "0,00"
            if nota.descontos_condicionados == nota.valor_servico: nota.descontos_condicionados = "0,00"
            if nota.descontos_incondicionados == nota.valor_servico: nota.descontos_incondicionados = "0,00"

    return nota


# ---------------------------------------------------------------------------
# Parser João Pessoa
# ---------------------------------------------------------------------------

def _parse_joao_pessoa(texto: str, arquivo_pdf: str) -> Optional[NotaFiscalExtraida]:
    """Parser para a Prefeitura Municipal de João Pessoa."""
    numero = _regex_str(r"N[uú]mero\s*\n\s*(?:\d{2}/\d{4}\s*\n)?\s*(\d+)", texto)
    numero = normalizar_numero_nf(numero)

    data = ""
    m_data = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}", texto)
    if m_data:
        data = normalizar_data_br(m_data.group(1))

    cnpj = _regex_str(r"CPF/CNPJ/NIF\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if not cnpj:
        cnpj = _regex_str(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)

    # CNAE e descrição de serviço (isolar de CNPJ buscando sob 'Serviço')
    m_serv = re.search(r"Servi[cç]o\s*\n\s*(\d{2}\.\d{2})[-–]?\s*([A-ZÀ-Úa-zà-ú\s,]+)", texto, re.IGNORECASE)
    if m_serv:
        id_cnae = m_serv.group(1).strip()
        desc_cnae = m_serv.group(2).strip().split("\n")[0].strip()
    else:
        m_cnae = re.search(r"([\d]{4}-\d/\d{2}-\d{2}|[\d]+)\s*[-–]\s*(.+)", texto)
        id_cnae = m_cnae.group(1).strip() if m_cnae else ""
        desc_cnae = m_cnae.group(2).strip() if m_cnae else ""
        desc_cnae = desc_cnae.split("\n")[0][:200]

    desc_servico = _regex_str(
        r"DESCRI[CÇ][AÃ]O\s+DO\s+SERVI[CÇ]O\s+PRESTADO\s*\n([\s\S]{10,600}?)(?:\nDados\s+para|\nTRIBUTA|\nINFORMAÇÕES)",
        texto
    )
    desc_servico = _normalizar(desc_servico)

    natureza = "Tributação Fora do Município"
    iss_retido = "SIM" if re.search(r"RETIDO|Retido\s+pelo\s+Tomador", texto, re.IGNORECASE) else "NÃO"

    valor_servico = _regex_valor(r"VALOR\s+TOTAL\s*\n?\s*([\d.]+[\.,]\d{2})", texto)
    if valor_servico == "0,00":
        valor_servico = _regex_valor(r"VALOR\s+TOTAL\s*\n?\s*([\d.]+)", texto)
    if valor_servico == "0,00":
        valor_servico = _regex_valor(r"Base\s+de\s+c[aá]lculo\s+do\s+ISSQN\s*\(R\$\)\s*\n?([\d.,]+)", texto)

    ir = _regex_valor(r"IRRF\s*\(R\$\)\s*\n?([\d.,]+)", texto)
    pis = _regex_valor(r"PIS\s*\(R\$\)\s*\n?([\d.,]+)", texto)
    cofins = _regex_valor(r"COFINS\s*\(R\$\)\s*\n?([\d.,]+)", texto)
    csrf = _regex_valor(r"CSLL\s*\(R\$\)\s*\n?([\d.,]+)", texto)
    inss = _regex_valor(r"INSS\s*\(R\$\)\s*\n?([\d.,]+)", texto)

    aliquota = _regex_str(r"Al[ií]q\.\s*\(%\)\s*\n?([\d,]+)", texto)
    if not aliquota:
        aliquota = _regex_str(r"([\d,]+)\s*%\s*\n?Al[ií]quota", texto)
        if not aliquota:
            aliquota = _regex_str(r"([\d.,]+)\s*\n?Al[ií]q", texto)
    if aliquota:
        aliquota = aliquota.replace(".", ",")
        if "," not in aliquota:
            aliquota += ",00"
        aliquota = f"{aliquota}%"

    local_m = re.search(r"Local\s+da\s+presta[cç][aã]o\s+do\s+servi[cç]o\s*\n([A-Za-zÀ-ú\s]+?)\s*/\s*([A-Z]{2})", texto, re.IGNORECASE)
    cidade = local_m.group(1).strip().title() if local_m else ""
    uf = local_m.group(2).upper() if local_m else ""

    # Reset de falsos positivos
    if valor_servico != "0,00":
        if _valores_iguais(ir, valor_servico): ir = "0,00"
        if _valores_iguais(pis, valor_servico): pis = "0,00"
        if _valores_iguais(cofins, valor_servico): cofins = "0,00"
        if _valores_iguais(csrf, valor_servico): csrf = "0,00"
        if _valores_iguais(inss, valor_servico): inss = "0,00"

    return NotaFiscalExtraida(
        arquivo_pdf=arquivo_pdf,
        prefeitura="Prefeitura Municipal de João Pessoa",
        cnpj_prestador=cnpj,
        numero_nf=numero,
        data_emissao=data,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=desc_servico,
        uf_local_prestacao=uf,
        cidade_local_prestacao=cidade,
        natureza_operacao=natureza,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        valor_deducoes="0,00",
        descontos_incondicionados="0,00",
        descontos_condicionados="0,00",
        outras_retencoes="0,00",
        ir=ir,
        pis_nao_retido=pis,
        cofins_nao_retido=cofins,
        csrf=csrf,
        inss=inss,
        aliquota=aliquota,
    )


# ---------------------------------------------------------------------------
# Função central de roteamento
# ---------------------------------------------------------------------------

def extrair_prefeitura(caminho_pdf: Path, nome_prefeitura: str) -> Optional[NotaFiscalExtraida]:
    """
    Extrai os campos da NF do PDF `caminho_pdf` usando o parser da `nome_prefeitura`.
    Retorna None quando a nota deve ser ignorada (PDF sem texto, OCR necessário, etc.).
    """
    # Lê o texto bruto (UTF-8 via fitz)
    texto = _ler_texto_pdf_utf8(caminho_pdf)

    # Ignora PDFs sem camada de texto (escaneados / necessitam de OCR)
    if len(texto.strip()) < 50:
        print(f"[IGNORADO] {caminho_pdf.name}: PDF sem camada de texto (OCR necessário).")
        return None

    nome = nome_prefeitura.strip()

    # Roteamento por prefeitura
    if nome == "Governo do Distrito Federal":
        return _parse_abrasf(texto, nome, caminho_pdf.name)

    if nome == "MUNICIPIO DE PETROLINA":
        return _parse_petrolina(texto, caminho_pdf.name)

    if nome == "NACIONAL":
        return _parse_abrasf(texto, nome, caminho_pdf.name)

    if nome == "PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS":
        return _parse_abrasf(texto, nome, caminho_pdf.name)

    if nome == "Prefeitura Municipal de Goiânia - GO":
        return _parse_abrasf(texto, nome, caminho_pdf.name)

    if nome == "PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE SP":
        return _parse_sao_roque(texto, caminho_pdf.name)

    if nome == "PREFEITURA MUNICIPAL DE BARUERI":
        return _parse_barueri(texto, caminho_pdf.name)

    if nome == "PREFEITURA MUNICIPAL DE CAMPO GRANDE":
        return _parse_campo_grande(texto, caminho_pdf.name)

    if nome in ("PREFEITURA MUNICIPAL DE EUSÉBIO", "PREFEITURA MUNICIPAL DE MARACANAÚ"):
        return _parse_sigiss_ce(texto, nome, caminho_pdf.name)

    if nome == "PREFEITURA MUNICIPAL DE FORTALEZA":
        return _parse_fortaleza(texto, caminho_pdf)

    if nome == "Prefeitura Municipal de João Pessoa":
        return _parse_joao_pessoa(texto, caminho_pdf.name)

    print(f"[IGNORADO] {caminho_pdf.name}: prefeitura '{nome}' não possui parser cadastrado.")
    return None


# ---------------------------------------------------------------------------
# Processamento em lote de uma pasta
# ---------------------------------------------------------------------------

def processar_pasta_prefeitura(
    pasta: Path,
    nome_prefeitura: str,
    callback_log=None,
    callback_progresso=None,
) -> tuple[list[NotaFiscalExtraida], Path | None]:
    """
    Processa todos os PDFs de `pasta` utilizando o parser de `nome_prefeitura`.

    Regras de movimentação:
    - PDFs extraídos com sucesso → movidos para `pasta/concluidas/`
    - PDFs com falha ou ignorados → permanecem na pasta original

    Retorna uma tupla (lista_de_registros_extraídos, caminho_xlsx_gerado).
    O XLSX é salvo em `pasta/nf_compilado.xlsx`.
    """
    def log(msg: str) -> None:
        print(msg)
        if callback_log:
            callback_log(msg)

    pdfs = sorted(pasta.glob("*.pdf"))
    # Ignora PDFs já movidos para a subpasta 'concluidas'
    pdfs = [p for p in pdfs if p.parent == pasta]

    if not pdfs:
        log(f"[{nome_prefeitura}] Nenhum PDF encontrado em: {pasta}")
        return [], None

    log(f"[{nome_prefeitura}] Encontrados {len(pdfs)} PDFs. Iniciando extração...")

    registros: list[NotaFiscalExtraida] = []
    total = len(pdfs)

    for indice, pdf in enumerate(pdfs, start=1):
        log(f"[{indice}/{total}] Processando: {pdf.name}")
        try:
            nota = extrair_prefeitura(pdf, nome_prefeitura)
        except Exception as e:
            log(f"[ERRO] {pdf.name}: {e}")
            nota = None

        if nota is not None:
            registros.append(nota)
            # Move o PDF com sucesso para a subpasta 'concluidas'
            _mover_para_concluidas(pdf)
            log(f"[OK] {pdf.name} extraído e movido para 'concluidas'.")
        else:
            log(f"[IGNORADO] {pdf.name}: não foi possível extrair dados.")

        if callback_progresso:
            callback_progresso(indice, total)

    if registros:
        caminho_xlsx = pasta / "nf_compilado.xlsx"
        log(f"[{nome_prefeitura}] Gerando planilha Excel consolidada... Aguarde.")
        gerar_xlsx(registros, caminho_xlsx)
        log(f"[{nome_prefeitura}] Planilha gerada: {caminho_xlsx}")
        return registros, caminho_xlsx
    else:
        log(f"[{nome_prefeitura}] Nenhuma nota extraída com sucesso.")
        return [], None


def identificar_nome_prefeitura(caminho_pdf: Path) -> Optional[str]:
    """
    Identifica dinamicamente a prefeitura correspondente a um PDF NFS-e sem usar IA.
    Retorna o nome canônico ou estruturado da prefeitura, ou None caso não identifique.
    """
    texto = _ler_texto_pdf_utf8(caminho_pdf)
    if not texto or len(texto.strip()) < 10:
        nome_arquivo_upper = caminho_pdf.name.upper()
        if "MACEIO" in nome_arquivo_upper or "MACEIÓ" in nome_arquivo_upper:
            return "PREFEITURA MUNICIPAL DE MACEIÓ"
        return None

    # Se tiver muitos caracteres de controle garbled (característico de São Paulo)
    control_chars_count = len([c for c in texto if ord(c) < 32 and c not in "\r\n\t"])
    if control_chars_count > 100:
        return "PREFEITURA DO MUNICÍPIO DE SÃO PAULO"

    from extrair_nf_pdfs import texto_para_busca
    texto_busca = texto_para_busca(texto)

    # 1. Nacional
    if "AMBIENTE NACIONAL" in texto_busca or "NFS-E NACIONAL" in texto_busca or "DANFSE NACIONAL" in texto_busca or "DANFESE" in texto_busca or "PORTAL NACIONAL" in texto_busca:
        return "NACIONAL"

    # 2. Distrito Federal
    if "DISTRITO FEDERAL" in texto_busca and ("SECRETARIA DE ESTADO" in texto_busca or "GDF" in texto_busca):
        return "Governo do Distrito Federal"

    # 3. Varredura por expressões regulares estruturadas
    padroes_pref = [
        r"PREFEITURA\s+MUNICIPAL\s+DE\s+([A-ZÀ-Úa-zà-ú'\-\s]{3,60})",
        r"PREFEITURA\s+DO\s+MUNICIPIO\s+DE\s+([A-ZÀ-Úa-zà-ú'\-\s]{3,60})",
        r"PREFEITURA\s+DA\s+EST[AÁ]NCIA\s+(?:TUR[IÍ]STICA\s+)?DE\s+([A-ZÀ-Úa-zà-ú'\-\s]{3,60})",
        r"MUNIC[IÍ]PIO\s+DE\s+([A-ZÀ-Úa-zà-ú'\-\s]{3,60})",
    ]

    for padrao in padroes_pref:
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            nome_cidade = m.group(1).strip()
            nome_cidade = nome_cidade.split("\n")[0].strip()
            nome_cidade = " ".join(nome_cidade.split())
            if len(nome_cidade) > 3 and not any(k in nome_cidade.upper() for k in ["CNPJ", "TELEFONE", "TOMADOR", "PRESTADOR", "ENDERECO", "CLIENTE"]):
                nome_cidade_upper = nome_cidade.upper()
                # Remove acentos para comparação robusta
                import unicodedata
                nome_busca = "".join(c for c in unicodedata.normalize("NFD", nome_cidade_upper) if unicodedata.category(c) != "Mn")
                if "PETROLINA" in nome_busca:
                    return "MUNICIPIO DE PETROLINA"
                if "FORTALEZA" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE FORTALEZA"
                if "BARUERI" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE BARUERI"
                if "CAMPO GRANDE" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE CAMPO GRANDE"
                if "EUSEBIO" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE EUSÉBIO"
                if "MARACANAU" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE MARACANAÚ"
                if "TEIXEIRA DE FREITAS" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS"
                if "GOIANIA" in nome_busca:
                    return "Prefeitura Municipal de Goiânia - GO"
                if "JOAO PESSOA" in nome_busca:
                    return "Prefeitura Municipal de João Pessoa"
                if "SAO ROQUE" in nome_busca:
                    return "PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE SP"
                if "SAO PAULO" in nome_busca:
                    return "PREFEITURA DO MUNICÍPIO DE SÃO PAULO"
                if "MACEIO" in nome_busca:
                    return "PREFEITURA MUNICIPAL DE MACEIÓ"
                
                # Para outras prefeituras dinâmicas
                if "MUNICIPAL DE" in m.group(0).upper():
                    return f"PREFEITURA MUNICIPAL DE {nome_cidade_upper}"
                elif "DO MUNICIPIO DE" in m.group(0).upper() or "DO MUNICÍPIO DE" in m.group(0).upper():
                    return f"PREFEITURA DO MUNICÍPIO DE {nome_cidade_upper}"
                else:
                    return f"PREFEITURA MUNICIPAL DE {nome_cidade_upper}"

    # 4. Fallbacks baseados em palavras-chave se as regexes estruturadas falharem
    if "FORTALEZA" in texto_busca and "SECRETARIA MUNICIPAL DAS FINANCAS" in texto_busca:
        return "PREFEITURA MUNICIPAL DE FORTALEZA"
    if "JOAO PESSOA" in texto_busca:
        return "Prefeitura Municipal de João Pessoa"
    if "PETROLINA" in texto_busca:
        return "MUNICIPIO DE PETROLINA"
    if "BARUERI" in texto_busca:
        return "PREFEITURA MUNICIPAL DE BARUERI"
    if "MACEIO" in texto_busca:
        return "PREFEITURA MUNICIPAL DE MACEIÓ"
    if "CAMPO GRANDE" in texto_busca and ("AGILI" in texto_busca or "PREFCAMPOGRANDE" in texto_busca):
        return "PREFEITURA MUNICIPAL DE CAMPO GRANDE"

    return None


def organizar_pasta_pdfs(
    pasta_origem: Path,
    callback_log=None,
    callback_progresso=None
) -> tuple[int, int]:
    """
    Organiza todos os PDFs da pasta_origem movendo-os para subpastas
    nomeadas com base na prefeitura identificada em cada PDF.
    Notas não identificadas não são movidas.
    """
    def log(msg: str) -> None:
        print(msg)
        if callback_log:
            callback_log(msg)

    pdfs = sorted(pasta_origem.glob("*.pdf"))
    # Ignora subpastas, analisa apenas arquivos na raiz
    pdfs = [p for p in pdfs if p.parent == pasta_origem]

    if not pdfs:
        log(f"Nenhum PDF encontrado na pasta: {pasta_origem}")
        return 0, 0

    log(f"Iniciando organização de {len(pdfs)} PDFs...")
    sucessos = 0
    ignorados = 0
    total = len(pdfs)

    for indice, pdf in enumerate(pdfs, start=1):
        try:
            nome_pref = identificar_nome_prefeitura(pdf)
            if nome_pref:
                # Cria a subpasta da prefeitura correspondente na própria pasta de origem
                pasta_destino = pasta_origem / nome_pref
                pasta_destino.mkdir(exist_ok=True)
                # Move o PDF
                destino_arquivo = pasta_destino / pdf.name
                shutil.move(str(pdf), str(destino_arquivo))
                log(f"[ORGANIZADO] {pdf.name} -> {nome_pref}/")
                sucessos += 1
            else:
                log(f"[IGNORADO] {pdf.name}: não foi possível identificar a prefeitura.")
                ignorados += 1
        except Exception as e:
            log(f"[ERRO] Falha ao processar {pdf.name}: {e}")
            ignorados += 1

        if callback_progresso:
            callback_progresso(indice, total)

    log(f"Organização concluída! Organizados: {sucessos}, Não movidos: {ignorados}")
    return sucessos, ignorados
