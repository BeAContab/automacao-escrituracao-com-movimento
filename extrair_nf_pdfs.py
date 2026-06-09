from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - fallback local
    fitz = None

from PyPDF2 import PdfReader

from iss_fortaleza_automacao import (
    executar_automacao_iss,
    interpretar_competencia,
    solicitar_competencia,
)
from cnae_final import enriquecer_registros_cnae_final
from gemini_extracao import extrair_campos_gemini
from tratamento_erros import registrar_erro, registrar_evento_execucao

VALOR_BR_PATTERN = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")
CNPJ_PATTERN = re.compile(
    r"\d{2}\s*\.?\s*\d{3}\s*\.?\s*\d{3}\s*/\s*\d{4}\s*-\s*\d{2}"
)

PADROES_CNPJ_PRESTADOR = [
    r"EMITENTE\s+PRESTADOR\s+DO\s+SERVICO.*?CPF\s*/\s*CNPJ\s*/\s*NIF\s*([0-9.\-/\s]+?)(?=\s*NOME|\s*ENDERECO|\s*SERVICO)",
    r"IDENTIFICACAO\s+DO\s+PRESTADOR.*?CNPJ\s*/\s*CPF\s*/\s*NIF:\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*TELEFONE|\s*NOME)",
    r"PRESTADOR\s+DE\s+SERVICOS?.*?CNPJ\s*/\s*CPF\s*:?\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*TELEFONE|\s*NOME|\s*RAZAO)",
    r"PRESTADOR\s+DO\s+SERVICO.*?CNPJ\s*:?\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*TELEFONE|\s*NOME|\s*RAZAO)",
    r"CPF\s*/\s*CNPJ\s*/\s*NIF\s*:?\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*NO\s*DA\s*NOTA|\s*NATUREZA|\s*NOME)",
    r"CNPJ\s*/\s*CPF\s*/\s*NIF\s*:?\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*TELEFONE|\s*NOME)",
    r"CNPJ\s*/\s*CPF\s*([0-9.\-/\s]+?)(?=\s*INSCRICAO|\s*TELEFONE|\s*NOME)",
    r"CNPJ\s*[:\-]\s*([0-9.\-/\s]+?)(?=\s*AGENCIA|\s*ENDERECO|\s*INSCRICAO|\s*RAZAO\s*SOCIAL)",
]

PADROES_NUMERO_NF = [
    r"NUMERO\s*DO\s*DOCUMENTO\s*(\d+)",
    r"NUMERO\s*DA\s*NFS-\s*E\s*(\d+)",
    r"NUMERO\s*DA\s*NFS\s*E\s*(\d+)",
    r"NFS-\s*E\s*N[Oº]?\s*(\d+)",
    r"NOTA\s*FISCAL\s*DE\s*SERVICOS\s*ELETRONICA\s*N[Oº]?\s*(\d+)",
    r"NUMERO\s*DA\s*NOTA\s*(\d+)",
    r"NO\s*DA\s*NOTA\s*FISCAL\s*(\d+)",
    r"NUMERO\s*DA\s*NOTA\s*FISCAL\s*(\d+)",
    r"NOTA\s*N[ÂºO]\s*(\d+)",
]

PADROES_DATA_EMISSAO = [
    r"DATA\s*DE\s*GERACAO\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
    r"DATA\s*FATO\s*GERADOR\s*(\d{2}/\d{2}/\d{4})",
    r"DATA\s*DA\s*EMISSAO:\s*(\d{2}/\d{2}/\d{4})",
    r"DATA\s*E\s*HORA\s*DA\s*EMISSAO\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
    r"DATA\s*E\s*HORA\s*DA\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
    r"EMISSAO\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"EMITIDA\s*EM\s*(\d{2}/\d{2}/\d{4})",
    r"COMPETENCIA\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
    r"COMPETENCIA\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
    r"DATA\s*DE\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
    r"DT\.\s*DE\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
]

MARCADORES_FIM_BLOCO = [
    "TOMADOR",
    "DADOS DO TOMADOR",
    "CALCULO",
    "CÁLCULO",
    "TRIBUTACAO",
    "TRIBUTAÇÃO",
    "VALOR",
    "ISS",
    "CODIGO",
    "CÓDIGO",
    "PAIS",
    "PAÍS",
    "NATUREZA",
]


@dataclass
class NotaFiscalExtraida:
    arquivo_pdf: str
    cnpj_prestador: str = ""
    numero_nf: str = ""
    data_emissao: str = ""
    id_cnae: str = ""
    desc_cnae: str = ""
    descricao_servico: str = ""
    uf_local_prestacao: str = ""
    cidade_local_prestacao: str = ""
    natureza_operacao: str = ""
    iss_retido: str = ""
    valor_servico: str = ""
    aliquota: str = ""
    id_cnae_final: str = ""
    desc_cnae_final: str = ""

    def como_linha(self) -> list[str]:
        return [
            self.arquivo_pdf,
            self.cnpj_prestador,
            self.numero_nf,
            self.data_emissao,
            self.id_cnae,
            self.desc_cnae,
            self.descricao_servico,
            self.uf_local_prestacao,
            self.cidade_local_prestacao,
            self.natureza_operacao,
            self.iss_retido,
            self.valor_servico,
            self.aliquota,
            self.id_cnae_final,
            self.desc_cnae_final,
        ]

    def campos_vazios(self) -> list[str]:
        """Lista as colunas que não foram preenchidas para esta NF."""
        campos = {
            "CNPJ_PRESTADOR": self.cnpj_prestador,
            "NUMERO_NF": self.numero_nf,
            "DATA_EMISSAO": self.data_emissao,
            "ID_CNAE": self.id_cnae,
            "DESC_CNAE": self.desc_cnae,
            "DESCRICAO_SERVICO": self.descricao_servico,
            "UF_LOCAL_PRESTACAO": self.uf_local_prestacao,
            "CIDADE_LOCAL_PRESTACAO": self.cidade_local_prestacao,
            "NATUREZA_OPERACAO": self.natureza_operacao,
            "ISS_RETIDO": self.iss_retido,
            "VALOR_SERVICO": self.valor_servico,
            "ALIQUOTA": self.aliquota,
        }
        return [nome for nome, valor in campos.items() if not str(valor).strip()]

    def resumo_campos(self) -> str:
        """Devolve um resumo curto para facilitar o diagnóstico no log."""
        return (
            f"CNPJ={self.cnpj_prestador or 'vazio'} | "
            f"NF={self.numero_nf or 'vazio'} | "
            f"Data={self.data_emissao or 'vazio'} | "
            f"Valor={self.valor_servico or 'vazio'} | "
            f"Alíquota={self.aliquota or 'vazio'}"
        )


def normalizar_texto(texto: str) -> str:
    """Consolida espaços para facilitar a extração por expressão regular."""
    return " ".join(texto.split())


def texto_para_busca(texto: str) -> str:
    """Remove acentos mantendo o tamanho do texto para preservar os índices.

    Os PDFs das três prefeituras usam rótulos parecidos, mas com acentos,
    espaços e quebras diferentes. Esta versão facilita a busca sem perder a
    possibilidade de recortar o conteúdo original.
    """
    equivalencias = {"º": "O", "ª": "A", "–": "-", "—": "-", "\u00a0": " "}
    caracteres: list[str] = []

    for caractere in texto:
        caractere = equivalencias.get(caractere, caractere)
        decomposicao = unicodedata.normalize("NFKD", caractere)
        base = "".join(c for c in decomposicao if not unicodedata.combining(c))
        caracteres.append((base[:1] or " ").upper())

    return "".join(caracteres)


def ler_texto_pdf(caminho_pdf: Path) -> str:
    """Lê o texto de todas as páginas do PDF."""
    if fitz is not None:
        documento = fitz.open(str(caminho_pdf))
        partes = [pagina.get_text("text") or "" for pagina in documento]
        return normalizar_texto("\n".join(partes))

    leitor = PdfReader(str(caminho_pdf))
    partes: list[str] = []
    for pagina in leitor.pages:
        partes.append(pagina.extract_text() or "")
    return normalizar_texto(" ".join(partes))


def buscar_primeiro_grupo(
    texto_original: str,
    texto_busca: str,
    padroes: list[str],
    grupo: int = 1,
) -> str:
    """Retorna no texto original o grupo encontrado no texto normalizado."""
    for padrao in padroes:
        encontrado = re.search(padrao, texto_busca, flags=re.IGNORECASE)
        if encontrado:
            return texto_original[encontrado.start(grupo) : encontrado.end(grupo)].strip()
    return ""


def pontuar_cnpj_por_contexto(busca: str, inicio: int, fim: int) -> int:
    """Pontua CNPJs pela proximidade com blocos de prestador.

    A regra evita depender de um layout específico: em muitos PDFs o primeiro
    CNPJ é o do prestador, mas a presença de PIX, banco ou tomador pode confundir
    a extração. A pontuação privilegia o contexto fiscal do emitente.
    """
    antes = busca[max(0, inicio - 400) : inicio]
    depois = busca[fim : min(len(busca), fim + 250)]
    contexto_curto = busca[max(0, inicio - 80) : min(len(busca), fim + 80)]
    pos_tomador = busca.find("TOMADOR")
    pontuacao = 0

    if any(marcador in antes for marcador in ["PRESTADOR", "EMITENTE", "FORNECEDOR"]):
        pontuacao += 8
    if any(marcador in contexto_curto for marcador in ["CNPJ / CPF", "CPF / CNPJ", "CNPJ:"]):
        pontuacao += 2
    if pos_tomador >= 0 and inicio < pos_tomador:
        pontuacao += 3
    if any(marcador in antes[-160:] for marcador in ["TOMADOR", "INTERMEDIARIO"]):
        pontuacao -= 6
    if any(marcador in contexto_curto for marcador in ["PIX", "CHAVE PIX", "BANCO", "AGENCIA", "CONTA"]):
        pontuacao -= 8
    if any(marcador in depois for marcador in ["RAZAO SOCIAL", "NOME", "INSCRICAO"]):
        pontuacao += 1

    return pontuacao


def extrair_cnpj_por_contexto(texto: str, busca: str) -> str:
    """Seleciona o CNPJ mais provável do prestador quando os rótulos variam."""
    candidatos: list[tuple[int, int, str]] = []
    for encontrado in CNPJ_PATTERN.finditer(busca):
        valor = texto[encontrado.start() : encontrado.end()]
        pontuacao = pontuar_cnpj_por_contexto(busca, encontrado.start(), encontrado.end())
        candidatos.append((pontuacao, encontrado.start(), valor))

    if not candidatos:
        return ""

    candidatos.sort(key=lambda item: (-item[0], item[1]))
    return limpar_cnpj(candidatos[0][2])


def extrair_cnpj_prestador(texto: str, busca: str) -> str:
    cnpj = buscar_primeiro_grupo(texto, busca, PADROES_CNPJ_PRESTADOR)
    if cnpj and len(re.sub(r"\D+", "", cnpj)) == 14:
        return limpar_cnpj(cnpj)
    return extrair_cnpj_por_contexto(texto, busca)


def extrair_numero_nf(texto: str, busca: str) -> str:
    trecho_antes = _trecho_antes_rotulo(
        texto,
        busca,
        ["NUMERO DA NFS E", "NUMERO DA NFS-E", "NUMERO DA NOTA FISCAL", "NUMERO DA NOTA"],
        janela=40,
    )
    encontrado_antes = re.search(r"\b(\d{1,12})\s*$", trecho_antes.strip())
    if encontrado_antes:
        return normalizar_numero_nf(encontrado_antes.group(1))

    numero = buscar_primeiro_grupo(
        texto,
        busca,
        [
            r"NUMERO\s*DA\s*NOTA\s*FISCAL\s*(\d+)",
            r"NUMERO\s*DA\s*NOTA\s*(\d+)",
            r"NUMERO\s*DA\s*NFS-\s*E\s*(\d+)",
            r"NUMERO\s*DA\s*NFS\s*E\s*(\d+)",
            r"NFS\s*E\s*N[OÂº]?\s*(\d+)",
            r"NFS-\s*E\s*N[OÂº]?\s*(\d+)",
            r"NOTA\s*FISCAL\s*DE\s*SERVICOS\s*ELETRONICA\s*N[OÂº]?\s*(\d+)",
            r"NO\s*DA\s*NOTA\s*FISCAL\s*(\d+)",
            r"NUMERO\s*DO\s*DOCUMENTO\s*(\d+)",
        ]
        + PADROES_NUMERO_NF,
    )
    if numero:
        return normalizar_numero_nf(numero)

    return ""


def extrair_data_emissao(texto: str, busca: str) -> str:
    data = buscar_primeiro_grupo(
        texto,
        busca,
        [
            r"DATA\s*DE\s*GERACAO\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*DE\s*GERACAO\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*FATO\s*GERADOR\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*E\s*HORA\s*DA\s*EMISSAO\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*E\s*HORA\s*DA\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*DA\s*EMISSAO:\s*(\d{2}/\d{2}/\d{4})",
            r"EMISSAO\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            r"EMITIDA\s*EM\s*(\d{2}/\d{2}/\d{4})",
            r"COMPETENCIA\s*DA\s*NFS-\s*E\s*(\d{2}/\d{2}/\d{4})",
            r"COMPETENCIA\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s*DE\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
            r"DT\.\s*DE\s*EMISSAO\s*(\d{2}/\d{2}/\d{4})",
        ]
        + PADROES_DATA_EMISSAO,
    )
    if data:
        return normalizar_data_br(data)

    # Fallback conservador: usa a primeira data próxima do cabeçalho de emissão.
    for encontrado in re.finditer(r"\d{2}/\d{2}/\d{4}", busca):
        contexto = busca[max(0, encontrado.start() - 120) : encontrado.start()]
        if any(marcador in contexto for marcador in ["EMISSAO", "GERACAO", "COMPETENCIA"]):
            return texto[encontrado.start() : encontrado.end()]

    return ""


def limpar_espacos(valor: str) -> str:
    return normalizar_texto(valor.strip())


def limpar_cnpj(cnpj: str) -> str:
    return re.sub(r"\s+", "", cnpj)


def normalizar_numero_nf(numero: str) -> str:
    """Remove zeros à esquerda sem alterar números já normalizados."""
    numero = re.sub(r"\D+", "", numero or "")
    if not numero:
        return ""
    return str(int(numero))


def normalizar_data_br(valor: str) -> str:
    """Retorna apenas a data DD/MM/AAAA, ignorando horário quando existir."""
    if not valor:
        return ""
    encontrado = re.search(r"\d{2}/\d{2}/\d{4}", valor)
    return encontrado.group(0) if encontrado else ""


def _trecho_apos_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 260) -> str:
    """Retorna o trecho logo após o primeiro rótulo encontrado."""

    for rotulo in rotulos:
        encontrado = re.search(texto_para_busca(rotulo), busca, flags=re.IGNORECASE)
        if encontrado:
            inicio = encontrado.end()
            return texto[inicio : min(len(texto), inicio + janela)]
    return ""


def _trecho_antes_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 160) -> str:
    """Retorna o trecho imediatamente anterior ao primeiro rótulo encontrado."""

    for rotulo in rotulos:
        encontrado = re.search(texto_para_busca(rotulo), busca, flags=re.IGNORECASE)
        if encontrado:
            fim = encontrado.start()
            return texto[max(0, fim - janela) : fim]
    return ""


def _capturar_numero_apos_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 260) -> str:
    trecho = _trecho_apos_rotulo(texto, busca, rotulos, janela=janela)
    if not trecho:
        return ""
    encontrado = re.search(r"\b0*\d{1,12}\b", trecho)
    return normalizar_numero_nf(encontrado.group(0)) if encontrado else ""


def _capturar_data_apos_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 260) -> str:
    trecho = _trecho_apos_rotulo(texto, busca, rotulos, janela=janela)
    if not trecho:
        return ""
    encontrado = re.search(r"\d{2}/\d{2}/\d{4}", trecho)
    return normalizar_data_br(encontrado.group(0)) if encontrado else ""


def _capturar_valor_apos_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 260) -> Decimal:
    trecho = _trecho_apos_rotulo(texto, busca, rotulos, janela=janela)
    if not trecho:
        return Decimal("0")
    encontrado = VALOR_BR_PATTERN.search(trecho)
    return converter_moeda_br_para_decimal(encontrado.group(0)) if encontrado else Decimal("0")


def _capturar_percentual_proximo_rotulo(texto: str, busca: str, rotulos: list[str], janela: int = 180) -> Decimal:
    trecho_antes = _trecho_antes_rotulo(texto, busca, rotulos, janela=janela)
    if trecho_antes:
        encontrado = re.search(r"(\d{1,2}(?:[.,]\d{1,4})?)\s*$", trecho_antes.strip())
        if encontrado:
            return converter_percentual_para_decimal(encontrado.group(1))

    trecho_depois = _trecho_apos_rotulo(texto, busca, rotulos, janela=janela)
    if trecho_depois:
        encontrado = re.search(r"(\d{1,2}(?:[.,]\d{1,4})?)", trecho_depois)
        if encontrado:
            return converter_percentual_para_decimal(encontrado.group(1))

    return Decimal("0")


def _ajustes_por_layout_especifico(texto: str, busca: str) -> dict[str, object]:
    """Aplica regras por família de NFS-e para layouts conhecidos do projeto."""

    ajustes: dict[str, object] = {}

    if "PREFEITURA DO MUNICIPIO DE SAO PAULO" in busca:
        numero_nf = _capturar_numero_apos_rotulo(
            texto,
            busca,
            ["NUMERO DA NOTA", "NUMERO DA NOTA FISCAL", "NUMERO DA NFS E", "NFS E"],
        )
        if numero_nf:
            ajustes["numero_nf"] = numero_nf

        data_emissao = _capturar_data_apos_rotulo(
            texto,
            busca,
            ["DATA E HORA DE EMISSAO", "DATA E HORA DE EMISSAO DA NFS E"],
        )
        if data_emissao:
            ajustes["data_emissao"] = data_emissao

        descricao = ""
        padrao_descricao = re.search(
            r"DISCRIMINACAO\s*DE\s*SERVICOS\s*(.+?)(?=\s*VALOR\s*TOTAL\s*DO\s*SERVICO|\s*CONTRIBUICAO\s*PREVIDENCIARIA|\s*OUTRAS\s*INFORMACOES|\s*$)",
            busca,
            flags=re.IGNORECASE,
        )
        if padrao_descricao:
            descricao = texto[padrao_descricao.start(1) : padrao_descricao.end(1)]
        if descricao:
            ajustes["descricao_servico"] = limpar_descricao_servico(descricao)

        padrao_cnae = re.search(
            r"CODIGO\s*DO\s*SERVICO\s*([0-9A-Z.\-/]+)\s*-\s*(.+?)(?=\s*VALOR\s*TOTAL\s*DAS\s*DEDUCOES|\s*VALOR\s*TOTAL\s*DO\s*SERVICO|\s*ALIQUOTA|\s*MUNICIPIO\s*DE\s*PRESTACAO|\s*$)",
            busca,
            flags=re.IGNORECASE,
        )
        if padrao_cnae:
            ajustes["id_cnae"] = limpar_espacos(texto[padrao_cnae.start(1) : padrao_cnae.end(1)])
            ajustes["desc_cnae"] = limpar_descricao_cnae(
                texto[padrao_cnae.start(2) : padrao_cnae.end(2)]
            )

        padrao_local = re.search(
            r"MUNICIPIO\s*DE\s*PRESTACAO\s*DO\s*SERVICO\s*([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)(?:\s*/\s*([A-Z]{2}))?(?=\s*$|\s*CHAVE|\s*VALOR|\s*OUTRAS\s*INFORMACOES)",
            texto,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_local:
            ajustes["cidade_local_prestacao"] = limpar_cidade_extraida(
                texto[padrao_local.start(1) : padrao_local.end(1)]
            )
            if padrao_local.lastindex and padrao_local.lastindex >= 2:
                ajustes["uf_local_prestacao"] = texto[padrao_local.start(2) : padrao_local.end(2)].strip().upper()
        elif "CAMPINA GRANDE" in busca:
            ajustes["cidade_local_prestacao"] = "Campina Grande"
            ajustes["uf_local_prestacao"] = "PB"

        valor_servico = _capturar_valor_apos_rotulo(
            texto,
            busca,
            ["VALOR TOTAL DO SERVICO", "VALOR TOTAL"],
        )
        if valor_servico > 0:
            ajustes["valor_servico"] = valor_servico

        aliquota = _capturar_percentual_proximo_rotulo(
            texto,
            busca,
            ["ALIQ. ISSQN", "ALIQ ISSQN", "ALIQUOTA"],
        )
        if aliquota <= 0:
            encontrado_aliquota = re.search(
                r"ALIQUOTA\s*\(%\).*?(\d{1,2}(?:[.,]\d{1,4})?)\s*%",
                busca,
                flags=re.IGNORECASE,
            )
            if encontrado_aliquota:
                aliquota = converter_percentual_para_decimal(
                    texto[encontrado_aliquota.start(1) : encontrado_aliquota.end(1)]
                )
        if aliquota > 0:
            ajustes["aliquota"] = aliquota

        cidade_atual = texto_para_busca(ajustes.get("cidade_local_prestacao", ""))
        if not cidade_atual or "NUMERO INSCRICAO" in cidade_atual:
            if "CAMPINA GRANDE" in busca:
                ajustes["cidade_local_prestacao"] = "Campina Grande"
                ajustes["uf_local_prestacao"] = "PB"

        padrao_prestador = re.search(
            r"PRESTADOR\s*DE\s*SERVICOS.*?CPF\s*/\s*CNPJ:\s*([0-9.\-/]+).*?NOME\s*/\s*RAZAO\s*SOCIAL:",
            busca,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_prestador:
            ajustes["cnpj_prestador"] = limpar_cnpj(
                texto[padrao_prestador.start(1) : padrao_prestador.end(1)]
            )

        padrao_municipio = re.search(
            r"MUNICIPIO\s*(?::|\s*DE\s*PRESTACAO\s*DO\s*SERVICO\s*[:\-]?)\s*([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)(?:\s*UF\s*[:\-]?\s*([A-Z]{2}))?(?=\s*$|\s*CHAVE|\s*VALOR|\s*OUTRAS\s*INFORMACOES)",
            busca,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_municipio:
            ajustes["cidade_local_prestacao"] = limpar_cidade_extraida(
                texto[padrao_municipio.start(1) : padrao_municipio.end(1)]
            )
            if padrao_municipio.lastindex and padrao_municipio.lastindex >= 2:
                ajustes["uf_local_prestacao"] = texto[
                    padrao_municipio.start(2) : padrao_municipio.end(2)
                ].strip().upper()
        elif "SÃO PAULO" in busca or "SAO PAULO" in busca:
            ajustes["cidade_local_prestacao"] = "São Paulo"
            ajustes["uf_local_prestacao"] = "SP"

        valor_servico = _capturar_valor_apos_rotulo(
            texto,
            busca,
            ["VALOR TOTAL DO SERVICO", "VALOR TOTAL DO SERVICO =", "VALOR TOTAL"],
        )
        if valor_servico > 0:
            ajustes["valor_servico"] = valor_servico

        aliquota = _capturar_percentual_proximo_rotulo(
            texto,
            busca,
            ["ALIQ. ISSQN", "ALIQ ISSQN", "ALIQUOTA", "ALIQUOTA (%)"],
        )
        if aliquota > 0:
            ajustes["aliquota"] = aliquota

        if "CAMPINA GRANDE" in busca:
            ajustes["cidade_local_prestacao"] = "Campina Grande"
            ajustes["uf_local_prestacao"] = "PB"

    elif "PREFEITURA MUNICIPAL DE FORTALEZA" in busca:
        numero_nf = _capturar_numero_apos_rotulo(
            texto,
            busca,
            ["NUMERO DA NFS E", "NUMERO DA NFS-E", "NFS E"],
        )
        if numero_nf:
            ajustes["numero_nf"] = numero_nf

        data_emissao = _capturar_data_apos_rotulo(
            texto,
            busca,
            ["DATA E HORA DA EMISSAO", "DATA E HORA DA EMISSAO DA NFS E"],
        )
        if data_emissao:
            ajustes["data_emissao"] = data_emissao

        padrao_cnae = re.search(
            r"C[ÁA]LCULO\s*DO\s*ISSQN\s*([0-9.\s/]+)\s*-\s*(.+?)(?=\s*E-MAIL|\s*VALOR\s*DOS\s*SERVICOS|\s*DADOS\s*DO\s*TOMADOR|\s*$)",
            texto,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_cnae:
            ajustes["id_cnae"] = limpar_espacos(texto[padrao_cnae.start(1) : padrao_cnae.end(1)])
            ajustes["desc_cnae"] = limpar_descricao_cnae(
                texto[padrao_cnae.start(2) : padrao_cnae.end(2)]
            )

        padrao_descricao = re.search(
            r"DISCRIMINAC[ÃA]O\s*DOS\s*SERVICOS\s*(.+?)(?=\s*DADOS\s*BANCARIOS|\s*C[ÁA]LCULO\s*DO\s*ISSQN|\s*TRIBUTA[ÇC][ÃA]O\s*NACIONAL|\s*TRIBUTA[ÇC][ÃA]O\s*MUNICIPAL|\s*$)",
            texto,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_descricao:
            ajustes["descricao_servico"] = limpar_descricao_servico(
                texto[padrao_descricao.start(1) : padrao_descricao.end(1)]
            )

        padrao_local = re.search(
            r"LOCAL\s*DA\s*PRESTACAO\s*[:\s]*([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)(?:\s*[-/]\s*([A-Z]{2}))?(?=\s*CHAVE|\s*TRIBUTACAO|\s*PIS|\s*$)",
            busca,
            flags=re.IGNORECASE,
        )
        if padrao_local:
            ajustes["cidade_local_prestacao"] = limpar_cidade_extraida(
                texto[padrao_local.start(1) : padrao_local.end(1)]
            )
            if padrao_local.lastindex and padrao_local.lastindex >= 2:
                ajustes["uf_local_prestacao"] = texto[padrao_local.start(2) : padrao_local.end(2)].strip().upper()

        valor_servico = _capturar_valor_apos_rotulo(
            texto,
            busca,
            ["VALOR DOS SERVICOS", "VALOR DO SERVICO"],
        )
        if valor_servico > 0:
            ajustes["valor_servico"] = valor_servico

        aliquota = _capturar_percentual_proximo_rotulo(
            texto,
            busca,
            ["ALIQUIDADE", "ALIQUOTA", "ALIQUOTA", "ALIQUOTA %", "ALIQ ISSQN"],
        )
        if aliquota <= 0:
            encontrado_aliquota = re.search(
                r"CPF\s*/\s*CNPJ\s+(\d{1,2}(?:[.,]\d{1,4})?)\s+VALOR\s*DOS\s*SERVICOS",
                busca,
                flags=re.IGNORECASE,
            )
            if encontrado_aliquota:
                aliquota = converter_percentual_para_decimal(
                    texto[encontrado_aliquota.start(1) : encontrado_aliquota.end(1)]
                )
        if aliquota > 0:
            ajustes["aliquota"] = aliquota

    elif "MUNICIPIO DE PETROLINA" in busca:
        numero_nf = _capturar_numero_apos_rotulo(
            texto,
            busca,
            ["NO DA NOTA FISCAL", "NUMERO DA NOTA FISCAL", "NUMERO DA NOTA", "NUMERO DA NFS E"],
        )
        if numero_nf:
            ajustes["numero_nf"] = numero_nf

        data_emissao = _capturar_data_apos_rotulo(
            texto,
            busca,
            ["DATA FATO GERADOR", "EMITIDO EM", "DATA DE EMISSAO"],
        )
        if data_emissao:
            ajustes["data_emissao"] = data_emissao

        padrao_local = re.search(
            r"LOCAL\s*DE\s*PRESTACAO.*?(\d{7})\s*-\s*([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)\s*-\s*([A-Z]{2})",
            busca,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_local:
            ajustes["cidade_local_prestacao"] = limpar_cidade_extraida(
                texto[padrao_local.start(2) : padrao_local.end(2)]
            )
            ajustes["uf_local_prestacao"] = texto[padrao_local.start(3) : padrao_local.end(3)].strip().upper()

        padrao_descricao = re.search(
            r"DISCRIMINAC[ÃA]O\s*DOS\s*SERVICOS\s*(.+?)(?=\s*VALOR\s*SERVICO|\s*VALOR\s*DOS\s*SERVICOS|\s*DADOS\s*BANCARIOS|\s*$)",
            busca,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if padrao_descricao:
            ajustes["descricao_servico"] = limpar_descricao_servico(
                texto[padrao_descricao.start(1) : padrao_descricao.end(1)]
            )

    return ajustes


def _somente_digitos(texto: str) -> str:
    return re.sub(r"\D+", "", texto or "")


def _campo_presente_no_texto(valor: str, texto: str, tipo: str) -> bool:
    """Confere se o valor local realmente aparece no texto bruto do PDF."""

    if not valor:
        return False

    if tipo in {"numero_nf", "cnpj_prestador", "valor_servico", "aliquota", "data_emissao"}:
        return _somente_digitos(valor) in _somente_digitos(texto)

    if tipo in {"cidade_local_prestacao", "desc_cnae", "descricao_servico", "uf_local_prestacao", "id_cnae"}:
        return texto_para_busca(valor) in texto_para_busca(texto)

    return bool(valor.strip())


def _mesclar_campo_textual(valor_local: str, valor_gemini: str, texto: str, tipo: str) -> str:
    """Usa o Gemini apenas quando o valor local esta vazio ou suspeito."""

    valor_gemini = (valor_gemini or "").strip()
    if not valor_gemini:
        return valor_local

    if not valor_local.strip():
        return valor_gemini

    if not _campo_presente_no_texto(valor_local, texto, tipo):
        return valor_gemini

    return valor_local


def _mesclar_campo_decimal(valor_local: Decimal, valor_gemini: str, texto: str, tipo: str) -> Decimal:
    """Converte e substitui valores monetarios ou percentuais quando o local falhou."""

    valor_gemini = (valor_gemini or "").strip()
    if not valor_gemini:
        return valor_local

    valor_local_texto = (
        formatar_moeda_br(valor_local) if tipo == "valor_servico" else formatar_percentual(valor_local)
    ) if valor_local else ""

    if not valor_local or not _campo_presente_no_texto(valor_local_texto, texto, tipo):
        if tipo == "valor_servico":
            return converter_moeda_br_para_decimal(valor_gemini)
        return converter_percentual_para_decimal(valor_gemini)

    return valor_local


def limpar_descricao_cnae(descricao: str) -> str:
    descricao = cortar_por_marcadores(
        descricao,
        [
            "CODIGO DE TRIBUTACAO MUNICIPAL",
            "CODIGO DE TRIBUTACAO NACIONAL",
            "CODIGO NBS",
            "LOCAL DA PRESTACAO",
            "LOCAL DE PRESTACAO",
            "VL. DO SERVICO",
            "VALOR DO SERVICO",
            "DESCRICAO DO SERVICO",
            "DESCRICAO DOS SERVICOS",
            "IMPOSTO SOBRE SERVICO",
            "NBS:",
        ],
    )
    return limpar_espacos(descricao)


def cortar_por_marcadores(texto: str, marcadores: list[str]) -> str:
    """Corta um trecho no primeiro marcador fiscal posterior encontrado."""

    busca = texto_para_busca(texto)
    fim = len(texto)
    for marcador in marcadores:
        pos = busca.find(texto_para_busca(marcador))
        if pos >= 0:
            fim = min(fim, pos)
    return texto[:fim]


def limpar_descricao_servico(descricao: str) -> str:
    return limpar_espacos(
        cortar_por_marcadores(
            descricao,
            [
                "VALOR DOS SERVICOS",
                "VALOR DO SERVICO",
                "VALOR TOTAL",
                "CALCULO DO ISS",
                "DADOS DO TOMADOR",
                "OUTRAS INFORMACOES",
                "TRIBUTOS FEDERAIS",
            ],
        )
    )


def limpar_cidade_extraida(cidade: str) -> str:
    cidade = limpar_espacos(cidade)
    cidade = re.sub(r"\s*[-/]\s*[A-Z]{2}$", "", cidade, flags=re.IGNORECASE)
    return cidade.title() if cidade.isupper() else cidade


def converter_moeda_br_para_decimal(valor: str) -> Decimal:
    valor = str(valor or "").strip()
    if not valor:
        return Decimal("0")
    valor = re.sub(r"[^\d,.-]", "", valor)
    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")
    try:
        return Decimal(valor)
    except Exception:
        return Decimal("0")


def converter_percentual_para_decimal(valor: str) -> Decimal:
    return converter_moeda_br_para_decimal(valor)


def formatar_moeda_br(valor: Decimal) -> str:
    valor = Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_percentual(valor: Decimal) -> str:
    valor = Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _extrair_por_regex_original(texto: str, busca: str, padroes: list[str]) -> str:
    return buscar_primeiro_grupo(texto, busca, padroes)


def extrair_id_desc_cnae(texto: str, busca: str) -> tuple[str, str]:
    padroes = [
        r"CODIGO\s*DO\s*SERVICO\s*([\d./\-\s]+)\s*[-:]\s*(.+?)(?=\s*DISCRIMINACAO|\s*DESCRICAO|\s*VALOR|\s*LOCAL|\s*DADOS)",
        r"CALCULO\s*DO\s*ISSQN\s*([\d./\-\s]+)\s*[-:]\s*(.+?)(?=\s*E-MAIL|\s*VALOR|\s*CODIGO\s*ART|\s*CNPJ)",
        r"SERVICO\s*:\s*([\d./\-\s]+)\s*[-:]\s*(.+?)(?=\s*VALOR|\s*ALIQUOTA|\s*ISS)",
    ]
    for padrao in padroes:
        encontrado = re.search(padrao, busca, flags=re.IGNORECASE | re.DOTALL)
        if encontrado:
            codigo = limpar_espacos(texto[encontrado.start(1) : encontrado.end(1)])
            descricao = limpar_descricao_cnae(texto[encontrado.start(2) : encontrado.end(2)])
            return codigo, descricao
    return "", ""


def extrair_descricao_servico(texto: str, busca: str) -> str:
    padroes = [
        r"DISCRIMINACAO\s*DOS\s*SERVICOS\s*(.+?)(?=\s*VALOR\s*TOTAL|\s*VALOR\s*DOS\s*SERVICOS|\s*CALCULO|\s*OUTRAS\s*INFORMACOES|\s*$)",
        r"DESCRICAO\s*DOS\s*SERVICOS\s*(.+?)(?=\s*VALOR\s*TOTAL|\s*VALOR\s*DOS\s*SERVICOS|\s*CALCULO|\s*OUTRAS\s*INFORMACOES|\s*$)",
        r"DESCRICAO\s*DO\s*SERVICO\s*(.+?)(?=\s*VALOR\s*TOTAL|\s*VALOR\s*DOS\s*SERVICOS|\s*CALCULO|\s*OUTRAS\s*INFORMACOES|\s*$)",
    ]
    for padrao in padroes:
        encontrado = re.search(padrao, busca, flags=re.IGNORECASE | re.DOTALL)
        if encontrado:
            return limpar_descricao_servico(texto[encontrado.start(1) : encontrado.end(1)])
    return ""


def extrair_cidade_uf_local_prestacao(texto: str, busca: str) -> tuple[str, str]:
    padroes = [
        r"MUNICIPIO\s*DE\s*PRESTACAO\s*DO\s*SERVICO\s*([A-ZÀ-Ü\s.'-]+?)\s*[-/]\s*([A-Z]{2})",
        r"LOCAL\s*DE\s*PRESTACAO.*?([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)\s*[-/]\s*([A-Z]{2})",
        r"PRESTADO\s*EM\s*([A-ZÀ-Ü][A-ZÀ-Ü\s.'-]+?)\s*[-/]\s*([A-Z]{2})",
    ]
    for padrao in padroes:
        encontrado = re.search(padrao, busca, flags=re.IGNORECASE | re.DOTALL)
        if encontrado:
            cidade = limpar_cidade_extraida(texto[encontrado.start(1) : encontrado.end(1)])
            uf = texto[encontrado.start(2) : encontrado.end(2)].strip().upper()
            return cidade, uf
    if "CAMPINA GRANDE" in busca:
        return "Campina Grande", "PB"
    return "", ""


def extrair_valor_servico(texto: str, busca: str) -> Decimal:
    padroes = [
        r"VALOR\s*TOTAL\s*DO\s*SERVICO\s*R\$?\s*([\d\.,]+)",
        r"VALOR\s*DOS\s*SERVICOS\s*R\$?\s*([\d\.,]+)",
        r"VALOR\s*DO\s*SERVICO\s*R\$?\s*([\d\.,]+)",
        r"VALOR\s*SERVICO\s*R\$?\s*([\d\.,]+)",
    ]
    return converter_moeda_br_para_decimal(_extrair_por_regex_original(texto, busca, padroes))


def extrair_aliquota(texto: str, busca: str) -> Decimal:
    padroes = [
        r"ALIQUOTA\s*\(%\)\s*([\d\.,]+)",
        r"ALIQUOTA\s*([\d\.,]+)\s*%",
        r"([\d\.,]+)\s*%\s*ISS",
        r"ALIQUOTA\s*([\d\.,]+)",
    ]
    return converter_percentual_para_decimal(_extrair_por_regex_original(texto, busca, padroes))


def extrair_valor_liquido(texto: str, busca: str) -> Decimal:
    padroes = [
        r"VALOR\s*LIQUIDO\s*DA\s*NFS-\s*E\s*R\$?\s*([\d\.,]+)",
        r"VALOR\s*LIQUIDO\s*DO\s*SERVICO\s*R\$?\s*([\d\.,]+)",
        r"VALOR\s*LIQUIDO\s*R\$?\s*([\d\.,]+)",
    ]
    return converter_moeda_br_para_decimal(_extrair_por_regex_original(texto, busca, padroes))


def determinar_natureza_operacao(cidade_local_prestacao: str) -> str:
    return (
        "Tributação no Município"
        if cidade_local_prestacao.strip().lower() == "fortaleza"
        else "Tributação Fora do Município"
    )


def extrair_iss_retido(texto: str, busca: str, valor_servico: Decimal, aliquota: Decimal) -> str:
    if re.search(r"O\s*ISS\s*DESTA\s*NFS\s*E\s*SERA\s*RETIDO\s*PELO\s*TOMADOR\s*DE\s*SERVICO", busca):
        return "SIM"
    if re.search(
        r"RETIDO\s*PELO\s*TOMADOR\s*DE\s*SERVICO|TIPO\s*DE\s*RETENCAO\s*RETIDO\s*PELO\s*TOMADOR|RETIDO\s*NA\s*FONTE|ISS\s*A\s*RETER.*?\bSIM\b",
        busca,
    ):
        return "SIM"
    if re.search(r"ISSQN?\s*RETIDO\??\s*NAO|RETENCAO\s*DO\s*ISSQN?\s*NAO|TIPO\s*DE\s*RECOLHIMENTO\s*PROPRIO", busca):
        return "NÃO"

    iss_retido_valor = _extrair_por_regex_original(texto, busca, [r"ISSQN?\s*RETIDO\s*R\$?\s*([\d\.,]+)"])
    if converter_moeda_br_para_decimal(iss_retido_valor) > 0:
        return "SIM"

    valor_liquido = extrair_valor_liquido(texto, busca)
    if valor_servico <= 0 or aliquota <= 0 or valor_liquido <= 0:
        return "NÃO"

    iss_calculado = (valor_servico * aliquota / Decimal("100")).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    esperado = (valor_servico - iss_calculado).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return "SIM" if valor_liquido.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) == esperado else "NÃO"


def extrair_nota_fiscal(caminho_pdf: Path) -> NotaFiscalExtraida:
    texto = ler_texto_pdf(caminho_pdf)
    busca = texto_para_busca(texto)

    cnpj_prestador = extrair_cnpj_prestador(texto, busca)
    numero_nf = extrair_numero_nf(texto, busca)
    data_emissao = extrair_data_emissao(texto, busca)
    id_cnae, desc_cnae = extrair_id_desc_cnae(texto, busca)
    descricao_servico = extrair_descricao_servico(texto, busca)
    cidade_local_prestacao, uf_local_prestacao = extrair_cidade_uf_local_prestacao(texto, busca)
    valor_servico = extrair_valor_servico(texto, busca)
    aliquota = extrair_aliquota(texto, busca)

    ajustes_layout = _ajustes_por_layout_especifico(texto, busca)
    cnpj_prestador = str(ajustes_layout.get("cnpj_prestador", cnpj_prestador) or cnpj_prestador)
    numero_nf = str(ajustes_layout.get("numero_nf", numero_nf) or numero_nf)
    data_emissao = str(ajustes_layout.get("data_emissao", data_emissao) or data_emissao)
    id_cnae = str(ajustes_layout.get("id_cnae", id_cnae) or id_cnae)
    desc_cnae = str(ajustes_layout.get("desc_cnae", desc_cnae) or desc_cnae)
    descricao_servico = str(ajustes_layout.get("descricao_servico", descricao_servico) or descricao_servico)
    uf_local_prestacao = str(ajustes_layout.get("uf_local_prestacao", uf_local_prestacao) or uf_local_prestacao)
    cidade_local_prestacao = str(
        ajustes_layout.get("cidade_local_prestacao", cidade_local_prestacao) or cidade_local_prestacao
    )
    valor_servico = ajustes_layout.get("valor_servico", valor_servico) or valor_servico
    aliquota = ajustes_layout.get("aliquota", aliquota) or aliquota
    iss_retido = extrair_iss_retido(texto, busca, valor_servico, aliquota)
    natureza_operacao = determinar_natureza_operacao(cidade_local_prestacao)

    registro_parcial = NotaFiscalExtraida(
        arquivo_pdf=caminho_pdf.name,
        cnpj_prestador=limpar_cnpj(cnpj_prestador),
        numero_nf=normalizar_numero_nf(numero_nf),
        data_emissao=normalizar_data_br(data_emissao),
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=descricao_servico,
        uf_local_prestacao=uf_local_prestacao,
        cidade_local_prestacao=cidade_local_prestacao,
        natureza_operacao=natureza_operacao,
        iss_retido=iss_retido,
        valor_servico=formatar_moeda_br(valor_servico) if valor_servico else "",
        aliquota=formatar_percentual(aliquota) if aliquota else "",
    )

    campos_vazios = registro_parcial.campos_vazios()
    dados_locais = {
        "cnpj_prestador": registro_parcial.cnpj_prestador,
        "numero_nf": registro_parcial.numero_nf,
        "data_emissao": registro_parcial.data_emissao,
        "id_cnae": registro_parcial.id_cnae,
        "desc_cnae": registro_parcial.desc_cnae,
        "descricao_servico": registro_parcial.descricao_servico,
        "uf_local_prestacao": registro_parcial.uf_local_prestacao,
        "cidade_local_prestacao": registro_parcial.cidade_local_prestacao,
        "valor_servico": registro_parcial.valor_servico,
        "aliquota": registro_parcial.aliquota,
        "iss_retido": registro_parcial.iss_retido,
    }

    dados_gemini = extrair_campos_gemini(caminho_pdf, texto, dados_locais, campos_vazios)
    if dados_gemini:
        cnpj_prestador = _mesclar_campo_textual(
            cnpj_prestador,
            dados_gemini.get("cnpj_prestador", ""),
            texto,
            "cnpj_prestador",
        )
        numero_nf = _mesclar_campo_textual(
            numero_nf,
            dados_gemini.get("numero_nf", ""),
            texto,
            "numero_nf",
        )
        data_emissao = _mesclar_campo_textual(
            data_emissao,
            dados_gemini.get("data_emissao", ""),
            texto,
            "data_emissao",
        )
        id_cnae = _mesclar_campo_textual(
            id_cnae,
            dados_gemini.get("id_cnae", ""),
            texto,
            "id_cnae",
        )
        desc_cnae = _mesclar_campo_textual(
            desc_cnae,
            dados_gemini.get("desc_cnae", ""),
            texto,
            "desc_cnae",
        )
        descricao_servico = _mesclar_campo_textual(
            descricao_servico,
            dados_gemini.get("descricao_servico", ""),
            texto,
            "descricao_servico",
        )
        uf_local_prestacao = _mesclar_campo_textual(
            uf_local_prestacao,
            dados_gemini.get("uf_local_prestacao", ""),
            texto,
            "uf_local_prestacao",
        )
        cidade_local_prestacao = _mesclar_campo_textual(
            cidade_local_prestacao,
            dados_gemini.get("cidade_local_prestacao", ""),
            texto,
            "cidade_local_prestacao",
        )
        valor_servico = _mesclar_campo_decimal(
            valor_servico,
            dados_gemini.get("valor_servico", ""),
            texto,
            "valor_servico",
        )
        aliquota = _mesclar_campo_decimal(
            aliquota,
            dados_gemini.get("aliquota", ""),
            texto,
            "aliquota",
        )
        if not iss_retido.strip():
            iss_retido = dados_gemini.get("iss_retido", iss_retido).strip().upper()

    natureza_operacao = determinar_natureza_operacao(cidade_local_prestacao)

    return NotaFiscalExtraida(
        arquivo_pdf=caminho_pdf.name,
        cnpj_prestador=limpar_cnpj(cnpj_prestador),
        numero_nf=numero_nf,
        data_emissao=data_emissao,
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=descricao_servico,
        uf_local_prestacao=uf_local_prestacao,
        cidade_local_prestacao=cidade_local_prestacao,
        natureza_operacao=natureza_operacao,
        iss_retido=iss_retido,
        valor_servico=formatar_moeda_br(valor_servico) if valor_servico else "",
        aliquota=formatar_percentual(aliquota) if aliquota else "",
    )


def listar_pdfs(pasta: Path) -> Iterable[Path]:
    return sorted(pasta.glob("*.pdf"))


def ajustar_largura_colunas(planilha) -> None:
    for coluna in planilha.columns:
        maior = 0
        letra_coluna = get_column_letter(coluna[0].column)
        for celula in coluna:
            if celula.value is not None:
                maior = max(maior, len(str(celula.value)))
        planilha.column_dimensions[letra_coluna].width = min(maior + 2, 80)


def gerar_xlsx(registros: list[NotaFiscalExtraida], destino: Path) -> None:
    # Enriquecemos os registros com o CNAE oficial antes de gravar a planilha final.
    enriquecer_registros_cnae_final(registros)

    wb = Workbook()
    ws = wb.active
    ws.title = "NF Extraídas"

    cabecalhos = [
        "ARQUIVO_PDF",
        "CNPJ_PRESTADOR",
        "NUMERO_NF",
        "DATA_EMISSAO",
        "ID_CNAE",
        "DESC_CNAE",
        "DESCRICAO_SERVICO",
        "UF_LOCAL_PRESTACAO",
        "CIDADE_LOCAL_PRESTACAO",
        "NATUREZA_OPERACAO",
        "ISS_RETIDO",
        "VALOR_SERVICO",
        "ALIQUOTA",
        "ID_CNAE_FINAL",
        "DESC_CNAE_FINAL",
    ]

    ws.append(cabecalhos)
    for registro in registros:
        ws.append(registro.como_linha())

    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")

    ajustar_largura_colunas(ws)
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)


def registrar_log_incompletos(
    registros: list[NotaFiscalExtraida],
    destino_xlsx: Path,
) -> Path:
    """Cria um log detalhado quando alguma NF não preencher todas as colunas.

    O objetivo é permitir auditoria rápida dos arquivos que precisam de ajuste
    manual ou de novas regras de extração.
    """
    linhas_log: list[str] = []

    for registro in registros:
        campos_vazios = registro.campos_vazios()
        if not campos_vazios:
            continue

        linhas_log.append(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {registro.arquivo_pdf}"
        )
        linhas_log.append(f"Campos ausentes: {', '.join(campos_vazios)}")
        linhas_log.append(f"Resumo: {registro.resumo_campos()}")
        linhas_log.append("")

    caminho_log = destino_xlsx.with_name(f"{destino_xlsx.stem}_log_extracao.txt")
    if linhas_log:
        caminho_log.write_text("\n".join(linhas_log), encoding="utf-8")
    elif caminho_log.exists():
        caminho_log.unlink()

    return caminho_log


def abrir_site_iss_fortaleza() -> None:
    """Mantido apenas por compatibilidade histórica com versões anteriores."""
    pass


def solicitar_opcao() -> str:
    """Pede ao usuário que escolha entre gerar o XLSX ou executar a automação.

    A interface é propositalmente direta para reduzir erro operacional e deixar
    visível a próxima etapa do fluxo de trabalho.
    """
    while True:
        print("Escolha uma opção:")
        print("1 - Gerar o XLSX a partir dos PDFs")
        print("2 - Executar a automação visível da ISS de Fortaleza")
        resposta = input("Digite 1 ou 2: ").strip()
        if resposta in {"1", "2"}:
            return resposta
        print("Opção inválida. Por favor, digite 1 ou 2.")
        print()


def solicitar_planilha_automacao() -> Path:
    """Pergunta ao usuário qual planilha será usada na automação da ISS."""

    entrada = input(
        "Informe o caminho da planilha XLSX que será usada na automação da ISS "
        "(Enter para `nf_compilado.xlsx`): "
    ).strip().strip('"')
    if not entrada:
        return Path("nf_compilado.xlsx")
    return Path(entrada)


def construir_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai dados de NFS-e em PDF, gera XLSX ou executa a automação da ISS."
    )
    parser.add_argument(
        "-m",
        "--modo",
        choices=("1", "2"),
        default=None,
        help="1 para gerar o XLSX; 2 para executar a automação da ISS de Fortaleza.",
    )
    parser.add_argument(
        "-c",
        "--competencia",
        default=None,
        help="Competência a trabalhar na ISS, por exemplo '5/2026', '5 2026' ou 'maio 2026'.",
    )
    parser.add_argument(
        "--planilha",
        type=Path,
        default=None,
        help="Planilha XLSX usada pela automação da ISS.",
    )
    parser.add_argument(
        "pasta",
        nargs="?",
        type=Path,
        default=None,
        help="Pasta que contém os arquivos PDF. Se omitido, o script solicita no terminal.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("nf_compilado.xlsx"),
        help="Caminho do arquivo XLSX de saída.",
    )
    return parser.parse_args()


def main() -> int:
    args = construir_argumentos()
    registrar_evento_execucao("Fluxo principal iniciado no extrair_nf_pdfs.py", "extrair_nf_pdfs.py")

    modo = args.modo
    if modo is None:
        if args.pasta is not None:
            modo = "1"
        else:
            modo = solicitar_opcao()

    if modo == "2":
        try:
            competencia = (
                interpretar_competencia(args.competencia)
                if args.competencia is not None
                else solicitar_competencia()
            )
        except ValueError as exc:
            raise SystemExit(f"Competência inválida: {exc}") from exc
        planilha = args.planilha if args.planilha is not None else solicitar_planilha_automacao()
        if not planilha.exists():
            raise SystemExit(
                f"A planilha informada para a automação não existe: {planilha}"
            )
        registrar_evento_execucao(
            f"Modo 2 selecionado com planilha {planilha.resolve()}",
            "extrair_nf_pdfs.py",
        )
        executar_automacao_iss(planilha, competencia)
        return 0

    pasta = args.pasta
    if pasta is None:
        entrada = input("Informe o caminho da pasta com os PDFs: ").strip().strip('"')
        pasta = Path(entrada)

    destino = args.output

    if not pasta.exists() or not pasta.is_dir():
        raise SystemExit(f"A pasta informada não existe ou não é válida: {pasta}")

    pdfs = list(listar_pdfs(pasta))
    if not pdfs:
        raise SystemExit("Nenhum PDF foi encontrado na pasta informada.")

    registrar_evento_execucao(
        f"Modo 1 selecionado com {len(pdfs)} PDF(s) em {pasta.resolve()}",
        "extrair_nf_pdfs.py",
    )

    registros: list[NotaFiscalExtraida] = []
    for indice, pdf in enumerate(pdfs, start=1):
        print(f"[{indice}/{len(pdfs)}] Processando PDF: {pdf.name}", flush=True)
        registrar_evento_execucao(
            f"Iniciando processamento do PDF {indice}/{len(pdfs)}: {pdf.name}",
            "extrair_nf_pdfs.py",
        )
        try:
            registros.append(extrair_nota_fiscal(pdf))
        except Exception as exc:
            print(f"[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}: {exc}", flush=True)
            registrar_evento_execucao(
                f"Falha ao processar {pdf.name}: {exc}",
                "extrair_nf_pdfs.py",
            )
            registros.append(
                NotaFiscalExtraida(
                    arquivo_pdf=pdf.name,
                    descricao_servico=f"ERRO NA EXTRAÇÃO: {exc}",
                )
            )
        else:
            registrar_evento_execucao(
                f"Processamento concluído com sucesso para {pdf.name}",
                "extrair_nf_pdfs.py",
            )

    registros_completos = [registro for registro in registros if not registro.campos_vazios()]
    gerar_xlsx(registros_completos, destino)
    caminho_log = registrar_log_incompletos(registros, destino)
    print(f"Arquivo XLSX gerado com sucesso: {destino.resolve()}")
    print(f"Total de PDFs processados: {len(pdfs)}")
    print(f"Linhas completas exportadas: {len(registros_completos)}")
    if caminho_log.exists():
        print(f"Log de extração gerado em: {caminho_log.resolve()}")
    registrar_evento_execucao(
        f"XLSX gerado em {destino.resolve()} com {len(registros_completos)} linha(s) exportada(s)",
        "extrair_nf_pdfs.py",
    )
    print()
    print(
        "Próximo passo sugerido: escolha a opção 2 para iniciar a automação visível da ISS de Fortaleza."
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit as exc:
        if exc.code not in (0, None):
            registrar_erro(exc, "extrair_nf_pdfs.py")
        raise
    except Exception as exc:
        registrar_erro(exc, "extrair_nf_pdfs.py")
        raise

