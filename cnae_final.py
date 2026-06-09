"""Classificacao do CNAE oficial para enriquecimento do XLSX da funcao 1."""

from __future__ import annotations

import json
import os
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from openpyxl import load_workbook

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL_PADRAO = "llama-3.1-8b-instant"
LIMITE_CANDIDATOS_GROQ = 10


@dataclass(frozen=True)
class CnaeOficial:
    """Representa uma linha do arquivo oficial de CNAE."""

    codigo: str
    descricao: str
    descricao_normalizada: str
    tokens: frozenset[str]


def _carregar_env_local(caminho: Path = Path(".env")) -> None:
    """Carrega variáveis locais do .env sem depender de bibliotecas extras."""

    if not caminho.exists():
        return

    for linha in caminho.read_text(encoding="utf-8").splitlines():
        texto = linha.strip()
        if not texto or texto.startswith("#") or "=" not in texto:
            continue

        chave, valor = texto.split("=", 1)
        chave = chave.strip()
        valor = valor.strip().strip('"').strip("'")
        if chave and chave not in os.environ:
            os.environ[chave] = valor


def _normalizar_texto(texto: str) -> str:
    """Normaliza acentos, pontuacao e espacos para comparar descricoes."""

    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(caractere for caractere in texto if unicodedata.category(caractere) != "Mn")
    texto = re.sub(r"[^0-9A-Za-z]+", " ", texto)
    return " ".join(texto.lower().split())


def _tokenizar(texto: str) -> frozenset[str]:
    """Separa a descricao em tokens uteis para pontuacao local."""

    palavras = {
        "de",
        "da",
        "do",
        "das",
        "dos",
        "e",
        "em",
        "para",
        "com",
        "sem",
        "ou",
        "a",
        "o",
        "as",
        "os",
        "na",
        "no",
        "nas",
        "nos",
        "servico",
        "servicos",
        "atividade",
        "atividades",
        "outro",
        "outros",
        "especificado",
        "especificados",
        "anteriormente",
    }
    tokens = [token for token in _normalizar_texto(texto).split() if token not in palavras]
    return frozenset(tokens)


IDS_QUE_FORCAM_ISS_RETIDO = {"932989910", "900190201"}
CODIGO_CNAE_EVENTOS = "932989910"


@lru_cache(maxsize=1)
def carregar_cnaes_oficiais(caminho: str = "cnae_oficial.xlsx") -> tuple[CnaeOficial, ...]:
    """Lê o arquivo oficial de CNAE e devolve a tabela pronta para busca."""

    arquivo = Path(caminho)
    if not arquivo.exists():
        raise FileNotFoundError(f"Arquivo oficial de CNAE não encontrado: {arquivo}")

    workbook = load_workbook(arquivo, data_only=True)
    planilha = workbook.active
    cabecalhos = [str(celula.value or "").strip() for celula in next(planilha.iter_rows(min_row=1, max_row=1))]
    mapa = {cabecalho.lower(): indice for indice, cabecalho in enumerate(cabecalhos)}

    indice_codigo = mapa.get("cnae", 0)
    indice_descricao = mapa.get("descricao", 1)

    cnaes: list[CnaeOficial] = []
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        if not linha or indice_codigo >= len(linha) or indice_descricao >= len(linha):
            continue

        codigo = str(linha[indice_codigo] or "").strip()
        descricao = str(linha[indice_descricao] or "").strip()
        if not codigo or not descricao:
            continue

        descricao_normalizada = _normalizar_texto(descricao)
        tokens = _tokenizar(descricao)
        cnaes.append(
            CnaeOficial(
                codigo=codigo,
                descricao=descricao,
                descricao_normalizada=descricao_normalizada,
                tokens=tokens,
            )
        )

    return tuple(cnaes)


def _pontuar_texto(consulta: str, candidata: CnaeOficial) -> float:
    """Pontua uma descricao oficial contra o texto enviado pelo PDF."""

    consulta_normalizada = _normalizar_texto(consulta)
    if not consulta_normalizada:
        return 0.0

    score = 0.0
    if consulta_normalizada == candidata.descricao_normalizada:
        score += 120.0
    if consulta_normalizada in candidata.descricao_normalizada:
        score += 60.0
    if candidata.descricao_normalizada in consulta_normalizada:
        score += 40.0

    tokens_consulta = _tokenizar(consulta)
    interseccao = len(tokens_consulta & candidata.tokens)
    score += interseccao * 10.0

    score += SequenceMatcher(None, consulta_normalizada, candidata.descricao_normalizada).ratio() * 25.0
    return score


def _pontuar_candidato(desc_cnae: str, descricao_servico: str, id_cnae: str, candidata: CnaeOficial) -> float:
    """Combina o texto do PDF e a descrição do serviço para ranquear o CNAE."""

    score = 0.0
    score += _pontuar_texto(desc_cnae, candidata) * 0.75
    score += _pontuar_texto(descricao_servico, candidata) * 0.25

    codigo_extraido = re.sub(r"\D+", "", id_cnae or "")
    codigo_oficial = re.sub(r"\D+", "", candidata.codigo)
    if codigo_extraido and codigo_extraido == codigo_oficial:
        score += 80.0
    elif codigo_extraido and codigo_extraido in codigo_oficial:
        score += 20.0

    return score


def _selecionar_candidatos(
    desc_cnae: str,
    descricao_servico: str,
    id_cnae: str,
    oficiais: tuple[CnaeOficial, ...],
    limite: int = LIMITE_CANDIDATOS_GROQ,
) -> list[CnaeOficial]:
    """Escolhe um subconjunto pequeno de CNAEs para a IA analisar com mais foco."""

    if not oficiais:
        return []

    ranqueados = sorted(
        oficiais,
        key=lambda candidata: _pontuar_candidato(desc_cnae, descricao_servico, id_cnae, candidata),
        reverse=True,
    )
    return ranqueados[:limite]


def _obter_chave_groq() -> str:
    """Lê a chave da Groq do ambiente ou do arquivo .env local."""

    _carregar_env_local()
    return os.getenv("GROQ_API_KEY", "").strip()


def _obter_modelo_groq() -> str:
    """Permite trocar o modelo por variavel de ambiente, mantendo um padrao economico."""

    _carregar_env_local()
    return os.getenv("GROQ_MODEL", GROQ_MODEL_PADRAO).strip() or GROQ_MODEL_PADRAO


def _normalizar_codigo_cnae(codigo: str) -> str:
    """Remove pontuação do código para facilitar regras exatas de negócio."""

    return re.sub(r"\D+", "", codigo or "")


def _pontuar_descricao_oficial(desc_cnae: str, candidata: CnaeOficial) -> float:
    """Mede o quanto a descri??o extra?da parece a descri??o oficial da CNAE."""

    score = _pontuar_texto(desc_cnae, candidata)
    descricao_normalizada = _normalizar_texto(desc_cnae)
    if descricao_normalizada == candidata.descricao_normalizada:
        score += 200.0
    elif descricao_normalizada in candidata.descricao_normalizada:
        score += 120.0
    elif candidata.descricao_normalizada in descricao_normalizada:
        score += 90.0
    return score


def _resolver_por_descricao_explicita(
    desc_cnae: str,
    oficiais: tuple[CnaeOficial, ...],
) -> tuple[str, str] | None:
    """Tenta resolver o CNAE apenas pela descri??o vis?vel no PDF."""

    if not desc_cnae.strip() or not oficiais:
        return None

    tokens_desc = _tokenizar(desc_cnae)
    if {"organizacao", "feiras", "congressos"}.issubset(tokens_desc) or {"organizacao", "feiras", "exposicoes"}.issubset(tokens_desc):
        for candidata in oficiais:
            if _normalizar_codigo_cnae(candidata.codigo) == CODIGO_CNAE_EVENTOS:
                return candidata.codigo, candidata.descricao
        return CODIGO_CNAE_EVENTOS, desc_cnae.strip()

    ranqueados = sorted(
        oficiais,
        key=lambda candidata: _pontuar_descricao_oficial(desc_cnae, candidata),
        reverse=True,
    )
    melhor = ranqueados[0]
    score_melhor = _pontuar_descricao_oficial(desc_cnae, melhor)
    segundo = _pontuar_descricao_oficial(desc_cnae, ranqueados[1]) if len(ranqueados) > 1 else 0.0

    interseccao = len(tokens_desc & melhor.tokens)
    cobertura = interseccao / max(len(melhor.tokens), 1)

    if score_melhor >= 170.0 or (cobertura >= 0.7 and score_melhor >= segundo + 20.0):
        return melhor.codigo, melhor.descricao

    return None


def _classificar_localmente(
    desc_cnae: str,
    descricao_servico: str,
    id_cnae: str,
    oficiais: tuple[CnaeOficial, ...],
) -> tuple[str, str]:
    """Fallback local quando a Groq não estiver disponível ou falhar."""

    candidatos = _selecionar_candidatos(desc_cnae, descricao_servico, id_cnae, oficiais, limite=1)
    if not candidatos:
        return id_cnae.strip(), desc_cnae.strip()
    melhor = candidatos[0]
    return melhor.codigo, melhor.descricao


def _chamar_groq(
    desc_cnae: str,
    descricao_servico: str,
    id_cnae: str,
    candidatos: list[CnaeOficial],
) -> tuple[str, str]:
    """Consulta a Groq com JSON Object Mode para escolher um único CNAE oficial."""

    chave = _obter_chave_groq()
    if not chave:
        raise RuntimeError("GROQ_API_KEY não configurada.")

    modelo = _obter_modelo_groq()
    payload = {
        "model": modelo,
        "temperature": 0,
        "max_completion_tokens": 120,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": (
                    "Você é um especialista em CNAE. "
                    "Escolha exatamente um item da lista oficial que melhor corresponda à NFS-e. "
                    "Baseie-se principalmente em DESC_CNAE_PDF e use DESCRICAO_SERVICO apenas como apoio. "
                    "Responda somente com JSON válido, sem texto extra, no formato "
                    '{"id_cnae_final":"...","desc_cnae_final":"..."}.'
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "desc_cnae_pdf": desc_cnae,
                        "descricao_servico": descricao_servico,
                        "id_cnae_extraido": id_cnae,
                        "candidatos_oficiais": [
                            {"codigo": candidata.codigo, "descricao": candidata.descricao}
                            for candidata in candidatos
                        ],
                    },
                    ensure_ascii=False,
                ),
            },
        ],
    }

    requisicao = Request(
        GROQ_API_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {chave}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(requisicao, timeout=60) as resposta:
            corpo = json.loads(resposta.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Falha ao consultar a Groq: {exc}") from exc

    try:
        conteudo = corpo["choices"][0]["message"]["content"]
        if isinstance(conteudo, str):
            resultado = json.loads(conteudo)
        else:
            resultado = conteudo
    except (KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
        raise RuntimeError("Resposta da Groq em formato inesperado.") from exc

    id_cnae_final = str(resultado.get("id_cnae_final", "")).strip()
    desc_cnae_final = str(resultado.get("desc_cnae_final", "")).strip()
    if not id_cnae_final or not desc_cnae_final:
        raise RuntimeError("Resposta da Groq sem CNAE final completo.")

    return id_cnae_final, desc_cnae_final


def resolver_cnae_final(
    desc_cnae: str,
    descricao_servico: str = "",
    id_cnae: str = "",
    caminho_oficial: str = "cnae_oficial.xlsx",
) -> tuple[str, str]:
    """Resolve o CNAE final com Groq quando poss?vel e fallback local quando necess?rio."""

    try:
        oficiais = carregar_cnaes_oficiais(caminho_oficial)
    except FileNotFoundError:
        return id_cnae.strip(), desc_cnae.strip()

    resolvido_por_descricao = _resolver_por_descricao_explicita(desc_cnae, oficiais)
    if resolvido_por_descricao is not None:
        return resolvido_por_descricao

    candidatos = _selecionar_candidatos(desc_cnae, descricao_servico, id_cnae, oficiais)
    if not candidatos:
        return id_cnae.strip(), desc_cnae.strip()

    try:
        return _chamar_groq(desc_cnae, descricao_servico, id_cnae, candidatos)
    except Exception:
        return _classificar_localmente(desc_cnae, descricao_servico, id_cnae, oficiais)


def enriquecer_registros_cnae_final(registros: list[Any]) -> None:
    """Preenche os campos finais de CNAE nos registros já extraídos dos PDFs."""

    cache: dict[tuple[str, str, str], tuple[str, str]] = {}
    for registro in registros:
        desc_cnae = str(getattr(registro, "desc_cnae", "") or "").strip()
        descricao_servico = str(getattr(registro, "descricao_servico", "") or "").strip()
        id_cnae = str(getattr(registro, "id_cnae", "") or "").strip()
        chave = (_normalizar_texto(desc_cnae), _normalizar_texto(descricao_servico), _normalizar_texto(id_cnae))

        if chave not in cache:
            cache[chave] = resolver_cnae_final(desc_cnae, descricao_servico, id_cnae)

        id_cnae_final, desc_cnae_final = cache[chave]
        setattr(registro, "id_cnae_final", id_cnae_final)
        setattr(registro, "desc_cnae_final", desc_cnae_final)

        if _normalizar_codigo_cnae(id_cnae_final) in IDS_QUE_FORCAM_ISS_RETIDO:
            setattr(registro, "iss_retido", "SIM")
