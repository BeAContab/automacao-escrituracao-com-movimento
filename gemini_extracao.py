from __future__ import annotations

import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any

try:
    from google import genai
except ImportError:  # pragma: no cover - dependência opcional em ambiente sem SDK
    genai = None

try:
    from google.genai import types
except ImportError:  # pragma: no cover - dependência opcional em ambiente sem SDK
    types = None

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - dependência opcional em ambiente sem SDK
    fitz = None


GEMINI_MODEL_PADRAO = "gemini-2.5-flash"
GEMINI_MODELOS_FALLBACK = [
    "gemini-2.5-flash-lite",
    "gemini-flash-lite-latest",
]

ESQUEMA_EXTRAIR_NF = {
    "type": "object",
    "properties": {
        "cnpj_prestador": {"type": ["string", "null"], "description": "CNPJ do prestador."},
        "numero_nf": {"type": ["string", "null"], "description": "Número da nota fiscal."},
        "data_emissao": {"type": ["string", "null"], "description": "Data de emissão em DD/MM/AAAA."},
        "id_cnae": {"type": ["string", "null"], "description": "Código CNAE ou código de serviço encontrado no documento."},
        "desc_cnae": {"type": ["string", "null"], "description": "Descrição oficial ou descritiva do CNAE."},
        "descricao_servico": {"type": ["string", "null"], "description": "Discriminação completa dos serviços."},
        "uf_local_prestacao": {"type": ["string", "null"], "description": "UF da prestação, com duas letras."},
        "cidade_local_prestacao": {"type": ["string", "null"], "description": "Cidade da prestação do serviço."},
        "valor_servico": {"type": ["string", "null"], "description": "Valor total do serviço no padrão brasileiro."},
        "aliquota": {"type": ["string", "null"], "description": "Alíquota percentual no padrão brasileiro."},
        "iss_retido": {"type": ["string", "null"], "description": "SIM ou NÃO."},
    },
    "required": [
        "cnpj_prestador",
        "numero_nf",
        "data_emissao",
        "id_cnae",
        "desc_cnae",
        "descricao_servico",
        "uf_local_prestacao",
        "cidade_local_prestacao",
        "valor_servico",
        "aliquota",
        "iss_retido",
    ],
    "additionalProperties": False,
}


def _carregar_env_local(caminho: Path = Path(".env")) -> None:
    """Carrega o .env local sem depender de bibliotecas extras."""

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
    """Normaliza acentos e pontuação para comparação simples."""

    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(caractere for caractere in texto if unicodedata.category(caractere) != "Mn")
    texto = re.sub(r"[^0-9A-Za-z]+", " ", texto)
    return " ".join(texto.upper().split())


def _somente_digitos(texto: str) -> str:
    return re.sub(r"\D+", "", texto or "")


def _texto_baixa_confianca(texto: str) -> bool:
    """Identifica textos com sinais fortes de extração ruim ou PDF escaneado."""

    texto = texto or ""
    if len(texto) < 250:
        return True

    caracteres_invalidos = sum(1 for caractere in texto if ord(caractere) < 32 and caractere not in "\n\r\t")
    if caracteres_invalidos / max(len(texto), 1) > 0.03:
        return True

    letras_e_espacos = sum(1 for caractere in texto if caractere.isalpha() or caractere.isspace())
    if letras_e_espacos / max(len(texto), 1) < 0.45:
        return True

    return False


def _numero_nf_suspeito(numero_nf: str, texto_local: str) -> bool:
    """Sinaliza números curtos demais quando o texto local já parece degradado."""

    if not _texto_baixa_confianca(texto_local):
        return False

    digitos = _somente_digitos(numero_nf)
    return not digitos or len(digitos) <= 3


def _presente_no_texto(texto: str, valor: str, tipo: str) -> bool:
    """Verifica se o valor local parece consistente com o texto bruto."""

    if not valor:
        return False

    if tipo in {"numero_nf", "cnpj_prestador", "valor_servico", "aliquota"}:
        return _somente_digitos(valor) in _somente_digitos(texto)

    if tipo == "data_emissao":
        return _somente_digitos(valor) in _somente_digitos(texto)

    if tipo in {"cidade_local_prestacao", "desc_cnae", "descricao_servico"}:
        return _normalizar_texto(valor) in _normalizar_texto(texto)

    if tipo == "uf_local_prestacao":
        return _normalizar_texto(valor) in _normalizar_texto(texto)

    return bool(valor.strip())


def deve_usar_gemini(
    texto_local: str,
    dados_locais: dict[str, str],
    campos_vazios: list[str],
) -> bool:
    """Define quando vale acionar o Gemini como fallback de extração."""

    campos_criticos = {
        "CNPJ_PRESTADOR",
        "NUMERO_NF",
        "DATA_EMISSAO",
        "DESCRICAO_SERVICO",
        "UF_LOCAL_PRESTACAO",
        "CIDADE_LOCAL_PRESTACAO",
        "VALOR_SERVICO",
        "ALIQUOTA",
    }

    if _texto_baixa_confianca(texto_local):
        return True

    if any(campo in campos_criticos for campo in campos_vazios):
        return True

    for chave, tipo in [
        ("numero_nf", "numero_nf"),
        ("cnpj_prestador", "cnpj_prestador"),
        ("data_emissao", "data_emissao"),
        ("cidade_local_prestacao", "cidade_local_prestacao"),
        ("uf_local_prestacao", "uf_local_prestacao"),
        ("valor_servico", "valor_servico"),
        ("aliquota", "aliquota"),
    ]:
        if dados_locais.get(chave, "").strip() and not _presente_no_texto(texto_local, dados_locais[chave], tipo):
            return True

    return False


def _obter_chave_gemini() -> str:
    _carregar_env_local()
    return (
        os.getenv("GEMINI_API_KEY", "").strip()
        or os.getenv("GOOGLE_API_KEY", "").strip()
    )


def _obter_modelo_gemini() -> str:
    _carregar_env_local()
    return os.getenv("GEMINI_MODEL", GEMINI_MODEL_PADRAO).strip() or GEMINI_MODEL_PADRAO


def _modelos_candidatos_gemini() -> list[str]:
    """Lista os modelos que vamos tentar em ordem de preferência."""

    candidatos: list[str] = []
    modelo_principal = _obter_modelo_gemini()
    if modelo_principal:
        candidatos.append(modelo_principal)

    for modelo in GEMINI_MODELOS_FALLBACK:
        if modelo not in candidatos:
            candidatos.append(modelo)

    return candidatos


def _normalizar_resposta(resultado: dict[str, Any]) -> dict[str, str]:
    """Garante strings limpas e previsíveis para a fusão com o parser local."""

    campos = {
        "cnpj_prestador": "",
        "numero_nf": "",
        "data_emissao": "",
        "id_cnae": "",
        "desc_cnae": "",
        "descricao_servico": "",
        "uf_local_prestacao": "",
        "cidade_local_prestacao": "",
        "valor_servico": "",
        "aliquota": "",
        "iss_retido": "",
    }

    for chave in campos:
        valor = resultado.get(chave, "")
        if valor is None:
            continue
        campos[chave] = str(valor).strip()

    return campos


def _refinar_campos_por_imagem(caminho_pdf: Path, cliente: Any) -> dict[str, str]:
    """Refina campos cr?ticos renderizando a primeira p?gina como imagem.

    Esse passo ? reservado para PDFs com texto local muito fraco, porque nesses
    casos o Gemini lendo o PDF bruto pode se apoiar em n?meros perif?ricos e
    devolver um valor incorreto para o campo mais sens?vel da extra??o.
    """

    if fitz is None:
        return {}

    try:
        documento = fitz.open(str(caminho_pdf))
        if documento.page_count <= 0:
            return {}

        pagina = documento[0]
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
        imagem_png = pixmap.tobytes("png")
    except Exception:
        return {}

    parte_imagem = types.Part.from_bytes(data=imagem_png, mime_type="image/png")
    prompt = (
        "Leia a imagem desta NFS-e e retorne apenas JSON v?lido com as chaves "
        '"numero_nf" e "iss_retido". O valor correto de numero_nf ? o campo '
        'explicitamente rotulado como "N?mero da Nota" ou equivalente. O valor '
        'de iss_retido deve ser SIM apenas se o documento deixar isso expl?cito, '
        'como em "O ISS desta NFS-e ser? RETIDO pelo Tomador de Servi?o". '
        "Ignore s?rie, protocolo, parcela, c?digo, p?gina, lote, autoriza??o e "
        "qualquer outro n?mero do documento."
    )

    configuracao = {
        "temperature": 0,
        "response_mime_type": "application/json",
    }

    for modelo in _modelos_candidatos_gemini():
        try:
            resposta = cliente.models.generate_content(
                model=modelo,
                contents=[parte_imagem, prompt],
                config=configuracao,
            )
        except Exception:
            continue

        texto_resposta = getattr(resposta, "text", "") or ""
        if not texto_resposta:
            continue

        try:
            bruto = json.loads(texto_resposta)
        except json.JSONDecodeError:
            match = re.search(r"\{.*\}", texto_resposta, flags=re.DOTALL)
            if not match:
                continue
            try:
                bruto = json.loads(match.group(0))
            except json.JSONDecodeError:
                continue

        if not isinstance(bruto, dict):
            continue

        numero_nf = str(bruto.get("numero_nf", "") or "").strip()
        iss_retido = str(bruto.get("iss_retido", "") or "").strip().upper()
        if numero_nf or iss_retido:
            return {
                "numero_nf": numero_nf,
                "iss_retido": iss_retido,
            }

    return {}

def extrair_campos_gemini(
    caminho_pdf: Path,
    texto_local: str,
    dados_locais: dict[str, str],
    campos_vazios: list[str],
) -> dict[str, str]:
    """Consulta o Gemini para completar ou corrigir campos problemáticos da NF."""

    if genai is None or types is None:
        return {}

    chave = _obter_chave_gemini()
    if not chave:
        return {}

    if not deve_usar_gemini(texto_local, dados_locais, campos_vazios):
        return {}

    cliente = genai.Client(api_key=chave)
    prompt = (
        "Você é um extrator fiscal. Leia o PDF anexado e retorne apenas JSON válido. "
        "Preencha os campos da NFS-e com o maior cuidado possível. "
        "Priorize o documento em caso de conflito e corrija valores que estejam visíveis, "
        "mas não foram capturados corretamente pelo parser local. "
        "Regras de formatação:\n"
        "- numero_nf: somente dígitos.\n"
        "- data_emissao: DD/MM/AAAA.\n"
        "- cnpj_prestador: 00.000.000/0000-00.\n"
        "- valor_servico: número no padrão brasileiro, sem símbolo de moeda.\n"
        "- aliquota: percentual numérico, por exemplo 5,00.\n"
        "- uf_local_prestacao: duas letras.\n"
        "- cidade_local_prestacao: nome da cidade.\n"
        "- iss_retido: SIM ou NÃO.\n"
        "Se algum campo realmente não existir no documento, devolva string vazia.\n\n"
        f"Campos já extraídos localmente:\n{json.dumps(dados_locais, ensure_ascii=False, indent=2)}\n\n"
        f"Campos em branco ou suspeitos: {', '.join(campos_vazios) if campos_vazios else 'nenhum'}"
    )

    # Enviamos o PDF em memória para evitar falhas com nomes de arquivo com acentos.
    conteudo_pdf = caminho_pdf.read_bytes()
    parte_pdf = types.Part.from_bytes(data=conteudo_pdf, mime_type="application/pdf")

    configuracao = {
        "temperature": 0,
        "response_mime_type": "application/json",
        "response_json_schema": ESQUEMA_EXTRAIR_NF,
    }

    for modelo in _modelos_candidatos_gemini():
        try:
            resposta = cliente.models.generate_content(
                model=modelo,
                contents=[parte_pdf, prompt],
                config=configuracao,
            )
        except Exception as exc:
            mensagem = str(exc).upper()
            if "RESOURCE_EXHAUSTED" in mensagem or "QUOTA" in mensagem:
                continue
            if "NOT_FOUND" in mensagem or "MODEL" in mensagem:
                continue
            continue

        texto_resposta = getattr(resposta, "text", "") or ""
        if not texto_resposta:
            continue

        try:
            bruto = json.loads(texto_resposta)
        except json.JSONDecodeError:
            match = re.search(r"\{.*\}", texto_resposta, flags=re.DOTALL)
            if not match:
                continue
            try:
                bruto = json.loads(match.group(0))
            except json.JSONDecodeError:
                continue

        if not isinstance(bruto, dict):
            continue

        campos_normalizados = _normalizar_resposta(bruto)
        if _numero_nf_suspeito(campos_normalizados.get("numero_nf", ""), texto_local) or (
            campos_normalizados.get("iss_retido", "").strip().upper() == "N?O"
            and _texto_baixa_confianca(texto_local)
        ):
            refinados = _refinar_campos_por_imagem(caminho_pdf, cliente)
            numero_nf_refinado = refinados.get("numero_nf", "").strip()
            iss_retido_refinado = refinados.get("iss_retido", "").strip().upper()
            if numero_nf_refinado:
                campos_normalizados["numero_nf"] = numero_nf_refinado
            if iss_retido_refinado == "SIM":
                campos_normalizados["iss_retido"] = iss_retido_refinado

        return campos_normalizados

    return {}
