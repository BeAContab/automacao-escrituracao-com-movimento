from __future__ import annotations

import json
import os
import re
import time
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

try:
    from groq import Groq
except ImportError:  # pragma: no cover
    Groq = None


class ErroCotaGemini(Exception):
    """Exceção levantada quando o limite de cota da API do Gemini é atingido."""
    pass


class ErroCotaGroq(Exception):
    """Exceção levantada quando o limite de cota da API do Groq é atingido."""
    pass


GEMINI_MODEL_PADRAO = "gemini-2.5-flash"
GEMINI_MODELOS_FALLBACK = [
    "gemini-2.5-pro",
    "gemini-2.5-flash-lite",
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
        # Campos adicionais: prefeitura emissora e tributos federais/deduções
        "prefeitura": {"type": ["string", "null"], "description": "Nome completo da prefeitura emissora, ex: PREFEITURA MUNICIPAL DE FORTALEZA."},
        "valor_deducoes": {"type": ["string", "null"], "description": "Valor total das deduções permitidas em lei, no padrão brasileiro (0,00 se ausente)."},
        "descontos_incondicionados": {"type": ["string", "null"], "description": "Valor do desconto incondicionado, no padrão brasileiro (0,00 se ausente)."},
        "descontos_condicionados": {"type": ["string", "null"], "description": "Valor do desconto condicionado, no padrão brasileiro (0,00 se ausente)."},
        "outras_retencoes": {"type": ["string", "null"], "description": "Valor de outras retenções federais não discriminadas, no padrão brasileiro (0,00 se ausente)."},
        "ir": {"type": ["string", "null"], "description": "Valor do IRRF retido, no padrão brasileiro (0,00 se ausente)."},
        "pis_nao_retido": {"type": ["string", "null"], "description": "Valor do PIS não retido, no padrão brasileiro (0,00 se ausente)."},
        "cofins_nao_retido": {"type": ["string", "null"], "description": "Valor do COFINS não retido, no padrão brasileiro (0,00 se ausente)."},
        "csrf": {"type": ["string", "null"], "description": "Valor da CSRF (CSLL+PIS+COFINS retidos), no padrão brasileiro (0,00 se ausente)."},
        "inss": {"type": ["string", "null"], "description": "Valor do INSS retido, no padrão brasileiro (0,00 se ausente)."},
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
        "prefeitura",
        "valor_deducoes",
        "descontos_incondicionados",
        "descontos_condicionados",
        "outras_retencoes",
        "ir",
        "pis_nao_retido",
        "cofins_nao_retido",
        "csrf",
        "inss",
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
        # Campos adicionais de prefeitura e tributos federais
        "prefeitura": "",
        "valor_deducoes": "",
        "descontos_incondicionados": "",
        "descontos_condicionados": "",
        "outras_retencoes": "",
        "ir": "",
        "pis_nao_retido": "",
        "cofins_nao_retido": "",
        "csrf": "",
        "inss": "",
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

    import time
    max_tentativas = 4
    tempo_espera = 12

    for modelo in _modelos_candidatos_gemini():
        try:
            resposta = cliente.models.generate_content(
                model=modelo,
                contents=[parte_imagem, prompt],
                config=configuracao,
            )
        except Exception as exc:
            mensagem = str(exc).upper()
            if "RESOURCE_EXHAUSTED" in mensagem or "QUOTA" in mensagem:
                raise ErroCotaGemini("Limite de cota da API do Gemini excedido (RESOURCE_EXHAUSTED).") from exc
            continue

        if resposta is None:
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
    api_key: str | None = None,
) -> dict[str, str]:
    """Consulta o Gemini para extrair todos os campos da NF."""

    if genai is None or types is None:
        return {}

    chave = api_key or _obter_chave_gemini()
    if not chave:
        return {}

    cliente = genai.Client(api_key=chave)
    from datetime import datetime
    data_atual_str = datetime.now().strftime("%d/%m/%Y")

    instrucoes_formatacao = (
        "Regras de formatação e Negócio:\n"
        "- numero_nf: somente dígitos. NUNCA invente números que não constam no documento.\n"
        f"- data_emissao: DD/MM/AAAA. A data NUNCA pode ser anterior a 3 meses da data atual ({data_atual_str}). Se for muito antiga ou não constar, retorne vazio.\n"
        "- cnpj_prestador: 00.000.000/0000-00. NUNCA invente um CNPJ que não consta no documento.\n"
        "- id_cnae e desc_cnae: As regras de extração dependem da Prefeitura emissora:\n"
        "  * Para PREFEITURA MUNICIPAL DE PETROLINA, extrair EXCLUSIVAMENTE do campo rotulado como 'SERVIÇO NACIONAL'. Ignore 'SERVIÇO NBS' ou 'SERVIÇO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE FORTALEZA, extrair EXCLUSIVAMENTE do campo rotulado como 'CÓDIGO DE ATIVIDADE CNAE'.\n"
        "  * Para Prefeitura Municipal de João Pessoa, extrair EXCLUSIVAMENTE do campo 'Serviço' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para 'Prefeitura Municipal de Goiânia' ou 'Prefeitura Municipal de Goiânia - GO', extrair EXCLUSIVAMENTE do campo 'Atividade Municipal' contido em 'DADOS DO SERVIÇO PRESTADO'.\n"
        "  * Para 'Município do Natal', 'PREFEITURA MUNICIPAL DE OLINDA', 'PREFEITURA MUNICIPAL DE VILA VELHA' ou 'PREFEITURA DA CIDADE DO RIO DE JANEIRO', extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA DO MUNICÍPIO DE SÃO PAULO, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * Para PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS, extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE MARACANAÚ, extrair EXCLUSIVAMENTE do campo 'CODIGO DA ATIVIDADE/SERVIÇO'.\n"
        "  * Para PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE/SP, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * IMPORTANTE: O campo id_cnae deve conter o código do CNAE ou serviço. NUNCA preencha este campo com o CPF, CNPJ ou número de telefone do prestador ou tomador.\n"
        "- descricao_servico: copie e cole o valor textual da discriminação/descrição dos serviços exatamente como aparece no documento, caractere por caractere, sem resumir, reescrever ou parafrasear.\n"
        "- valor_servico: número no padrão brasileiro, sem símbolo de moeda.\n"
        "- aliquota: percentual numérico, por exemplo 5,00.\n"
        "- uf_local_prestacao: duas letras. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- cidade_local_prestacao: nome da cidade. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- iss_retido: SIM ou NÃO.\n"
        "- prefeitura: nome completo da prefeitura emissora em maiúsculas, ex: PREFEITURA MUNICIPAL DE FORTALEZA.\n"
        "- valor_deducoes: valor das deduções em lei no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Deduções') ou parecidos.\n"
        "- descontos_incondicionados: valor dos descontos incondicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Incondicionado') ou parecidos.\n"
        "- descontos_condicionados: valor dos descontos condicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Condicionado') ou parecidos.\n"
        "- outras_retencoes: outras retenções federais não classificadas, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Outras Retenções') ou parecidos.\n"
        "- ir: valor do IRRF retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('IR'), parecidos, ou próximo de 'IRRF' ou similares.\n"
        "- pis_nao_retido: valor do PIS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'PIS' ou parecidos.\n"
        "- cofins_nao_retido: valor do COFINS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'COFINS' ou parecidos.\n"
        "- csrf: valor da CSRF (CSLL+PIS+COFINS retidos juntos), no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('CSRF'), parecidos, ou próximo de 'Contribuições Sociais Retidas' ou similares.\n"
        "- inss: valor do INSS retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('INSS'), parecidos, ou próximo de 'Contribuição Previdenciária Retida' ou similares.\n"
        "Se algum campo realmente não existir no documento, devolva string vazia para textuais ou 0,00 para monetários.\n"
    )

    prompt = (
        "Você é um extrator fiscal. Leia o PDF anexado e retorne apenas JSON válido. "
        "Preencha todos os campos da NFS-e com o maior cuidado possível, extraindo as informações diretamente do documento.\n\n"
        f"{instrucoes_formatacao}"
    )

    # Enviamos o PDF em memória para evitar falhas com nomes de arquivo com acentos.
    conteudo_pdf = caminho_pdf.read_bytes()
    parte_pdf = types.Part.from_bytes(data=conteudo_pdf, mime_type="application/pdf")

    configuracao = {
        "temperature": 0,
        "response_mime_type": "application/json",
        "response_json_schema": ESQUEMA_EXTRAIR_NF,
    }

    import time
    max_tentativas = 4
    tempo_espera = 12

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
                raise ErroCotaGemini("Limite de cota da API do Gemini excedido (RESOURCE_EXHAUSTED).") from exc
            continue

        if resposta is None:
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


def extrair_campos_gemini_somente_texto(
    caminho_pdf: Path,
    texto_local: str,
    api_key: str | None = None,
) -> dict[str, str]:
    """Consulta o Gemini para extrair todos os campos da NF enviando apenas o texto extraído localmente.

    Esta abordagem evita o processamento de imagens e arquivos binários, poupando tokens.
    """

    if genai is None or types is None:
        return {}

    chave = api_key or _obter_chave_gemini()
    if not chave:
        return {}

    cliente = genai.Client(api_key=chave)
    from datetime import datetime
    data_atual_str = datetime.now().strftime("%d/%m/%Y")

    instrucoes_formatacao = (
        "Regras de formatação e Negócio:\n"
        "- numero_nf: somente dígitos. NUNCA invente números que não constam no documento.\n"
        f"- data_emissao: DD/MM/AAAA. A data NUNCA pode ser anterior a 3 meses da data atual ({data_atual_str}). Se for muito antiga ou não constar, retorne vazio.\n"
        "- cnpj_prestador: 00.000.000/0000-00. NUNCA invente um CNPJ que não consta no documento.\n"
        "- id_cnae e desc_cnae: As regras de extração dependem da Prefeitura emissora:\n"
        "  * Para PREFEITURA MUNICIPAL DE PETROLINA, extrair EXCLUSIVAMENTE do campo rotulado como 'SERVIÇO NACIONAL'. Ignore 'SERVIÇO NBS' ou 'SERVIÇO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE FORTALEZA, extrair EXCLUSIVAMENTE do campo rotulado como 'CÓDIGO DE ATIVIDADE CNAE'.\n"
        "  * Para Prefeitura Municipal de João Pessoa, extrair EXCLUSIVAMENTE do campo 'Serviço' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para 'Prefeitura Municipal de Goiânia' ou 'Prefeitura Municipal de Goiânia - GO', extrair EXCLUSIVAMENTE do campo 'Atividade Municipal' contido em 'DADOS DO SERVIÇO PRESTADO'.\n"
        "  * Para 'Município do Natal', 'PREFEITURA MUNICIPAL DE OLINDA', 'PREFEITURA MUNICIPAL DE VILA VELHA' ou 'PREFEITURA DA CIDADE DO RIO DE JANEIRO', extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA DO MUNICÍPIO DE SÃO PAULO, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * Para PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS, extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE MARACANAÚ, extrair EXCLUSIVAMENTE do campo 'CODIGO DA ATIVIDADE/SERVIÇO'.\n"
        "  * Para PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE/SP, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * IMPORTANTE: O campo id_cnae deve conter o código do CNAE ou serviço. NUNCA preencha este campo com o CPF, CNPJ ou número de telefone do prestador ou tomador.\n"
        "- descricao_servico: copie e cole o valor textual da discriminação/descrição dos serviços exatamente como aparece no documento, caractere por caractere, sem resumir, reescrever ou parafrasear.\n"
        "- valor_servico: número no padrão brasileiro, sem símbolo de moeda.\n"
        "- aliquota: percentual numérico, por exemplo 5,00.\n"
        "- uf_local_prestacao: duas letras. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- cidade_local_prestacao: nome da cidade. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- iss_retido: SIM ou NÃO.\n"
        "- prefeitura: nome completo da prefeitura emissora em maiúsculas, ex: PREFEITURA MUNICIPAL DE FORTALEZA.\n"
        "- valor_deducoes: valor das deduções em lei no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Deduções') ou parecidos.\n"
        "- descontos_incondicionados: valor dos descontos incondicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Incondicionado') ou parecidos.\n"
        "- descontos_condicionados: valor dos descontos condicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Condicionado') ou parecidos.\n"
        "- outras_retencoes: outras retenções federais não classificadas, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Outras Retenções') ou parecidos.\n"
        "- ir: valor do IRRF retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('IR'), parecidos, ou próximo de 'IRRF' ou similares.\n"
        "- pis_nao_retido: valor do PIS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'PIS' ou parecidos.\n"
        "- cofins_nao_retido: valor do COFINS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'COFINS' ou parecidos.\n"
        "- csrf: valor da CSRF (CSLL+PIS+COFINS retidos juntos), no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('CSRF'), parecidos, ou próximo de 'Contribuições Sociais Retidas' ou similares.\n"
        "- inss: valor do INSS retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('INSS'), parecidos, ou próximo de 'Contribuição Previdenciária Retida' ou similares.\n"
        "Se algum campo realmente não existir no documento, devolva string vazia para textuais ou 0,00 para monetários.\n"
    )

    prompt = (
        "Você é um extrator fiscal. Leia o texto extraído do documento PDF abaixo e retorne apenas JSON válido. "
        "Preencha todos os campos da NFS-e com o maior cuidado possível, extraindo as informações diretamente do texto apresentado.\n\n"
        f"{instrucoes_formatacao}\n\n"
        f"--- INÍCIO DO TEXTO EXTRAÍDO DO PDF ---\n"
        f"{texto_local}\n"
        f"--- FIM DO TEXTO EXTRAÍDO DO PDF ---"
    )

    configuracao = {
        "temperature": 0,
        "response_mime_type": "application/json",
        "response_json_schema": ESQUEMA_EXTRAIR_NF,
    }

    import time
    max_tentativas = 4
    tempo_espera = 12

    for modelo in _modelos_candidatos_gemini():
        try:
            resposta = cliente.models.generate_content(
                model=modelo,
                contents=[prompt],
                config=configuracao,
            )
        except Exception as exc:
            mensagem = str(exc).upper()
            if "RESOURCE_EXHAUSTED" in mensagem or "QUOTA" in mensagem:
                raise ErroCotaGemini("Limite de cota da API do Gemini excedido (RESOURCE_EXHAUSTED).") from exc
            continue

        if resposta is None:
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
        return campos_normalizados

    return {}


def _obter_chave_groq() -> str:
    _carregar_env_local()
    return os.getenv("GROQ_API_KEY", "").strip()


def _obter_modelo_groq() -> str:
    _carregar_env_local()
    return os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile").strip()


def extrair_campos_groq(texto_local: str, api_key: str | None = None) -> dict[str, str]:
    """Consulta o Groq para extrair todos os campos da NF usando o LLaMA 3."""
    if Groq is None:
        return {}

    chave = api_key or _obter_chave_groq()
    if not chave:
        return {}

    cliente = Groq(api_key=chave)
    from datetime import datetime
    data_atual_str = datetime.now().strftime("%d/%m/%Y")

    instrucoes_formatacao = (
        "Regras de formatação e Negócio:\n"
        "- numero_nf: somente dígitos. NUNCA invente números que não constam no documento.\n"
        f"- data_emissao: DD/MM/AAAA. A data NUNCA pode ser anterior a 3 meses da data atual ({data_atual_str}). Se for muito antiga ou não constar, retorne vazio.\n"
        "- cnpj_prestador: 00.000.000/0000-00. NUNCA invente um CNPJ que não consta no documento.\n"
        "- id_cnae e desc_cnae: As regras de extração dependem da Prefeitura emissora:\n"
        "  * Para PREFEITURA MUNICIPAL DE PETROLINA, extrair EXCLUSIVAMENTE do campo rotulado como 'SERVIÇO NACIONAL'. Ignore 'SERVIÇO NBS' ou 'SERVIÇO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE FORTALEZA, extrair EXCLUSIVAMENTE do campo rotulado como 'CÓDIGO DE ATIVIDADE CNAE'.\n"
        "  * Para Prefeitura Municipal de João Pessoa, extrair EXCLUSIVAMENTE do campo 'Serviço' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para 'Prefeitura Municipal de Goiânia' ou 'Prefeitura Municipal de Goiânia - GO', extrair EXCLUSIVAMENTE do campo 'Atividade Municipal' contido em 'DADOS DO SERVIÇO PRESTADO'.\n"
        "  * Para 'Município do Natal', 'PREFEITURA MUNICIPAL DE OLINDA', 'PREFEITURA MUNICIPAL DE VILA VELHA' ou 'PREFEITURA DA CIDADE DO RIO DE JANEIRO', extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA DO MUNICÍPIO DE SÃO PAULO, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * Para PREFEITURA MUNICIPAL DE TEIXEIRA DE FREITAS, extrair EXCLUSIVAMENTE do campo 'Código de Tributação Nacional' contido em 'SERVIÇO PRESTADO'.\n"
        "  * Para PREFEITURA MUNICIPAL DE MARACANAÚ, extrair EXCLUSIVAMENTE do campo 'CODIGO DA ATIVIDADE/SERVIÇO'.\n"
        "  * Para PREFEITURA DA ESTANCIA TURISTICA DE SAO ROQUE/SP, extrair EXCLUSIVAMENTE do campo 'Código do Serviço'.\n"
        "  * IMPORTANTE: O campo id_cnae deve conter o código do CNAE ou serviço. NUNCA preencha este campo com o CPF, CNPJ ou número de telefone do prestador ou tomador.\n"
        "- descricao_servico: copie e cole o valor textual da discriminação/descrição dos serviços exatamente como aparece no documento, caractere por caractere, sem resumir, reescrever ou parafrasear.\n"
        "- valor_servico: número no padrão brasileiro, sem símbolo de moeda.\n"
        "- aliquota: percentual numérico, por exemplo 5,00.\n"
        "- uf_local_prestacao: duas letras. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- cidade_local_prestacao: nome da cidade. Deve receber as informações que estiverem mais próximas do nome 'PRESTADOR' no documento (evitando confusões com dados do tomador).\n"
        "- iss_retido: SIM ou NÃO.\n"
        "- prefeitura: nome completo da prefeitura emissora em maiúsculas, ex: PREFEITURA MUNICIPAL DE FORTALEZA.\n"
        "- valor_deducoes: valor das deduções em lei no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Deduções') ou parecidos.\n"
        "- descontos_incondicionados: valor dos descontos incondicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Incondicionado') ou parecidos.\n"
        "- descontos_condicionados: valor dos descontos condicionados no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Desconto Condicionado') ou parecidos.\n"
        "- outras_retencoes: outras retenções federais não classificadas, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('Outras Retenções') ou parecidos.\n"
        "- ir: valor do IRRF retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('IR'), parecidos, ou próximo de 'IRRF' ou similares.\n"
        "- pis_nao_retido: valor do PIS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'PIS' ou parecidos.\n"
        "- cofins_nao_retido: valor do COFINS não retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos como 'COFINS' ou parecidos.\n"
        "- csrf: valor da CSRF (CSLL+PIS+COFINS retidos juntos), no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('CSRF'), parecidos, ou próximo de 'Contribuições Sociais Retidas' ou similares.\n"
        "- inss: valor do INSS retido, no padrão brasileiro (use 0,00 se não houver). Note que este valor GERALMENTE, mas não obrigatoriamente, fica próximo de textos com o mesmo nome ('INSS'), parecidos, ou próximo de 'Contribuição Previdenciária Retida' ou similares.\n"
        "Se algum campo realmente não existir no documento, devolva string vazia para textuais ou 0,00 para monetários.\n"
        "O retorno deve ser um JSON válido onde as chaves correspondem exatamente a esses campos informados."
    )

    prompt = (
        "Você é um extrator fiscal. Leia o texto extraído do documento PDF abaixo e retorne apenas JSON válido.\n\n"
        f"{instrucoes_formatacao}\n\n"
        "Seu retorno DEVE ser um objeto JSON contendo exatamente as seguintes chaves:\n"
        "{\n"
        '  "cnpj_prestador": string,\n'
        '  "numero_nf": string,\n'
        '  "data_emissao": string,\n'
        '  "id_cnae": string,\n'
        '  "desc_cnae": string,\n'
        '  "descricao_servico": string,\n'
        '  "uf_local_prestacao": string,\n'
        '  "cidade_local_prestacao": string,\n'
        '  "valor_servico": string,\n'
        '  "aliquota": string,\n'
        '  "iss_retido": string,\n'
        '  "prefeitura": string,\n'
        '  "valor_deducoes": string,\n'
        '  "descontos_incondicionados": string,\n'
        '  "descontos_condicionados": string,\n'
        '  "outras_retencoes": string,\n'
        '  "ir": string,\n'
        '  "pis_nao_retido": string,\n'
        '  "cofins_nao_retido": string,\n'
        '  "csrf": string,\n'
        '  "inss": string\n'
        "}\n\n"
        f"--- INÍCIO DO TEXTO EXTRAÍDO DO PDF ---\n"
        f"{texto_local}\n"
        f"--- FIM DO TEXTO EXTRAÍDO DO PDF ---\n"
        "Retorne APENAS o JSON, sem markdown ou texto explicativo ao redor."
    )

    try:
        chat_completion = cliente.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": "Você é um assistente útil e sempre retorna um objeto JSON perfeitamente válido."
                },
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            model=_obter_modelo_groq(),
            temperature=0,
            response_format={"type": "json_object"},
        )
        
        texto_resposta = chat_completion.choices[0].message.content
        if not texto_resposta:
            return {}

        bruto = json.loads(texto_resposta)
        return _normalizar_resposta(bruto)
    except Exception as exc:
        mensagem = str(exc).upper()
        if "RATE_LIMIT_EXCEEDED" in mensagem or "429" in mensagem or "QUOTA" in mensagem or "LIMIT 100000" in mensagem:
            raise ErroCotaGroq("Limite de cota da API do Groq excedido.") from exc
        print(f"Erro na extração via Groq: {exc}", flush=True)
        return {}

