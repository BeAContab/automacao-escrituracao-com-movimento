"""Download de NFS-e do Portal Nacional (API do ADN) via mTLS com certificado digital.

Réplica em Python da lógica da extensão Chrome "baixar-nfse-portal-nacional"
(background.js), sem depender de navegador. Suporta dois modos de autenticação:

- **Certificado instalado no Windows** (recomendado): usa o `curl.exe` nativo do
  Windows (backend Schannel) referenciando o certificado por thumbprint
  (`CurrentUser\\MY\\<thumbprint>`) — a chave privada nunca sai do repositório do
  Windows/CNG, funcionando inclusive para certificados A3 (token/smartcard).
- **Arquivo .pfx/.p12 + senha** (alternativo): carrega a chave via `cryptography`
  e autentica via `requests`/OpenSSL — usado quando o certificado não está
  importado no repositório do Windows.
"""

import base64
import gzip
import json
import os
import re
import shutil
import stat
import subprocess
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Optional

import requests
from cryptography.hazmat.primitives.serialization import (
    Encoding,
    NoEncryption,
    PrivateFormat,
    pkcs12,
)
from pydantic import SecretStr

BASE_URL = "https://adn.nfse.gov.br/contribuintes/DFe"
MAX_TENTATIVAS = 5
DELAY_INICIAL_S = 1.0
FATOR_BACKOFF = 2
TIMEOUT_REQUISICAO_S = 15

# Caminho absoluto do curl nativo do Windows (backend Schannel, suporta certificado
# do repositório do Windows). Nunca usar "curl" solto: em máquinas com Git for Windows
# instalado, o curl do PATH costuma ser o do MinGW (backend OpenSSL), que não suporta
# a sintaxe "CurrentUser\MY\<thumbprint>".
CURL_PATH = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "curl.exe")

RE_EMIT_CNPJ = re.compile(r"<emit>.*?<CNPJ>(\d{14})</CNPJ>", re.DOTALL)
RE_TOMA_CNPJ = re.compile(r"<toma>.*?<CNPJ>(\d{14})</CNPJ>", re.DOTALL)
RE_DATA_EMISSAO = re.compile(r"<dhEmi>(\d{4})-(\d{2})")


class CancelamentoSolicitado(Exception):
    """Sinaliza que o operador pediu para cancelar o download em andamento."""


def _log_padrao(msg: str, is_error: bool = False) -> None:
    print(f"[ERRO] {msg}" if is_error else msg)


def _progresso_padrao(pct: int, status: str) -> None:
    print(f"[{pct}%] {status}")


@dataclass
class RespostaHTTP:
    status_code: int
    texto: Optional[str]
    retry_after: Optional[str] = None


# ----------------------------------------------------------------
# Modo 1 (recomendado): certificado instalado no repositório do Windows,
# autenticando via curl.exe nativo (Schannel) por thumbprint.
# ----------------------------------------------------------------
class _ClienteCurlSchannel:
    """Autentica via certificado do repositório do Windows (thumbprint), delegando
    toda a operação de chave privada ao CNG do Windows — a chave nunca é extraída,
    o que também permite usar certificados A3 (token/smartcard) com PIN nativo."""

    def __init__(self, thumbprint: str):
        if not os.path.exists(CURL_PATH):
            raise RuntimeError(
                f"curl.exe nativo do Windows não encontrado em {CURL_PATH}. "
                "Verifique se o Windows está atualizado (curl é nativo desde o Windows 10 1803+)."
            )
        self.thumbprint = thumbprint

    def get(self, url: str) -> RespostaHTTP:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".body") as arquivo_corpo:
            caminho_corpo = Path(arquivo_corpo.name)

        try:
            argumentos = [
                CURL_PATH,
                "--cert", f"CurrentUser\\MY\\{self.thumbprint}",
                "-sS",
                "--max-time", str(TIMEOUT_REQUISICAO_S),
                "-o", str(caminho_corpo),
                "-w", "%{http_code}|%header{retry-after}",
                url,
            ]
            resultado = subprocess.run(
                argumentos,
                capture_output=True,
                text=True,
                timeout=TIMEOUT_REQUISICAO_S + 10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if resultado.returncode != 0:
                raise RuntimeError(
                    f"Falha ao autenticar/conectar via curl (código {resultado.returncode}): "
                    f"{resultado.stderr.strip() or 'sem detalhe do curl'}"
                )

            status_str, _, retry_after = resultado.stdout.strip().partition("|")
            texto = caminho_corpo.read_text(encoding="utf-8", errors="replace") if caminho_corpo.exists() else None
            return RespostaHTTP(
                status_code=int(status_str) if status_str.isdigit() else 0,
                texto=texto,
                retry_after=retry_after or None,
            )
        finally:
            caminho_corpo.unlink(missing_ok=True)


# ----------------------------------------------------------------
# Modo 2 (alternativo): certificado em arquivo .pfx/.p12 + senha.
# ----------------------------------------------------------------
def _carregar_certificado_pfx(caminho_pfx: Path, senha_certificado: SecretStr):
    """Carrega chave privada e certificado de um arquivo .pfx/.p12 (certificado A1)."""
    dados_pfx = caminho_pfx.read_bytes()
    chave_privada, certificado, cadeia_adicional = pkcs12.load_key_and_certificates(
        dados_pfx, senha_certificado.get_secret_value().encode("utf-8")
    )
    if chave_privada is None or certificado is None:
        raise ValueError("Não foi possível extrair a chave privada/certificado do arquivo .pfx informado.")
    return chave_privada, certificado, cadeia_adicional


def _escrever_pem_temporarios(chave_privada, certificado, cadeia_adicional) -> tuple[Path, Path, Path]:
    """Escreve cert+chave em arquivos PEM temporários (requests exige caminho de arquivo,
    não aceita bytes em memória). Retorna (cert_path, key_path, pasta_temporaria).

    A pasta retornada deve ser removida pelo chamador assim que a sessão HTTP não for
    mais necessária — a chave privada fica em texto plano em disco enquanto ela existir.
    """
    pasta_tmp = Path(tempfile.mkdtemp(prefix="beacontab_mtls_"))
    try:
        os.chmod(pasta_tmp, stat.S_IRWXU)
    except Exception:
        pass

    cert_pem = certificado.public_bytes(Encoding.PEM)
    if cadeia_adicional:
        for cert_intermediario in cadeia_adicional:
            cert_pem += cert_intermediario.public_bytes(Encoding.PEM)
    cert_path = pasta_tmp / "cert.pem"
    cert_path.write_bytes(cert_pem)

    key_pem = chave_privada.private_bytes(Encoding.PEM, PrivateFormat.PKCS8, NoEncryption())
    key_path = pasta_tmp / "key.pem"
    key_path.write_bytes(key_pem)
    try:
        os.chmod(key_path, stat.S_IRUSR | stat.S_IWUSR)
    except Exception:
        pass

    return cert_path, key_path, pasta_tmp


class _ClienteRequestsPfx:
    """Modo alternativo: certificado em arquivo .pfx/.p12 + senha, autenticando via
    requests/OpenSSL. Usado quando o certificado não está importado no Windows."""

    def __init__(self, caminho_certificado_pfx: str, senha_certificado: SecretStr):
        chave_privada, certificado, cadeia_adicional = _carregar_certificado_pfx(
            Path(caminho_certificado_pfx), senha_certificado
        )
        cert_path, key_path, self.pasta_pem = _escrever_pem_temporarios(chave_privada, certificado, cadeia_adicional)
        self.session = requests.Session()
        self.session.cert = (str(cert_path), str(key_path))

    def get(self, url: str) -> RespostaHTTP:
        try:
            resposta = self.session.get(url, timeout=TIMEOUT_REQUISICAO_S)
        except requests.RequestException as erro:
            raise ConnectionError(str(erro)) from erro
        return RespostaHTTP(
            status_code=resposta.status_code,
            texto=resposta.text,
            retry_after=resposta.headers.get("Retry-After"),
        )

    def fechar(self) -> None:
        self.session.close()
        shutil.rmtree(self.pasta_pem, ignore_errors=True)


# ----------------------------------------------------------------
# Chamada à API do ADN, com retry/backoff
# ----------------------------------------------------------------
def _montar_url(nsu: int, cnpj_filial: Optional[str]) -> str:
    url = f"{BASE_URL}/{nsu}"
    if cnpj_filial:
        url += f"?cnpjConsulta={cnpj_filial}"
    return url


# Valores de StatusProcessamento que a API do ADN usa para sinalizar "não há mais
# documentos a partir deste NSU" — não é um erro, é o fim natural da varredura.
# "SEM_DOCUMENTOS" é documentado; "NENHUM_DOCUMENTO_LOCALIZADO" (visto em produção,
# retornado com HTTP 404 e código de erro E2220) tem o mesmo significado.
STATUS_SEM_MAIS_DOCUMENTOS = {"SEM_DOCUMENTOS", "NENHUM_DOCUMENTO_LOCALIZADO"}


def _interpretar_resposta(corpo_json: Optional[dict]) -> dict:
    if not corpo_json:
        return {"sem_mais_documentos": True}

    if corpo_json.get("StatusProcessamento") in STATUS_SEM_MAIS_DOCUMENTOS:
        return {"sem_mais_documentos": True}

    lote = corpo_json.get("LoteDFe")
    if not lote:
        return {"sem_mais_documentos": True}

    nsus_do_lote = [doc.get("NSU", 0) for doc in lote if isinstance(doc.get("NSU", 0), int)]
    maior_nsu = max(nsus_do_lote) if nsus_do_lote else 0
    return {
        "sem_mais_documentos": False,
        "proximo_nsu": maior_nsu + 1,
        "documentos": lote,
    }


def _buscar_lote_com_retry(
    cliente,
    nsu: int,
    cnpj_filial: Optional[str],
    log: Callable[[str, bool], None],
    callback_cancelamento: Optional[Callable[[], bool]],
) -> dict:
    tentativa = 0
    delay = DELAY_INICIAL_S

    while tentativa < MAX_TENTATIVAS:
        if callback_cancelamento and callback_cancelamento():
            raise CancelamentoSolicitado()

        try:
            resposta = cliente.get(_montar_url(nsu, cnpj_filial))
        except (ConnectionError, RuntimeError, TimeoutError) as erro:
            tentativa += 1
            if tentativa >= MAX_TENTATIVAS:
                raise
            log(f"Falha de conexão no NSU {nsu}: {erro}. Tentativa {tentativa}/{MAX_TENTATIVAS}...", True)
            time.sleep(delay)
            delay *= FATOR_BACKOFF
            continue

        if resposta.status_code == 204:
            return {"sem_mais_documentos": True}

        if resposta.status_code == 429:
            tentativa += 1
            if tentativa >= MAX_TENTATIVAS:
                raise RuntimeError(f"Limite de requisições (HTTP 429) excedido após {MAX_TENTATIVAS} tentativas no NSU {nsu}.")
            espera = float(resposta.retry_after) if resposta.retry_after else delay
            log(f"Limite de requisições (HTTP 429) no NSU {nsu}. Aguardando {espera:.0f}s antes de retomar...", True)
            time.sleep(espera)
            delay = max(delay * FATOR_BACKOFF, espera)
            continue

        if resposta.status_code in (500, 503):
            tentativa += 1
            log(f"API instável (HTTP {resposta.status_code}) no NSU {nsu}. Tentativa {tentativa}/{MAX_TENTATIVAS}...", True)
            time.sleep(delay)
            delay *= FATOR_BACKOFF
            continue

        try:
            corpo_json = json.loads(resposta.texto) if resposta.texto else None
        except json.JSONDecodeError:
            corpo_json = None

        if not (200 <= resposta.status_code < 300):
            # A API do ADN às vezes sinaliza "fim da fila" com um status de erro HTTP
            # (ex.: 404 + StatusProcessamento=NENHUM_DOCUMENTO_LOCALIZADO) em vez de um
            # 204/StatusProcessamento=SEM_DOCUMENTOS "limpo" — trata isso como fim normal.
            if corpo_json and corpo_json.get("StatusProcessamento") in STATUS_SEM_MAIS_DOCUMENTOS:
                return {"sem_mais_documentos": True}
            raise RuntimeError(f"HTTP {resposta.status_code} inesperado no NSU {nsu}: {(resposta.texto or '')[:300]}")

        return _interpretar_resposta(corpo_json)

    raise RuntimeError(f"Número máximo de tentativas ({MAX_TENTATIVAS}) atingido para o NSU {nsu}.")


# ----------------------------------------------------------------
# Decodificação e organização dos documentos
# ----------------------------------------------------------------
def _decodificar_documento(doc: dict) -> Optional[str]:
    payload_base64 = doc.get("ArquivoXml") or doc.get("xmlNFSe") or doc.get("XML")
    if not payload_base64:
        return None

    bruto = base64.b64decode(payload_base64)
    if bruto[:2] == b"\x1f\x8b":
        bruto = gzip.decompress(bruto)
    return bruto.decode("utf-8")


def _determinar_subpasta(xml_texto: str, tipo_doc: str, cnpj_filial: Optional[str]) -> str:
    if tipo_doc == "evento":
        return "Eventos"
    cnpj_emit = RE_EMIT_CNPJ.search(xml_texto)
    cnpj_toma = RE_TOMA_CNPJ.search(xml_texto)
    if cnpj_filial and cnpj_emit and cnpj_emit.group(1) == cnpj_filial:
        return "Notas Emitidas"
    if cnpj_filial and cnpj_toma and cnpj_toma.group(1) == cnpj_filial:
        return "Notas Tomadas"
    return "Outras"


def _extrair_ano_mes(xml_texto: str) -> tuple[Optional[str], Optional[str]]:
    match_data = RE_DATA_EMISSAO.search(xml_texto)
    if match_data:
        return match_data.group(1), match_data.group(2)
    return None, None


def _organizar_e_salvar(
    xml_texto: str, nsu: int, tipo_doc: str, subpasta: str, cnpj_filial: Optional[str],
    pasta_destino: Path, ano_doc: Optional[str], mes_doc: Optional[str],
) -> Path:
    cnpj_pasta = cnpj_filial or "sem_cnpj"
    pasta_ano_mes = f"{ano_doc}/{mes_doc}" if ano_doc and mes_doc else "Sem_Data"

    pasta_final = pasta_destino / cnpj_pasta / subpasta / pasta_ano_mes
    pasta_final.mkdir(parents=True, exist_ok=True)

    caminho_arquivo = pasta_final / f"nsu_{nsu}_{tipo_doc}.xml"
    caminho_arquivo.write_text(xml_texto, encoding="utf-8")
    return caminho_arquivo


def _processar_lote(
    documentos: list,
    cnpj_filial: Optional[str],
    pasta_destino: Path,
    log: Callable[[str, bool], None],
    apenas_notas_tomadas: bool = False,
    ano_filtro: Optional[int] = None,
    mes_filtro: Optional[int] = None,
) -> int:
    contador = 0
    for doc in documentos:
        try:
            xml_texto = _decodificar_documento(doc)
            if not xml_texto:
                continue
            nsu_doc = doc.get("NSU", 0)
            tipo_doc = str(doc.get("TipoDocumento", "NFSE")).lower()
            subpasta = _determinar_subpasta(xml_texto, tipo_doc, cnpj_filial)
            if apenas_notas_tomadas and subpasta != "Notas Tomadas":
                continue

            ano_doc, mes_doc = _extrair_ano_mes(xml_texto)
            if ano_filtro and ano_doc != f"{ano_filtro:04d}":
                continue
            if mes_filtro and mes_doc != f"{mes_filtro:02d}":
                continue

            _organizar_e_salvar(xml_texto, nsu_doc, tipo_doc, subpasta, cnpj_filial, pasta_destino, ano_doc, mes_doc)
            contador += 1
        except Exception as erro:
            log(f"Falha ao processar um documento do lote: {erro}", True)
    return contador


RE_NSU_NO_NOME_ARQUIVO = re.compile(r"^nsu_(\d+)_", re.IGNORECASE)


def _todos_nsus_na_pasta(pasta_destino: Path) -> set[int]:
    """Varre recursivamente `pasta_destino` e retorna o conjunto de todos os NSUs
    já baixados, lidos do padrão de nome `nsu_<N>_<tipo>.xml` (mesma nomenclatura
    usada por esta função e pela extensão Chrome irmã)."""
    nsus: set[int] = set()
    if not pasta_destino.exists():
        return nsus
    for caminho_arquivo in pasta_destino.rglob("nsu_*.xml"):
        match = RE_NSU_NO_NOME_ARQUIVO.match(caminho_arquivo.name)
        if match:
            nsus.add(int(match.group(1)))
    return nsus


def detectar_maior_nsu_na_pasta(pasta_destino: str) -> Optional[int]:
    """Retorna o maior NSU já baixado encontrado em `pasta_destino`, ou None se a
    pasta estiver vazia/sem nenhum arquivo reconhecível."""
    nsus = _todos_nsus_na_pasta(Path(pasta_destino))
    return max(nsus) if nsus else None


def detectar_nsu_para_retomar(pasta_destino: str, cnpj_filial: Optional[str] = None) -> Optional[int]:
    """Detecta o NSU mais adequado para retomar a varredura a partir do conteúdo
    real da pasta — considerando tanto o caso normal (continuar depois do maior
    NSU já baixado) quanto o caso de um "buraco" na sequência, deixado por arquivos
    apagados manualmente (ex.: operador apagou os XMLs de um ano inteiro para
    baixá-los de novo). O buraco pode estar em qualquer ponto — inclusive bem no
    início do que já foi baixado (o ano mais antigo) — não só entre dois arquivos
    ainda presentes: por isso o limite inferior da varredura de lacunas usa o menor
    NSU que a API já retornou de fato para este CNPJ em qualquer execução anterior
    (`obter_primeiro_nsu_visto`), não apenas o menor NSU ainda presente na pasta
    (que sozinho não distingue "nunca existiu antes disso" de "existia e foi
    apagado"). Sem esse histórico (ex.: primeira vez usando esta função), cai de
    volta no menor NSU presente na pasta — mais conservador, mas sem falso positivo.

    Como a varredura da API é sequencial e contínua, retomar a partir da PRIMEIRA
    lacuna encontrada é suficiente para recuperar TODAS as lacunas seguintes numa
    única passada (não é preciso localizar cada uma individualmente) — a varredura
    volta a re-percorrer, sem problema, os NSUs mais novos que já existem (os
    arquivos são só sobrescritos com o mesmo conteúdo) até alcançar de novo o fim
    real da fila. Sem nenhuma lacuna, retorna o maior NSU + 1 (continuação normal).
    Retorna None se a pasta não tiver nenhum arquivo reconhecível.

    Uso: só como SUGESTÃO para a GUI pré-preencher o campo "NSU inicial" — o
    operador pode sempre sobrescrever. `executar_download_nfse_nacional` NUNCA
    chama esta função por conta própria nem usa o resultado dela para sobrescrever
    o NSU informado."""
    nsus = _todos_nsus_na_pasta(Path(pasta_destino))
    if not nsus:
        return None

    menor_na_pasta = min(nsus)
    primeiro_visto = obter_primeiro_nsu_visto(cnpj_filial)
    limite_inferior = min(primeiro_visto, menor_na_pasta) if primeiro_visto is not None else menor_na_pasta
    maior = max(nsus)
    for candidato in range(limite_inferior, maior + 1):
        if candidato not in nsus:
            return candidato
    return maior + 1


# ----------------------------------------------------------------
# Persistência do último NSU processado (retomada de varredura)
# ----------------------------------------------------------------
def _caminho_estado() -> Path:
    appdata = os.environ.get("APPDATA", str(Path.home()))
    caminho = Path(appdata) / "BeAContab" / "brain" / "nfse_nacional_estado.json"
    caminho.parent.mkdir(parents=True, exist_ok=True)
    return caminho


def _chave_estado(cnpj_filial: Optional[str]) -> str:
    return cnpj_filial or "_geral"


def obter_ultimo_nsu_salvo(cnpj_filial: Optional[str]) -> Optional[int]:
    caminho = _caminho_estado()
    if not caminho.exists():
        return None
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    registro = estado.get(_chave_estado(cnpj_filial))
    return registro.get("ultimo_nsu") if registro else None


def obter_ultimo_certificado_salvo() -> Optional[dict]:
    """Retorna o dicionário completo (thumbprint, nome, documento, validade_ate, exibicao, ...)
    do último certificado do Windows usado com sucesso, para a GUI pré-selecionar
    (nenhum desses dados é segredo — são só metadados públicos do certificado)."""
    caminho = _caminho_estado()
    if not caminho.exists():
        return None
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    return estado.get("_ultimo_certificado")


def salvar_ultimo_certificado(certificado: dict) -> None:
    """Persiste o dicionário retornado por `windows_certstore.selecionar_certificado_windows`."""
    caminho = _caminho_estado()
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8")) if caminho.exists() else {}
    except (json.JSONDecodeError, OSError):
        estado = {}
    estado["_ultimo_certificado"] = certificado
    try:
        caminho.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def salvar_ultimo_nsu(cnpj_filial: Optional[str], nsu: int) -> None:
    caminho = _caminho_estado()
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8")) if caminho.exists() else {}
    except (json.JSONDecodeError, OSError):
        estado = {}
    registro = estado.get(_chave_estado(cnpj_filial), {})
    registro["ultimo_nsu"] = nsu
    registro["atualizado_em"] = datetime.now(timezone.utc).isoformat()
    estado[_chave_estado(cnpj_filial)] = registro
    try:
        caminho.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def atualizar_primeiro_nsu_visto(cnpj_filial: Optional[str], nsu_visto: int) -> None:
    """Atualiza (só para baixo) o menor NSU já visto de fato retornado pela API para
    este CNPJ, em qualquer execução. Diferente de "a partir de qual NSU o operador
    pediu para começar" (que normalmente é 0, mesmo que o primeiro documento real
    do CNPJ esteja bem mais à frente, já que o NSU é uma numeração global do ADN) —
    isto é o que `detectar_nsu_para_retomar` precisa para diferenciar corretamente
    "nunca existiu nada antes disso" de "existia e foi apagado da pasta"."""
    caminho = _caminho_estado()
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8")) if caminho.exists() else {}
    except (json.JSONDecodeError, OSError):
        estado = {}
    chave = _chave_estado(cnpj_filial)
    registro = estado.get(chave, {})
    atual = registro.get("primeiro_nsu_visto")
    if atual is None or nsu_visto < atual:
        registro["primeiro_nsu_visto"] = nsu_visto
        estado[chave] = registro
        try:
            caminho.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass


def obter_primeiro_nsu_visto(cnpj_filial: Optional[str]) -> Optional[int]:
    caminho = _caminho_estado()
    if not caminho.exists():
        return None
    try:
        estado = json.loads(caminho.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    registro = estado.get(_chave_estado(cnpj_filial))
    return (registro or {}).get("primeiro_nsu_visto")


# ----------------------------------------------------------------
# Função principal
# ----------------------------------------------------------------
def executar_download_nfse_nacional(
    pasta_destino: str,
    nsu_inicial: int = 0,
    cnpj_filial: Optional[str] = None,
    apenas_notas_tomadas: bool = False,
    ano_filtro: Optional[int] = None,
    mes_filtro: Optional[int] = None,
    certificado_thumbprint: Optional[str] = None,
    caminho_certificado_pfx: Optional[str] = None,
    senha_certificado: Optional[SecretStr] = None,
    callback_log: Optional[Callable[[str, bool], None]] = None,
    callback_progresso: Optional[Callable[[int, str], None]] = None,
    callback_cancelamento: Optional[Callable[[], bool]] = None,
) -> dict:
    """Varre a API pública do ADN (adn.nfse.gov.br) via mTLS, baixando e organizando
    os XMLs de NFS-e por CNPJ/tipo/ano-mês.

    Autenticação: informe `certificado_thumbprint` (certificado já instalado no
    repositório do Windows — modo recomendado, ver `windows_certstore.py`) OU
    `caminho_certificado_pfx`+`senha_certificado` (arquivo .pfx/.p12 — modo alternativo).

    `apenas_notas_tomadas`: quando True, só grava em disco os documentos identificados
    como "Notas Tomadas" (onde `cnpj_filial` é o tomador do serviço) — os demais
    (Notas Emitidas, Outras, Eventos) são descartados sem serem salvos.

    `ano_filtro`/`mes_filtro` (opcionais): só grava em disco os documentos cuja data
    de emissão (`dhEmi`) bate com o ano/mês informado. IMPORTANTE: a API do ADN não
    tem parâmetro de consulta por data — só por NSU — então isto NÃO acelera nem
    reduz a varredura em si; o download continua percorrendo sequencialmente todos
    os NSUs a partir do ponto de início (todo o histórico, se `nsu_inicial=0`),
    baixando e decodificando cada documento normalmente, só descartando sem salvar
    os que não batem com o período pedido. Útil para reprocessar/conferir um mês
    específico sem misturar XMLs de outros períodos na pasta, não para economizar
    tempo de download.

    Em ambos os filtros acima, o NSU continua avançando normalmente sobre TODOS os
    documentos do lote recebido da API, filtrados ou não — são filtros só do que é
    persistido em disco, não do que é varrido (não afetam a detecção de lacunas).

    Retorna um resumo: {"total_notas": int, "total_lotes": int, "ultimo_nsu": int}.
    """
    log = callback_log or _log_padrao
    progresso = callback_progresso or _progresso_padrao
    cnpj_filial = cnpj_filial or None

    pasta_destino_path = Path(pasta_destino)
    pasta_destino_path.mkdir(parents=True, exist_ok=True)

    # O NSU inicial informado pelo chamador é sempre respeitado ao pé da letra — nunca
    # é sobrescrito automaticamente aqui (ver `detectar_nsu_para_retomar`, que existe
    # só como sugestão para a GUI, nunca aplicada por conta própria dentro desta função).
    nsu_atual = max(0, int(nsu_inicial))

    total_notas = 0
    total_lotes = 0
    cliente = None

    try:
        progresso(0, "Preparando autenticação com o certificado digital...")
        if certificado_thumbprint:
            log("Autenticando com certificado do repositório do Windows...")
            cliente = _ClienteCurlSchannel(certificado_thumbprint)
        elif caminho_certificado_pfx:
            log("Autenticando com certificado em arquivo (.pfx)...")
            cliente = _ClienteRequestsPfx(caminho_certificado_pfx, senha_certificado)
        else:
            raise ValueError("Informe um certificado: selecione um instalado no Windows ou um arquivo .pfx.")

        log(f"Iniciando varredura a partir do NSU {nsu_atual}...")
        progresso(1, "Iniciando varredura...")

        while True:
            if callback_cancelamento and callback_cancelamento():
                log("Download cancelado pelo operador.", True)
                break

            resultado = _buscar_lote_com_retry(cliente, nsu_atual, cnpj_filial, log, callback_cancelamento)

            if resultado.get("sem_mais_documentos"):
                log("Varredura concluída. Nenhum documento novo encontrado.")
                break

            nsus_do_lote = [doc.get("NSU") for doc in resultado["documentos"] if isinstance(doc.get("NSU"), int)]
            if nsus_do_lote:
                atualizar_primeiro_nsu_visto(cnpj_filial, min(nsus_do_lote))

            notas_processadas = _processar_lote(
                resultado["documentos"], cnpj_filial, pasta_destino_path, log,
                apenas_notas_tomadas, ano_filtro, mes_filtro,
            )
            total_notas += notas_processadas
            total_lotes += 1
            nsu_atual = resultado["proximo_nsu"]

            salvar_ultimo_nsu(cnpj_filial, nsu_atual)
            log(f"Processando NSU: {nsu_atual} — {total_notas} nota(s) encontrada(s) em {total_lotes} lote(s).")
            progresso(min(99, total_lotes), f"NSU {nsu_atual} — {total_notas} nota(s) baixada(s)")

        return {"total_notas": total_notas, "total_lotes": total_lotes, "ultimo_nsu": nsu_atual}

    except CancelamentoSolicitado:
        log("Download cancelado pelo operador.", True)
        return {"total_notas": total_notas, "total_lotes": total_lotes, "ultimo_nsu": nsu_atual}
    finally:
        salvar_ultimo_nsu(cnpj_filial, nsu_atual)
        if isinstance(cliente, _ClienteRequestsPfx):
            cliente.fechar()
