"""Seletor nativo de certificado digital do repositório do Windows.

Usa `CryptUIDlgSelectCertificateFromStore` (cryptui.dll) — o mesmo diálogo nativo
"Selecionar um Certificado" usado pelo Chrome/Edge — para escolher um certificado
já instalado no repositório "Pessoal" (MY) do usuário atual, sem precisar de
arquivo .pfx nem senha.
"""

import ctypes
import re
from ctypes import wintypes
from typing import Optional

CERT_STORE_PROV_SYSTEM_W = 10
CERT_SYSTEM_STORE_CURRENT_USER = 1 << 16
CERT_HASH_PROP_ID = 3
X509_ASN_ENCODING = 0x00000001
PKCS_7_ASN_ENCODING = 0x00010000


class CERT_CONTEXT(ctypes.Structure):
    _fields_ = [
        ("dwCertEncodingType", wintypes.DWORD),
        ("pbCertEncoded", ctypes.POINTER(ctypes.c_byte)),
        ("cbCertEncoded", wintypes.DWORD),
        ("pCertInfo", ctypes.c_void_p),
        ("hCertStore", ctypes.c_void_p),
    ]


PCCERT_CONTEXT = ctypes.POINTER(CERT_CONTEXT)

_crypt32 = ctypes.WinDLL("crypt32.dll")
_cryptui = ctypes.WinDLL("cryptui.dll")

_crypt32.CertOpenStore.restype = ctypes.c_void_p
_crypt32.CertOpenStore.argtypes = [
    ctypes.c_void_p,  # lpszStoreProvider (aqui: inteiro do provider, castado)
    wintypes.DWORD,
    ctypes.c_void_p,
    wintypes.DWORD,
    ctypes.c_wchar_p,
]

_crypt32.CertCloseStore.restype = wintypes.BOOL
_crypt32.CertCloseStore.argtypes = [ctypes.c_void_p, wintypes.DWORD]

_crypt32.CertFreeCertificateContext.restype = wintypes.BOOL
_crypt32.CertFreeCertificateContext.argtypes = [PCCERT_CONTEXT]

_crypt32.CertGetCertificateContextProperty.restype = wintypes.BOOL
_crypt32.CertGetCertificateContextProperty.argtypes = [
    PCCERT_CONTEXT,
    wintypes.DWORD,
    ctypes.c_void_p,
    ctypes.POINTER(wintypes.DWORD),
]

_cryptui.CryptUIDlgSelectCertificateFromStore.restype = PCCERT_CONTEXT
_cryptui.CryptUIDlgSelectCertificateFromStore.argtypes = [
    ctypes.c_void_p,   # hCertStore
    wintypes.HWND,     # hwnd
    ctypes.c_wchar_p,  # pwszTitle
    ctypes.c_wchar_p,  # pwszDisplayString
    wintypes.DWORD,    # dwDontUseColumn
    wintypes.DWORD,    # dwFlags
    ctypes.c_void_p,   # pvReserved
]


def _obter_thumbprint(cert_context: PCCERT_CONTEXT) -> str:
    tamanho = wintypes.DWORD(0)
    _crypt32.CertGetCertificateContextProperty(cert_context, CERT_HASH_PROP_ID, None, ctypes.byref(tamanho))
    buffer = (ctypes.c_byte * tamanho.value)()
    if not _crypt32.CertGetCertificateContextProperty(cert_context, CERT_HASH_PROP_ID, buffer, ctypes.byref(tamanho)):
        raise OSError("Falha ao obter o thumbprint do certificado selecionado.")
    return bytes(buffer).hex().upper()


def _formatar_documento(digitos: str) -> str:
    if len(digitos) == 14:
        return f"{digitos[0:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:14]}"
    if len(digitos) == 11:
        return f"{digitos[0:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:11]}"
    return digitos


def _extrair_info_certificado(cert_context: PCCERT_CONTEXT) -> dict:
    """Lê o certificado (DER) e extrai campos legíveis para exibir ao operador.

    Certificados ICP-Brasil (e-CNPJ/e-CPF) trazem o CN no formato
    "RAZÃO SOCIAL:CNPJ" ou "NOME:CPF" — aqui isso é separado em campos próprios
    em vez de exibir o Subject X.500 cru (`CN=...,OU=...,O=...`).
    """
    from cryptography import x509
    from cryptography.x509.oid import NameOID

    tamanho = cert_context.contents.cbCertEncoded
    ponteiro = ctypes.cast(cert_context.contents.pbCertEncoded, ctypes.POINTER(ctypes.c_byte * tamanho))
    der_bytes = bytes(bytearray(ponteiro.contents))
    certificado = x509.load_der_x509_certificate(der_bytes)
    subject = certificado.subject

    try:
        cn = subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value
    except IndexError:
        cn = subject.rfc4514_string()

    nome, documento, cnpj = cn, None, None
    if ":" in cn:
        possivel_nome, possivel_doc = cn.rsplit(":", 1)
        digitos = re.sub(r"\D", "", possivel_doc)
        if len(digitos) in (11, 14):
            nome = possivel_nome.strip()
            documento = _formatar_documento(digitos)
            if len(digitos) == 14:
                cnpj = digitos  # CNPJ em dígitos puros — pronto para o campo "CNPJ da Filial"

    tipo_certificado = None
    try:
        for ou in subject.get_attributes_for_oid(NameOID.ORGANIZATIONAL_UNIT_NAME):
            valor = ou.value.lower()
            if "certificado" in valor and ("a1" in valor or "a3" in valor):
                tipo_certificado = ou.value
                break
    except Exception:
        pass

    validade_dt = getattr(certificado, "not_valid_after_utc", None) or certificado.not_valid_after
    validade_ate = validade_dt.strftime("%d/%m/%Y")

    return {
        "nome": nome,
        "documento": documento,
        "cnpj": cnpj,
        "tipo_certificado": tipo_certificado,
        "validade_ate": validade_ate,
        "subject_raw": subject.rfc4514_string(),
    }


def selecionar_certificado_windows(
    titulo: str = "Automações ISS — Selecionar Certificado Digital",
    instrucao: str = "Escolha o certificado digital (A1/A3) que será usado para autenticar no Portal Nacional NFS-e.",
) -> Optional[dict]:
    """Abre o seletor nativo do Windows sobre o repositório 'Pessoal' do usuário atual.

    Retorna um dicionário com o certificado escolhido:
    {"thumbprint", "nome", "documento", "cnpj", "tipo_certificado", "validade_ate", "subject_raw", "exibicao"}
    ("cnpj" vem em dígitos puros, só quando o certificado é e-CNPJ — pronto para
    preencher o campo "CNPJ da Filial"; campos derivados podem vir None quando não
    reconhecidos), ou None se o operador cancelar o diálogo.
    """
    hcertstore = _crypt32.CertOpenStore(
        ctypes.c_void_p(CERT_STORE_PROV_SYSTEM_W),
        0,
        None,
        CERT_SYSTEM_STORE_CURRENT_USER,
        "MY",
    )
    if not hcertstore:
        raise OSError("Não foi possível abrir o repositório de certificados 'Pessoal' do Windows.")

    try:
        cert_context = _cryptui.CryptUIDlgSelectCertificateFromStore(
            hcertstore, None, titulo, instrucao, 0, 0, None
        )
        if not cert_context:
            return None

        try:
            thumbprint = _obter_thumbprint(cert_context)
            info = _extrair_info_certificado(cert_context)

            exibicao = info["nome"]
            if info["documento"]:
                exibicao += f" — {info['documento']}"
            exibicao += f" (válido até {info['validade_ate']})"

            return {"thumbprint": thumbprint, "exibicao": exibicao, **info}
        finally:
            _crypt32.CertFreeCertificateContext(cert_context)
    finally:
        _crypt32.CertCloseStore(hcertstore, 0)
