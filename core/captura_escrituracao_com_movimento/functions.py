import re
from datetime import datetime
from pathlib import Path
import sys, os
from typing import Any, Type, TypeVar


def remover_pontuacao_cpf_cnpj(cnpj: str) -> str:
    """
        Remove toda a pontuação de um CNPJ.
    Args:
        cnpj: string do CNPJ a remover pontuação

    Returns:
        Nova string de *cnpj* com toda a pontuação removida.
    """
    return re.sub(r"[.\-/]", "", cnpj)


def remover_caracteres_escape(texto):
    """Remove todas as ocorrências de caracteres de escape (\n, \t, \r, etc.) de uma string."""
    # Remove caracteres de controle comuns (escape sequences)
    return re.sub(r'[\n\t\r\v\f\b]', '', texto)


def aplicar_mascara_cnpj(numeros: str) -> str:
    n = numeros.strip().replace(".", "")
    len_numeros = len(n)
    if len_numeros != 14:
        raise ValueError(f"A string deve conter exatamente 14 dígitos, ´{n}´ contém {len_numeros}.")

    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


def timestamp_as_file_name(file_extension: str) -> str:
    """
    Return the current timestamp as a file name.

    Args:
        file_extension: The file extension (without the dot)

    Example:
            timestamp_as_file_name("pdf") -> "121629.628.pdf"
    """
    now = datetime.now()
    return f"{now.strftime("%H%M%S.%f")}.{file_extension}"

def resource_path(relative: str) -> str:
    """Resolve caminho de assets tanto no IDE quanto no .exe (PyInstaller)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return str(Path(os.path.join(base, relative)))


def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent  # pasta do .exe
    else:
        return  Path(__file__).parent  # pasta do app.py no IDE


def verifica_arquivo_existe(nome_arquivo: str, diretorio: str) -> str:
    """
    Verifica se um arquivo existe em um diretório específico.

    Args:
        nome_arquivo: Nome do arquivo a ser procurado (ex: 'dados.txt')
        diretorio: Caminho do diretório onde procurar o arquivo

    Returns:
        O caminho completo do arquivo, caso ele exista.

    Raises:
        FileNotFoundError: Se o arquivo não for encontrado no diretório.
    """
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    if not os.path.isfile(caminho_completo):
        raise FileNotFoundError(
            f"Arquivo '{nome_arquivo}' não encontrado no diretório '{diretorio}'"
        )

    return caminho_completo

T = TypeVar("T")

def safe_cast(value: Any, target_type: Type[T]) -> T:
    """
    Casts a value to the target type only if it's not already an instance of that type.

    Args:
        value: The value to potentially cast.
        target_type: The type to cast to.

    Returns:
        The value as the target type, either cast or unchanged.

    Raises:
        TypeError: If the value cannot be cast to the target type.
        ValueError: If the value is incompatible with the target type.
    """
    if isinstance(value, target_type):
        return value
    return target_type(value)
