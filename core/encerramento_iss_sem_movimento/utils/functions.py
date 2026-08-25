import re
import os
import sys
from typing import TypeVar, Type, Any
from datetime import datetime, timedelta
import json
import tomllib
from datetime import date
from .constants import FORMATO_DATA
from loguru import logger


def resource_path(relative_path: str, fallback: str = ".") -> str:
    """Resolve a diferença entre o caminho de um recurso quando o programa é
        executado via script (poetry run ...) ou via executável (.exe).
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(fallback))
    return os.path.join(base_path, relative_path)


def remover_pontuacao_cnpj(cnpj: str) -> str:
    """
        Remove toda a pontuação de um CNPJ.
    Args:
        cnpj: string do CNPJ a remover pontuação

    Returns:
        Nova string de *cnpj* com toda a pontuação removida.
    """
    return re.sub(r"[.\-/]", "", cnpj)


def aplicar_mascara_cnpj(numeros: str) -> str:
    """
    Aplica máscara de CNPJ a uma string de 14 dígitos.
    Formato: XX.XXX.XXX/XXXX-XX
    """
    n = numeros.strip().replace(".", "")
    len_numeros = len(n)
    if len_numeros != 14:
        raise ValueError(f"A string deve conter exatamente 14 dígitos, ´{n}´ contém {len_numeros}.")

    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


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


def timestamp_as_directory(
    usar_mes_competencia: bool = False
) -> str:
    """
    Retorna o mês e ano atual no formato de diretório.
    Se *usar_mes_competencia* for True, o mês será reduzido em 1.

    Example:
            Executando em 29/04/2026
            >>> timestamp_as_directory() # Retorna "2026/04"
            >>> timestamp_as_directory(usar_mes_competencia=True) # Retorna "2026/03"
    """
    now = datetime.now()
    mes = now.month - 1 if usar_mes_competencia else now.month
    return f"{now.year}/{mes:02d}/{now.day:02d}"


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



def validate_day(day: int) -> None:
    """Validates if a given day is within the valid range for a calendar month (1–31).

    Args:
        day: An integer representing the day of the month.

    Raises:
        ValueError: If the day is outside the valid range [1, 31].
    """
    if not 1 <= day <= 31:
        raise ValueError(f"Invalid day: {day}. Day must be between 1 and 31.")


def validate_time(time_str: str) -> None:
    """Validates if a given string is a valid time in HH:MM format.

    Args:
        time_str: A string representing the time in HH:MM format (00:00–23:59).

    Raises:
        ValueError: If the string is not a valid time in HH:MM format.
    """
    if not re.fullmatch(r"\d{2}:\d{2}", time_str):
        raise ValueError(
            f"Invalid time format: '{time_str}'. Expected HH:MM (e.g. '09:30')."
        )

    hours, minutes = int(time_str[:2]), int(time_str[3:])

    if not 0 <= hours <= 23:
        raise ValueError(f"Invalid hours: {hours}. Hours must be between 0 and 23.")

    if not 0 <= minutes <= 59:
        raise ValueError(
            f"Invalid minutes: {minutes}. Minutes must be between 0 and 59."
        )


def extract_fields_from_pydantic_json(json_str: str) -> str:
    """ Extrai o nome dos campos que levantaram ValidationError de um JSON do Pydantic.

    Args:
        json_str: String JSON, retornado pelo metodo ValidationError.json()

    Returns:
        String com os nomes de todos os campos que levantaram erro, separados por ","
    """
    dados = json.loads(json_str)
    locs = [".".join(str(x) for x in item["loc"]) for item in dados]
    return ", ".join(locs)


def get_project_metadata_from_toml(toml_path: str | None = None) -> dict:
    """
    Lê um arquivo pyproject.toml e retorna os metadados da seção [project],
    excluindo campos relacionados a dependências.

    Args:
        toml_path: Caminho para o arquivo pyproject.toml.

    Returns:
        Dicionário com os metadados do projeto, sem chaves de dependências.
    """
    campos_dependencias = {
        "dependencies",
        "optional-dependencies",
        "requires-python",
    }

    with open(toml_path or resource_path("pyproject.toml"), "rb") as f:
        dados = tomllib.load(f)

    metadados_projeto = dados.get("project", {})

    return {
        chave: valor
        for chave, valor in metadados_projeto.items()
        if chave not in campos_dependencias
    }

def get_last_month(day: int = 1) -> date:
    """Retorna a data do último mês no dia `day`.

    :param int day: Data retornada virá configurada com este dia. Padrão é 1
    :return date: Objeto `date`
    """
    validate_day(day)
    
    last_day_last_month = date.today().replace(day=1) - timedelta(days=1)
    try:
        return last_day_last_month.replace(day=day)
    except ValueError as v_err:
        # Caso last_day_last_month não tiver dia `day`. Ex.: meses com menos de 30 dias.
        raise ValueError(
            f"Dia `{day}` inválido para o mês `{last_day_last_month.month}`"
        ) from v_err


def formata_date_para_br(data: date) -> str:
    return data.strftime(FORMATO_DATA)


def safe_division(*valores: float) -> float:
    """Divide uma sequência de números na ordem em que foram declarados.

    Os valores são divididos da esquerda para a direita
    (ex.: valores[0] / valores[1] / valores[2] / ...). Caso qualquer um dos
    valores seja 0, a função retorna 0 em vez de levantar ZeroDivisionError.

    Args:
        *valores: Dois ou mais números a serem divididos, na ordem em que
            devem ser divididos.

    Returns:
        O resultado da divisão de todos os valores da esquerda para a
        direita, ou 0 caso algum dos valores seja 0.

    Raises:
        ValueError: Se forem fornecidos menos de 2 valores.

    Example:
        >>> safe_division(10, 2)
        5.0
        >>> safe_division(100, 5, 2)
        10.0
        >>> safe_division(10, 0)
        0
        >>> safe_division(10, 2, 0, 5)
        0
    """
    if len(valores) < 2:
        raise ValueError(
            f"safe_division() requer ao menos 2 parâmetros, {len(valores)} foram fornecidos."
        )

    if any(valor == 0 for valor in valores):
        logger.warning(f"Divisão por 0 capturada em `safe_division` com os argumentos: {valores}")
        return 0

    resultado = valores[0]
    for divisor in valores[1:]:
        resultado /= divisor

    return resultado
