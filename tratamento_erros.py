"""Utilitarios para registrar erros em arquivo de log."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import traceback


RAIZ_PROJETO = Path(__file__).resolve().parent
LOG_ERRO_PADRAO = RAIZ_PROJETO / "log_erro.txt"
LOG_EXECUCAO_PADRAO = RAIZ_PROJETO / "log_execucao.txt"


def registrar_erro(erro: BaseException, contexto: str, caminho_log: Path = LOG_ERRO_PADRAO) -> Path:
    """Registra um erro com contexto, mensagem e stack trace no `log_erro.txt`."""

    caminho_log.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    traceback_texto = "".join(
        traceback.format_exception(type(erro), erro, erro.__traceback__)
    ).strip()

    bloco = [
        "=" * 80,
        f"[{timestamp}] CONTEXTO: {contexto}",
        f"TIPO: {erro.__class__.__name__}",
        f"MENSAGEM: {erro}",
        "TRACEBACK:",
        traceback_texto or "(sem stack trace disponível)",
        "",
    ]
    with caminho_log.open("a", encoding="utf-8") as arquivo:
        arquivo.write("\n".join(bloco))
        arquivo.write("\n")
    return caminho_log


def registrar_evento_execucao(
    evento: str,
    contexto: str,
    caminho_log: Path = LOG_EXECUCAO_PADRAO,
) -> Path:
    """Registra um marco da execução do programa em `log_execucao.txt`."""

    caminho_log.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bloco = [
        "=" * 80,
        f"[{timestamp}] CONTEXTO: {contexto}",
        f"EVENTO: {evento}",
        "",
    ]
    with caminho_log.open("a", encoding="utf-8") as arquivo:
        arquivo.write("\n".join(bloco))
        arquivo.write("\n")
    return caminho_log
