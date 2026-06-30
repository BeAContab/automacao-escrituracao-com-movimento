"""Utilitarios para registrar erros em arquivo de log."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import traceback


RAIZ_PROJETO = Path(__file__).resolve().parent
PASTAS_LOGS: list[Path] = [RAIZ_PROJETO]


def configurar_pasta_logs(pasta_pdf: Path) -> None:
    """Configura o diretório de logs para a pasta 'log' dentro de pasta_pdf."""
    global PASTAS_LOGS
    pasta_log = pasta_pdf if pasta_pdf.name == "log" else pasta_pdf / "log"
    pasta_log.mkdir(parents=True, exist_ok=True)
    PASTAS_LOGS = [pasta_log]


def configurar_pastas_logs(pastas: list[Path]) -> None:
    """Configura múltiplos diretórios de logs (ex: origem dos PDFs e destino da planilha)."""
    global PASTAS_LOGS
    PASTAS_LOGS = []
    for p in pastas:
        if p:
            pasta_log = p if p.name == "log" else p / "log"
            pasta_log.mkdir(parents=True, exist_ok=True)
            PASTAS_LOGS.append(pasta_log)


def registrar_erro(erro: BaseException, contexto: str, caminhos_log: list[Path] | Path | None = None) -> list[Path] | Path:
    """Registra um erro com contexto, mensagem e stack trace em log_erro.txt."""

    if caminhos_log:
        if isinstance(caminhos_log, Path):
            caminhos = [caminhos_log]
        else:
            caminhos = caminhos_log
    else:
        caminhos = [p / "log_erro.txt" for p in PASTAS_LOGS]

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
    texto = "\n".join(bloco) + "\n"

    for caminho in caminhos:
        try:
            caminho.parent.mkdir(parents=True, exist_ok=True)
            with caminho.open("a", encoding="utf-8") as arquivo:
                arquivo.write(texto)
        except Exception:
            pass
            
    return caminhos[0] if len(caminhos) == 1 else caminhos


def registrar_evento_execucao(
    evento: str,
    contexto: str,
    caminhos_log: list[Path] | Path | None = None,
) -> list[Path] | Path:
    """Registra um marco da execução do programa em log_execucao.txt."""

    if caminhos_log:
        if isinstance(caminhos_log, Path):
            caminhos = [caminhos_log]
        else:
            caminhos = caminhos_log
    else:
        caminhos = [p / "log_execucao.txt" for p in PASTAS_LOGS]

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bloco = [
        "=" * 80,
        f"[{timestamp}] CONTEXTO: {contexto}",
        f"EVENTO: {evento}",
        "",
    ]
    texto = "\n".join(bloco) + "\n"

    for caminho in caminhos:
        try:
            caminho.parent.mkdir(parents=True, exist_ok=True)
            with caminho.open("a", encoding="utf-8") as arquivo:
                arquivo.write(texto)
        except Exception:
            pass
            
    return caminhos[0] if len(caminhos) == 1 else caminhos
