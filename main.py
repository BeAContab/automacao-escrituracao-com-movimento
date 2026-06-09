"""Ponto de entrada principal do projeto."""

from extrair_nf_pdfs import main
from tratamento_erros import registrar_erro, registrar_evento_execucao


if __name__ == "__main__":
    try:
        registrar_evento_execucao("Programa iniciado pelo main.py", "main.py")
        resultado = main()
    except SystemExit as exc:
        if exc.code not in (0, None):
            registrar_erro(exc, "main.py")
        raise
    except Exception as exc:
        registrar_erro(exc, "main.py")
        raise
    else:
        registrar_evento_execucao(
            f"Programa finalizado com código de saída {resultado}", "main.py"
        )
        raise SystemExit(resultado)
