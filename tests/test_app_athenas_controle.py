"""Testes do backend do app para "Importar para Athenas".

NENHUM teste toca arquivo real do ISS nem gera planilha de verdade: `processar_arquivo`
(chamado por `app.athenas_processar_arquivo`) é totalmente dublado.
"""

import json
import re
import shutil
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

import app


class _JanelaFalsa:
    def __init__(self):
        self.js = []

    def evaluate_js(self, codigo):
        self.js.append(re.sub(r"\\u([0-9a-fA-F]{4})", lambda a: chr(int(a.group(1), 16)), codigo))

    def resumo(self):
        for c in self.js:
            if c.startswith("window.mostrar_resumo_athenas("):
                return json.loads(c[len("window.mostrar_resumo_athenas("):-1])
        return None


class _Base(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="app_athenas_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.janela = _JanelaFalsa()
        self.api = app.BeAContabAPI(self.janela)
        self._original_processar = app.athenas_processar_arquivo
        self._original_log_erros = app.athenas_salvar_log_erros

    def tearDown(self):
        app.athenas_processar_arquivo = self._original_processar
        app.athenas_salvar_log_erros = self._original_log_erros

    def logs(self) -> str:
        arquivos = sorted(self.pasta.glob("log_execucao_*.txt"))
        self.assertTrue(arquivos, "nenhum arquivo de log foi criado")
        return "\n".join(a.read_text(encoding="utf-8") for a in arquivos)

    def _esperar_fim(self, ciclos=200):
        for _ in range(ciclos):
            if not self.api._athenas_em_execucao:
                return
            threading.Event().wait(0.02)
        self.fail("a execução não terminou a tempo")

    def _iniciar(self, arquivos, regimes):
        self.api.iniciar_athenas_gui(arquivos, regimes, str(self.pasta))
        self._esperar_fim()


class TestFluxoNormal(_Base):
    def test_processa_cada_arquivo_com_seu_proprio_regime(self):
        chamadas = []
        app.athenas_processar_arquivo = lambda caminho, regime, pasta, log=None: (
            chamadas.append((caminho, regime)) or {"tomados": 2, "prestados": 1}
        )
        self._iniciar(["a.xlsx", "b.xlsx"], ["normal", "simples"])
        self.assertEqual(chamadas, [("a.xlsx", "normal"), ("b.xlsx", "simples")])

    def test_resumo_soma_os_totais_de_todos_os_arquivos(self):
        respostas = iter([{"tomados": 2, "prestados": 1}, {"tomados": 0, "prestados": 3}])
        app.athenas_processar_arquivo = lambda *a, **k: next(respostas)
        self._iniciar(["a.xlsx", "b.xlsx"], ["normal", "normal"])
        resumo = self.janela.resumo()
        self.assertEqual(resumo["tomados"], 2)
        self.assertEqual(resumo["prestados"], 4)
        self.assertEqual(resumo["erros"], 0)
        self.assertIsNone(resumo["caminho_log_erros"])

    def test_flag_de_execucao_e_liberada_ao_terminar(self):
        app.athenas_processar_arquivo = lambda *a, **k: {"tomados": 1, "prestados": 0}
        self._iniciar(["a.xlsx"], ["normal"])
        self.assertFalse(self.api._athenas_em_execucao)

    def test_log_registra_o_total_final(self):
        app.athenas_processar_arquivo = lambda *a, **k: {"tomados": 5, "prestados": 2}
        self._iniciar(["a.xlsx"], ["normal"])
        self.assertIn("TOTAL: 1 empresa(s) | 5 tomados | 2 prestados", self.logs())

    def test_nao_inicia_uma_segunda_importacao_enquanto_a_primeira_roda(self):
        liberar = threading.Event()
        app.athenas_processar_arquivo = lambda *a, **k: (liberar.wait(5), {"tomados": 1, "prestados": 0})[1]

        self.api.iniciar_athenas_gui(["a.xlsx"], ["normal"], str(self.pasta))
        for _ in range(100):
            if self.api._athenas_em_execucao:
                break
            threading.Event().wait(0.02)
        self.assertTrue(self.api._athenas_em_execucao)

        self.janela.js.clear()
        self.api.iniciar_athenas_gui(["b.xlsx"], ["normal"], str(self.pasta))
        self.assertTrue(any("já está em andamento" in c for c in self.janela.js))

        liberar.set()
        self._esperar_fim()


class TestErros(_Base):
    def test_erro_em_um_arquivo_nao_interrompe_os_demais_e_gera_log_de_erros(self):
        caminho_log_erros = self.pasta / "Log_Erros_ISS_Athenas.xlsx"
        app.athenas_salvar_log_erros = lambda erros, pasta: caminho_log_erros

        def falso(caminho, regime, pasta, log=None):
            if caminho == "com_erro.xlsx":
                raise ValueError("Nenhuma aba de Tomados ou Prestados encontrada.")
            return {"tomados": 1, "prestados": 0}

        app.athenas_processar_arquivo = falso
        self._iniciar(["ok.xlsx", "com_erro.xlsx", "ok2.xlsx"], ["normal", "normal", "normal"])

        resumo = self.janela.resumo()
        self.assertEqual(resumo["erros"], 1)
        self.assertEqual(resumo["tomados"], 2)  # os 2 arquivos "ok" processaram normalmente
        self.assertEqual(resumo["caminho_log_erros"], str(caminho_log_erros))
        self.assertEqual([r["arquivo"] for r in resumo["resultados"] if not r["ok"]], ["com_erro.xlsx"])
        self.assertIn("ERRO", self.logs())

    def test_lista_vazia_nao_quebra_e_libera_a_trava(self):
        app.athenas_processar_arquivo = mock.Mock(side_effect=RuntimeError("não deveria ser chamado"))
        self._iniciar([], [])
        app.athenas_processar_arquivo.assert_not_called()
        self.assertFalse(self.api._athenas_em_execucao)
        self.assertIn("TOTAL: 0 empresa(s)", self.logs())


if __name__ == "__main__":
    unittest.main()
