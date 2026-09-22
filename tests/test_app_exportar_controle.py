"""Testes do backend do app para Pausar/Continuar/Parar em "Exportar XML de Prestados".

NENHUM teste abre um navegador nem toca o portal: `executar_exportacao_xml_prestados` é
totalmente dublada (patch em `app.executar_exportacao_xml_prestados`).
"""

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
        self.js.append(codigo)


class _Base(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="app_exportar_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.janela = _JanelaFalsa()
        self.api = app.BeAContabAPI(self.janela)
        self._original = app.executar_exportacao_xml_prestados

    def tearDown(self):
        app.executar_exportacao_xml_prestados = self._original

    def logs(self) -> str:
        arquivos = sorted(self.pasta.glob("log_execucao_*.txt"))
        self.assertTrue(arquivos, "nenhum arquivo de log foi criado")
        return "\n".join(a.read_text(encoding="utf-8") for a in arquivos)

    def _esperar_fim(self, timeout_ciclos=200):
        controle = self.api._controles_xml["exportar"]
        for _ in range(timeout_ciclos):
            if not controle.em_execucao:
                return
            threading.Event().wait(0.03)
        self.fail("a execução não terminou a tempo")


class TestFluxoNormal(_Base):
    def test_callbacks_de_pausa_e_cancelamento_sao_passados_para_a_funcao(self):
        capturado = {}

        def falso(**kwargs):
            capturado.update(kwargs)

        app.executar_exportacao_xml_prestados = falso
        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "09/2026")
        self._esperar_fim()

        self.assertIn("callback_pausa", capturado)
        self.assertIn("callback_cancelamento", capturado)
        controle = self.api._controles_xml["exportar"]
        self.assertEqual(capturado["callback_pausa"], controle.aguardar_se_pausado)
        self.assertEqual(capturado["callback_cancelamento"], controle.cancelado)

    def test_flag_de_execucao_e_liberada_ao_terminar(self):
        app.executar_exportacao_xml_prestados = lambda **kwargs: None
        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "09/2026")
        self._esperar_fim()
        self.assertFalse(self.api._controles_xml["exportar"].em_execucao)

    def test_nao_inicia_uma_segunda_execucao_enquanto_a_primeira_roda(self):
        liberar = threading.Event()
        app.executar_exportacao_xml_prestados = lambda **kwargs: liberar.wait(5)

        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "09/2026")
        for _ in range(100):
            if self.api._controles_xml["exportar"].em_execucao:
                break
            threading.Event().wait(0.02)
        self.assertTrue(self.api._controles_xml["exportar"].em_execucao)

        self.janela.js.clear()
        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "10/2026")
        self.assertTrue(any("já está em andamento" in c for c in self.janela.js))

        liberar.set()
        self._esperar_fim()


class TestPausaEParada(_Base):
    def test_pausar_e_retomar_sao_registrados_no_log_e_o_robo_realmente_espera(self):
        chamadas = []

        def falso(**kwargs):
            kwargs["callback_pausa"]()  # 1ª checagem: ainda não pausado
            self.api.pausar_processamento_xml("exportar")
            fio_continuou = threading.Event()

            def esperar_e_marcar():
                kwargs["callback_pausa"]()  # deve bloquear até "Continuar"
                fio_continuou.set()

            fio = threading.Thread(target=esperar_e_marcar, daemon=True)
            fio.start()
            fio.join(timeout=0.3)
            chamadas.append(("bloqueado_enquanto_pausado", not fio_continuou.is_set()))
            self.api.retomar_processamento_xml("exportar")
            fio.join(timeout=2)
            chamadas.append(("liberou_apos_continuar", fio_continuou.is_set()))

        app.executar_exportacao_xml_prestados = falso
        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "09/2026")
        self._esperar_fim()

        self.assertEqual(dict(chamadas), {"bloqueado_enquanto_pausado": True, "liberou_apos_continuar": True})
        log = self.logs()
        self.assertIn("PAUSA solicitada pelo operador", log)
        self.assertIn("RETOMADA pelo operador", log)

    def test_parar_e_registrado_no_log_e_o_callback_de_cancelamento_reflete_isso(self):
        resultado = {}

        def falso(**kwargs):
            resultado["antes"] = kwargs["callback_cancelamento"]()
            self.api.cancelar_processamento_xml("exportar")
            resultado["depois"] = kwargs["callback_cancelamento"]()

        app.executar_exportacao_xml_prestados = falso
        self.api.iniciar_exportacao_xml_gui(str(self.pasta), "09/2026")
        self._esperar_fim()

        self.assertEqual(resultado, {"antes": False, "depois": True})
        log = self.logs()
        self.assertIn("PARADA solicitada pelo operador", log)
        self.assertIn("navegador não é fechado", log)


if __name__ == "__main__":
    unittest.main()
