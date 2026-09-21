import shutil
import tempfile
import threading
import unittest
from pathlib import Path

import app
from tests.test_escrituracao_multi_cnpj import AcessoFalso, _linhas, _planilha, CNPJ_A, CNPJ_B


class _JanelaFalsa:
    def __init__(self):
        self.js = []

    def evaluate_js(self, codigo):
        self.js.append(codigo)


class TestControleEscrituracaoMulti(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="escmulti_app_"))
        self.planilha = _planilha(
            self.pasta,
            _linhas(CNPJ_A, "11111111111", "SenhaA", 2) + _linhas(CNPJ_B, "22222222222", "SenhaB", 2),
        )
        self.janela = _JanelaFalsa()
        self.api = app.BeAContabAPI(self.janela)
        self._original = app.AcessoIssSelenium

    def tearDown(self):
        app.AcessoIssSelenium = self._original
        shutil.rmtree(self.pasta, ignore_errors=True)

    def _log_em_arquivo(self) -> str:
        arquivos = list(self.pasta.glob("log_execucao_*.txt"))
        self.assertEqual(len(arquivos), 1)
        return arquivos[0].read_text(encoding="utf-8")

    def _rodar(self, ao_logar):
        acesso = AcessoFalso(ao_logar=ao_logar)
        app.AcessoIssSelenium = lambda **_: acesso
        controle = self.api._controles_xml["esc_multi"]
        controle.reiniciar()
        controle.em_execucao = True
        return acesso, controle

    def test_pausa_no_meio_e_avisada_no_log_da_tela_e_do_arquivo_e_retoma(self):
        def ao_logar(login):
            if login == "11111111111":  # operador clica em Pausar durante a 1ª empresa
                self.api.pausar_processamento_xml("esc_multi")

        acesso, controle = self._rodar(ao_logar)
        fio = threading.Thread(target=self.api._executar_escrituracao_multi, args=(str(self.planilha),))
        fio.start()
        fio.join(timeout=1.0)
        self.assertTrue(fio.is_alive(), "deveria estar pausado antes da 2ª empresa")
        self.assertNotIn(("login", "22222222222"), acesso.chamadas)

        self.api.retomar_processamento_xml("esc_multi")
        fio.join(timeout=5)
        self.assertFalse(fio.is_alive())
        self.assertIn(("login", "22222222222"), acesso.chamadas)

        log = self._log_em_arquivo()
        self.assertIn("PAUSA solicitada pelo operador", log)
        self.assertIn("Execução PAUSADA no ponto em que estava", log)
        self.assertIn("RETOMADA pelo operador", log)
        self.assertTrue(any("PAUSA solicitada" in c for c in self.janela.js))  # também na tela
        self.assertFalse(controle.em_execucao)

    def test_pausa_na_ultima_empresa_avisa_que_nao_chegou_a_valer(self):
        def ao_logar(login):
            if login == "22222222222":
                self.api.pausar_processamento_xml("esc_multi")

        self._rodar(ao_logar)
        self.api._executar_escrituracao_multi(str(self.planilha))
        log = self._log_em_arquivo()
        self.assertIn("PAUSA solicitada pelo operador", log)
        self.assertIn("A pausa solicitada não chegou a valer", log)

    def test_parada_e_registrada_e_cancela_as_demais(self):
        def ao_logar(login):
            self.api.cancelar_processamento_xml("esc_multi")

        acesso, _ = self._rodar(ao_logar)
        self.api._executar_escrituracao_multi(str(self.planilha))
        log = self._log_em_arquivo()
        self.assertIn("PARADA solicitada pelo operador", log)
        self.assertIn("Execução cancelada", log)
        self.assertNotIn(("login", "22222222222"), acesso.chamadas)

    def test_sem_execucao_em_andamento_nao_registra_nada(self):
        self.api.pausar_processamento_xml("esc_multi")
        self.api.retomar_processamento_xml("esc_multi")
        self.api.cancelar_processamento_xml("esc_multi")
        self.assertEqual(self.janela.js, [])

    def test_cliques_repetidos_nao_duplicam_o_aviso(self):
        avisos = []
        controle = self.api._controles_xml["esc_multi"]
        controle.reiniciar()
        controle.log = lambda msg, erro=False: avisos.append(msg)
        self.api.pausar_processamento_xml("esc_multi")
        self.api.pausar_processamento_xml("esc_multi")
        self.api.retomar_processamento_xml("esc_multi")
        self.api.retomar_processamento_xml("esc_multi")
        self.assertEqual(len(avisos), 2)


if __name__ == "__main__":
    unittest.main()
