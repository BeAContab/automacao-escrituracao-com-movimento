"""Testes do backend do app para a Escrituração Multi-CNPJ (modos validar/simular/escriturar).

NENHUM teste grava nota nem toca o portal: acesso, preenchimento e gravação do robô são dublês e o
clique real em "Gravar" é trocado por um detonador.
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
import escrituracao_multi_cnpj as m
import escrituracao_multi_robo as robo
from tests.test_escrituracao_multi_cnpj import AcessoFalso
from tests.test_escrituracao_multi_etapa2 import CABECALHOS, ROTULO, _DriverFalso, _linha
from tests.test_escrituracao_multi_cnpj import _planilha

CNPJ_A = "00584628000181"
CNPJ_B = "21345512000160"


class _JanelaFalsa:
    def __init__(self):
        self.js = []

    def evaluate_js(self, codigo):
        # o app envia JSON com escapes \uXXXX; guardamos já decodificado para comparar com o texto real
        self.js.append(re.sub(r"\\u([0-9a-fA-F]{4})", lambda achado: chr(int(achado.group(1), 16)), codigo))

    def resumo(self):
        for c in self.js:
            if c.startswith("window.mostrar_resumo_esc_multi("):
                return json.loads(c[len("window.mostrar_resumo_esc_multi("):-1])
        return None


class _Base(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="app_etapa2_"))
        self.planilha = _planilha(
            self.pasta,
            [
                _linha(CNPJ_A, "11111111111", "SenhaA", 1),
                _linha(CNPJ_A, "11111111111", "SenhaA", 2),
                _linha(CNPJ_B, "22222222222", "SenhaB", 3),
            ],
            cabecalhos=CABECALHOS,
        )
        self.janela = _JanelaFalsa()
        self.api = app.BeAContabAPI(self.janela)
        self.acessos = []
        self.gravadas = []
        self._acesso_original = app.AcessoIssSelenium

        def criar_acesso(**_kw):
            acesso = AcessoFalso()
            acesso.driver = _DriverFalso()
            self.acessos.append(acesso)
            return acesso

        app.AcessoIssSelenium = criar_acesso
        robo.configurar_saida()
        robo.configurar_controle()

        def detonador(*_a, **_k):
            raise AssertionError("TENTATIVA DE CLICAR EM GRAVAR NUM TESTE")

        for alvo, novo in (
            ("_clicar_gravar_documento", detonador),
            ("preparar_tela_digitar_documento", lambda *a, **k: None),
            ("preencher_dados_prestador", lambda *a, **k: None),
            ("preencher_documento_servico", lambda *a, **k: None),
            ("_clicar_com_espera", lambda *a, **k: None),
            ("_gravar_documento_com_confirmacoes", self._gravar_falso),
        ):
            patch = mock.patch.object(robo, alvo, novo)
            patch.start()
            self.addCleanup(patch.stop)
        sono = mock.patch.object(robo.time, "sleep")
        sono.start()
        self.addCleanup(sono.stop)
        self.addCleanup(self._restaurar)

    def _gravar_falso(self, driver, nf, timeout=15):
        self.gravadas.append(nf)
        return "sucesso"

    def _restaurar(self):
        app.AcessoIssSelenium = self._acesso_original
        robo.configurar_saida()
        robo.configurar_controle()
        shutil.rmtree(self.pasta, ignore_errors=True)

    def logs(self) -> str:
        arquivos = sorted(self.pasta.glob("log_execucao_*.txt"))
        return "\n".join(a.read_text(encoding="utf-8") for a in arquivos)

    def iniciar_e_esperar(self, modo, competencia=ROTULO, confirmado=False):
        self.api.iniciar_escrituracao_multi_gui(str(self.planilha), modo, competencia, confirmado)
        controle = self.api._controles_xml["esc_multi"]
        for _ in range(200):
            if not controle.em_execucao:
                return
            threading.Event().wait(0.05)
        self.fail("a execução não terminou a tempo")


class TestRecusas(_Base):
    def _recusou(self, texto):
        self.assertTrue(any(texto in c for c in self.janela.js), self.janela.js)
        self.assertFalse(self.api._controles_xml["esc_multi"].em_execucao)
        self.assertFalse(self.api._automacao_em_execucao)
        self.assertEqual(self.acessos, [])  # nem abriu navegador

    def test_escriturar_sem_confirmacao_e_recusado(self):
        self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "escriturar", ROTULO, False)
        self._recusou("exige a confirmação")
        self.assertEqual(self.gravadas, [])

    def test_confirmacao_precisa_ser_exatamente_verdadeira(self):
        for valor in ("true", 1, "sim", None):
            self.janela.js.clear()
            self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "escriturar", ROTULO, valor)
            self._recusou("exige a confirmação")

    def test_competencia_invalida_e_recusada_em_simular_e_escriturar(self):
        for modo in ("simular", "escriturar"):
            self.janela.js.clear()
            self.api.iniciar_escrituracao_multi_gui(str(self.planilha), modo, "13/2026", True)
            self._recusou("Competência inválida")
        self.janela.js.clear()
        self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "simular", "", True)
        self._recusou("Competência inválida")

    def test_modo_desconhecido_e_planilha_inexistente(self):
        self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "apagar_tudo", ROTULO, True)
        self._recusou("Modo desconhecido")
        self.janela.js.clear()
        self.api.iniciar_escrituracao_multi_gui(str(self.pasta / "nao_existe.xlsx"), "validar")
        self._recusou("Selecione a planilha")

    def test_nao_inicia_se_a_escrituracao_comum_estiver_rodando(self):
        self.api._automacao_em_execucao = True
        self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "simular", ROTULO)
        self.assertTrue(any("Outra automação" in c for c in self.janela.js))
        self.assertFalse(self.api._controles_xml["esc_multi"].em_execucao)
        self.assertTrue(self.api._automacao_em_execucao)  # não mexeu na trava de quem estava rodando

    def test_validar_nao_exige_competencia_nem_confirmacao(self):
        self.iniciar_e_esperar("validar", competencia="")
        self.assertEqual(self.gravadas, [])
        self.assertEqual(self.janela.resumo()["contagem"], {"Acesso validado": 2})


class TestSimular(_Base):
    def test_simulacao_de_ponta_a_ponta_nao_grava_e_registra_tudo_na_pasta_da_planilha(self):
        self.iniciar_e_esperar("simular")
        self.assertEqual(self.gravadas, [])
        resumo = self.janela.resumo()
        self.assertEqual(resumo["contagem"], {"Simulada (nada foi gravado)": 2})
        self.assertTrue(all(e["ok"] for e in resumo["empresas"]))
        # logs unificados, na pasta da planilha e sem subpasta
        self.assertTrue((self.pasta / robo.ARQUIVO_SIMULADAS).exists())
        self.assertFalse((self.pasta / robo.ARQUIVO_SUCESSO).exists())
        self.assertFalse((self.pasta / "log").exists())
        self.assertFalse(m.caminho_estado(self.planilha).exists())
        texto = self.logs()
        self.assertIn("Modo: Simulação (nada é gravado)", texto)
        self.assertIn(f"Competência: {ROTULO}", texto)
        self.assertIn("Pré-análise da planilha", texto)
        self.assertIn("[00.584.628/0001-81] Processando linha", texto)
        self.assertNotIn("SenhaA", texto)

    def test_controle_e_saida_do_robo_sao_zerados_ao_terminar(self):
        self.iniciar_e_esperar("simular")
        self.assertIsNone(robo._CALLBACK_PAUSA)
        self.assertIsNone(robo._CALLBACK_CANCELAMENTO)
        self.assertIsNone(robo._CALLBACK_SAIDA)
        self.assertFalse(robo._GRAVACAO_PERMITIDA)
        self.assertFalse(self.api._automacao_em_execucao)

    def test_marcos_internos_vao_so_para_o_arquivo(self):
        def prestador(driver, doc, depuracao=False):
            robo.registrar_evento_execucao("Botão Consultar acionado", "ISS Fortaleza")

        with mock.patch.object(robo, "preencher_dados_prestador", prestador):
            self.iniciar_e_esperar("simular")
        self.assertIn("[detalhe] [00.584.628/0001-81] ISS Fortaleza: Botão Consultar acionado", self.logs())
        self.assertFalse(any("Botão Consultar acionado" in c for c in self.janela.js))  # a tela fica limpa


class TestEscriturar(_Base):
    def test_escriturar_grava_persiste_o_estado_e_evita_repeticao(self):
        self.iniciar_e_esperar("escriturar", confirmado=True)
        self.assertEqual(sorted(self.gravadas), ["1", "2", "3"])
        self.assertEqual(self.janela.resumo()["contagem"], {"Concluída": 2})
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, ROTULO), {CNPJ_A, CNPJ_B})
        self.assertTrue((self.pasta / robo.ARQUIVO_SUCESSO).exists())
        self.assertIn("[00.584.628/0001-81]", (self.pasta / robo.ARQUIVO_SUCESSO).read_text(encoding="utf-8"))
        self.assertFalse(robo._GRAVACAO_PERMITIDA)

        self.gravadas.clear()
        self.janela.js.clear()
        self.iniciar_e_esperar("escriturar", confirmado=True)  # execução repetida por engano
        self.assertEqual(self.gravadas, [])
        self.assertEqual(self.janela.resumo()["contagem"], {"Já concluída (pulada)": 2})

    def test_escriturar_com_erro_em_nota_fica_com_pendencias_e_reexecucao_refaz_so_ela(self):
        falhar = {"nf": "2"}

        def gravar(driver, nf, timeout=15):
            if nf == falhar["nf"]:
                raise RuntimeError("portal recusou")
            self.gravadas.append(nf)
            return "sucesso"

        with mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar):
            self.iniciar_e_esperar("escriturar", confirmado=True)
        self.assertEqual(self.janela.resumo()["contagem"], {"Concluída com pendências": 1, "Concluída": 1})
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, ROTULO), {CNPJ_B})

        falhar["nf"] = None
        self.gravadas.clear()
        self.janela.js.clear()
        with mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar):
            self.iniciar_e_esperar("escriturar", confirmado=True)
        self.assertEqual(self.gravadas, ["2"])  # só a nota que tinha falhado


class TestPausaNoMeioDaEmpresa(_Base):
    def test_pausa_vale_dentro_da_empresa_e_e_registrada(self):
        """Os helpers reais chamam `_verificar_pausa()` a cada clique; o dublê faz o mesmo, provando a ligação
        entre o botão Pausar (app) e o robô."""
        chamadas = []

        def prestador(driver, doc, depuracao=False):
            robo._verificar_pausa()
            chamadas.append(doc.numero_nf)
            if doc.numero_nf == "1":
                self.api.pausar_processamento_xml("esc_multi")

        with mock.patch.object(robo, "preencher_dados_prestador", prestador):
            self.api.iniciar_escrituracao_multi_gui(str(self.planilha), "simular", ROTULO)
            controle = self.api._controles_xml["esc_multi"]
            # a nota 1 pausa; a nota 2 (mesma empresa) deve travar dentro de `_verificar_pausa`
            for _ in range(60):
                if "Execução PAUSADA" in self.logs():
                    break
                threading.Event().wait(0.05)
            self.assertIn("Execução PAUSADA no ponto em que estava", self.logs())
            self.assertEqual(chamadas, ["1", "2"][: len(chamadas)])
            self.assertLessEqual(len(chamadas), 2)  # parado antes de concluir a empresa A
            self.assertTrue(controle.em_execucao)
            self.assertNotIn("RETOMADA", self.logs())

            self.api.retomar_processamento_xml("esc_multi")
            for _ in range(200):
                if not controle.em_execucao:
                    break
                threading.Event().wait(0.05)
        self.assertFalse(controle.em_execucao)
        self.assertEqual(chamadas, ["1", "2", "3"])
        self.assertIn("RETOMADA pelo operador", self.logs())

    def test_parar_no_meio_da_empresa_cancela_e_zera_o_controle(self):
        def prestador(driver, doc, depuracao=False):
            if doc.numero_nf == "2":
                self.api.cancelar_processamento_xml("esc_multi")

        with mock.patch.object(robo, "preencher_dados_prestador", prestador):
            self.iniciar_e_esperar("simular")
        resumo = self.janela.resumo()
        self.assertTrue(resumo["cancelado"])
        self.assertEqual(self.gravadas, [])
        texto = self.logs()
        self.assertIn("PARADA solicitada pelo operador", texto)
        self.assertIn("Logout confirmado", texto)
        self.assertIsNone(robo._CALLBACK_CANCELAMENTO)
        self.assertFalse(robo._GRAVACAO_PERMITIDA)


if __name__ == "__main__":
    unittest.main()
