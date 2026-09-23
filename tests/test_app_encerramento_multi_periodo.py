"""Testes do backend do app para o Encerramento ISS — Múltiplos Meses.

NENHUM teste toca o portal ou abre um navegador: `ControladorEncerramentoISSMultiPeriodo` é
totalmente dublado (patch em `app.ControladorEncerramentoISSMultiPeriodo`).
"""

import json
import re
import shutil
import tempfile
import threading
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import app
from core.encerramento_iss_sem_movimento.controladores.encerramento_iss import CodigoEncerramentoISS, StatusEncerramentoISS
from core.encerramento_iss_sem_movimento.controladores.encerramento_iss_multi_periodo import (
    ResultadoEncerramentoISSMultiPeriodo,
    ResultadoMesEmpresa,
)
from core.encerramento_iss_sem_movimento.iss.empresa import EmpresaSemMovimentoISSFortaleza


class _JanelaFalsa:
    def __init__(self):
        self.js = []

    def evaluate_js(self, codigo):
        # o app envia JSON com escapes \uXXXX; decodificado aqui para comparar com texto real
        self.js.append(re.sub(r"\\u([0-9a-fA-F]{4})", lambda achado: chr(int(achado.group(1), 16)), codigo))

    def resumo(self):
        for c in self.js:
            if c.startswith("window.mostrar_resumo_encerramento_multi("):
                return json.loads(c[len("window.mostrar_resumo_encerramento_multi("):-1])
        return None


def _empresa() -> EmpresaSemMovimentoISSFortaleza:
    return EmpresaSemMovimentoISSFortaleza(
        codigo=1, nome="EMPRESA TESTE", cnpj="12345678000199", responsavel="SM - FULANO",
        municipio="FORTALEZA", linha_planilha_fiscal=2,
    )


class _Base(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="app_encerr_multi_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        (self.pasta / "planilha.xlsx").write_text("fingido")  # só precisa existir p/ Path(...); não é lido de verdade
        self.janela = _JanelaFalsa()
        self.api = app.BeAContabAPI(self.janela)
        self._original_controlador = app.ControladorEncerramentoISSMultiPeriodo

    def tearDown(self):
        app.ControladorEncerramentoISSMultiPeriodo = self._original_controlador

    def _resultado_padrao(self):
        emp = _empresa()
        resultados = [
            ResultadoMesEmpresa(emp, date(2026, 1, 1), encerrada=True, dt_encerramento=date(2026, 1, 5)),
            ResultadoMesEmpresa(emp, date(2026, 2, 1), encerrada=False, problemas=["SERVIÇOS PENDENTES"]),
        ]
        return ResultadoEncerramentoISSMultiPeriodo(
            status=StatusEncerramentoISS(CodigoEncerramentoISS.SUCCESS, "ok"),
            resultados=resultados,
            report_paths=[self.pasta / "saida" / "2026-01" / "relatorio.xlsx", self.pasta / "saida" / "2026-02" / "relatorio.xlsx"],
        )

    def _instalar_controlador_falso(self, capturado: dict, resultado=None, erro: Exception | None = None):
        def construtor(**kwargs):
            capturado.update(kwargs)
            controlador = mock.Mock()
            if erro is not None:
                controlador.executar_processo.side_effect = erro
            else:
                controlador.executar_processo.return_value = resultado or self._resultado_padrao()
            return controlador

        app.ControladorEncerramentoISSMultiPeriodo = construtor

    def _iniciar_e_esperar(self, competencias, cpf="12345678901", senha="segredo", **extra):
        self.api.iniciar_encerramento_multi_periodo_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), competencias,
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", cpf, senha, **extra,
        )
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                return
            threading.Event().wait(0.05)
        self.fail("a execução não terminou a tempo")

    def logs(self) -> str:
        arquivos = sorted(self.pasta.glob("saida/log_execucao_*.txt"))
        self.assertTrue(arquivos, "nenhum arquivo de log foi criado")
        return "\n".join(a.read_text(encoding="utf-8") for a in arquivos)


class TestFluxoNormal(_Base):
    def test_competencias_sao_convertidas_para_date_na_ordem_recebida(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026", "02/2026"])
        self.assertEqual(capturado["competencias"], [date(2026, 1, 1), date(2026, 2, 1)])

    def test_resumo_mandado_para_a_gui_tem_os_totais_e_os_relatorios(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026", "02/2026"])
        resumo = self.janela.resumo()
        self.assertEqual(resumo["processadas"], 2)
        self.assertEqual(resumo["encerradas"], 1)
        self.assertEqual(resumo["problemas"], 1)
        self.assertEqual(len(resumo["report_paths"]), 2)
        self.assertEqual(resumo["pasta_saida"], str(self.pasta / "saida"))

    def test_log_arquivo_registra_a_conclusao(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertIn("Múltiplos Meses concluído", self.logs())

    def test_credenciais_nunca_aparecem_no_log(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"], cpf="98765432100", senha="SenhaSuperSecreta")
        texto = self.logs()
        self.assertNotIn("SenhaSuperSecreta", texto)
        self.assertNotIn("98765432100", texto)

    def test_flag_de_execucao_liberada_ao_terminar(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertFalse(self.api._encerramento_multi_periodo_em_execucao)


class TestErrosECancelamento(_Base):
    def test_erro_inesperado_e_reportado_no_log_e_libera_a_trava(self):
        capturado = {}
        self._instalar_controlador_falso(capturado, erro=RuntimeError("o portal caiu"))
        self._iniciar_e_esperar(["01/2026"])
        self.assertIn("Erro crítico no encerramento", self.logs())
        self.assertFalse(self.api._encerramento_multi_periodo_em_execucao)

    def test_parar_nao_derruba_o_callback_imediato_mas_liga_o_callback_de_parada(self):
        capturado = {}
        observado = {}

        def construtor(**kwargs):
            capturado.update(kwargs)

            def executar():
                observado["antes"] = kwargs["callback_parar"]()
                kwargs["check_thread_stopped_callback"]()  # parada imediata desligada: nunca levanta
                self.api.cancelar_encerramento_multi_periodo()  # operador clica em "Parar"
                kwargs["check_thread_stopped_callback"]()  # continua sem levantar no meio de um encerramento
                observado["depois"] = kwargs["callback_parar"]()
                return self._resultado_padrao()

            fake = mock.Mock()
            fake.executar_processo.side_effect = executar
            return fake

        app.ControladorEncerramentoISSMultiPeriodo = construtor
        self._iniciar_e_esperar(["01/2026"])

        self.assertEqual(observado, {"antes": False, "depois": True})
        self.assertNotIn("Erro crítico", self.logs())

    def test_nao_inicia_uma_segunda_execucao_enquanto_a_primeira_esta_rodando(self):
        liberar = threading.Event()
        capturado = {}

        def construtor(**kwargs):
            capturado.update(kwargs)
            fake = mock.Mock()
            fake.executar_processo.side_effect = lambda: (liberar.wait(5), self._resultado_padrao())[1]
            return fake

        app.ControladorEncerramentoISSMultiPeriodo = construtor
        self.api.iniciar_encerramento_multi_periodo_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), ["01/2026"],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        for _ in range(100):
            if self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.02)
        self.assertTrue(self.api._encerramento_multi_periodo_em_execucao)

        self.janela.js.clear()
        self.api.iniciar_encerramento_multi_periodo_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), ["02/2026"],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        self.assertTrue(any("já está em andamento" in c for c in self.janela.js))

        liberar.set()
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.02)
        self.assertFalse(self.api._encerramento_multi_periodo_em_execucao)


class TestPausaEParadaNoApp(_Base):
    def test_controlador_recebe_callbacks_de_pausa_e_parada_e_o_modo(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertEqual(capturado["modo"], "encerrar")
        self.assertTrue(callable(capturado["callback_pausa"]))
        self.assertTrue(callable(capturado["callback_parar"]))
        self.assertIsNone(capturado["alvos"])

    def test_pausa_e_retomada_sao_avisadas_no_log(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)

        def construtor(**kwargs):
            capturado.update(kwargs)

            def executar():
                # o robô chega ao limite seguro com a execução em andamento (log do controle ativo)
                self.api.pausar_processamento_xml("encerramento_multi")
                threading.Timer(0.1, lambda: self.api.retomar_processamento_xml("encerramento_multi")).start()
                kwargs["callback_pausa"]()  # bloqueia até o Timer retomar
                return self._resultado_padrao()

            fake = mock.Mock()
            fake.executar_processo.side_effect = executar
            return fake

        app.ControladorEncerramentoISSMultiPeriodo = construtor
        self._iniciar_e_esperar(["01/2026"])
        texto = self.logs()
        self.assertIn("PAUSA solicitada pelo operador", texto)
        self.assertIn("RETOMADA pelo operador", texto)
        self.assertIn("nunca no meio de um encerramento", texto)

    def test_iniciar_devolve_true_ao_iniciar_e_false_se_ja_ha_execucao_para_a_tela_destravar(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        args = (
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), ["01/2026"],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        self.assertTrue(self.api.iniciar_encerramento_multi_periodo_gui(*args))
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.05)
        self.api._encerramento_multi_periodo_em_execucao = True  # simula outra execução em andamento
        self.assertFalse(self.api.iniciar_encerramento_multi_periodo_gui(*args))
        self.api._encerramento_multi_periodo_em_execucao = False

    def test_fim_da_execucao_avisa_a_tela_para_destravar_os_botoes(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertTrue(any("encerramento_multi_finalizado" in c for c in self.janela.js))

    def test_parar_registra_aviso_e_libera_uma_pausa_em_andamento(self):
        controle = self.api._controles_xml["encerramento_multi"]
        controle.reiniciar()
        controle.log = lambda *_a, **_k: None
        self.api.pausar_processamento_xml("encerramento_multi")
        self.assertFalse(controle.pausa.is_set())
        self.api.cancelar_processamento_xml("encerramento_multi")
        self.assertTrue(controle.cancelado())
        self.assertTrue(controle.pausa.is_set())  # a parada acorda o robô pausado

    def test_iniciar_zera_uma_parada_de_execucao_anterior(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self.api.cancelar_encerramento_multi_periodo()  # sobra de uma execução anterior
        self._iniciar_e_esperar(["01/2026"])
        self.assertFalse(capturado["callback_parar"]())

    def test_parado_pelo_operador_loga_parado_e_nao_concluido(self):
        capturado = {}
        resultado = self._resultado_padrao()
        resultado.status = StatusEncerramentoISS(CodigoEncerramentoISS.USER_ENDED_PROCESS, "parado")
        self._instalar_controlador_falso(capturado, resultado=resultado)
        self._iniciar_e_esperar(["01/2026"])
        texto = self.logs()
        self.assertIn("PARADO pelo operador", texto)
        self.assertNotIn("Múltiplos Meses concluído", texto)


class TestModoVerificarNoApp(_Base):
    def _resultado_verificacao(self):
        emp = _empresa()
        emp.cnpj_m = "12.345.678/0001-99"
        resultados = [
            ResultadoMesEmpresa(emp, date(2026, 1, 1), encerrada=False, acao="apta", situacao="Aberta - Normal 01/01/2026"),
            ResultadoMesEmpresa(emp, date(2026, 2, 1), encerrada=True, acao="ja_encerrada_sem_certificado"),
            ResultadoMesEmpresa(emp, date(2026, 3, 1), encerrada=True, acao="ja_encerrada_certificado_existente"),
            ResultadoMesEmpresa(emp, date(2026, 4, 1), encerrada=False, problemas=["SERVIÇOS PRESTADOS"]),
        ]
        return ResultadoEncerramentoISSMultiPeriodo(
            status=StatusEncerramentoISS(CodigoEncerramentoISS.SUCCESS, "ok"), resultados=resultados, report_paths=[],
        )

    def test_modo_verificar_e_repassado_ao_controlador(self):
        capturado = {}
        self._instalar_controlador_falso(capturado, resultado=self._resultado_verificacao())
        self._iniciar_e_esperar(["01/2026"], modo="verificar")
        self.assertEqual(capturado["modo"], "verificar")

    def test_resumo_da_verificacao_traz_a_lista_com_o_que_e_executavel(self):
        capturado = {}
        self._instalar_controlador_falso(capturado, resultado=self._resultado_verificacao())
        self._iniciar_e_esperar(["01/2026"], modo="verificar")
        resumo = self.janela.resumo()
        self.assertEqual(resumo["modo"], "verificar")
        self.assertEqual((resumo["aptas"], resumo["ja_encerradas"], resumo["problemas"], resumo["encerradas_agora"]), (1, 2, 1, 0))
        itens = {i["competencia"]: i for i in resumo["verificacao"]}
        self.assertEqual(sorted(itens), ["01/2026", "02/2026", "03/2026", "04/2026"])
        self.assertTrue(itens["01/2026"]["executavel"])  # apta
        self.assertTrue(itens["02/2026"]["executavel"])  # já encerrada, falta o certificado
        self.assertFalse(itens["03/2026"]["executavel"])  # já encerrada, certificado existe
        self.assertFalse(itens["04/2026"]["executavel"])  # problema
        self.assertEqual(itens["04/2026"]["motivo"], "SERVIÇOS PRESTADOS")
        self.assertEqual(itens["01/2026"]["situacao"], "Aberta - Normal 01/01/2026")
        self.assertEqual(itens["01/2026"]["cnpj"], "12345678000199")

    def test_log_da_verificacao_diz_que_nada_foi_encerrado(self):
        capturado = {}
        self._instalar_controlador_falso(capturado, resultado=self._resultado_verificacao())
        self._iniciar_e_esperar(["01/2026"], modo="verificar")
        texto = self.logs()
        self.assertIn("nada foi encerrado nem baixado", texto)
        self.assertNotIn("Múltiplos Meses concluído", texto)

    def test_modo_encerrar_nao_manda_lista_de_verificacao(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertNotIn("verificacao", self.janela.resumo())


class TestExecutarVerificadasNoApp(_Base):
    def _executar_e_esperar(self, itens):
        self.api.executar_verificadas_encerramento_multi_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), itens,
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                return
            threading.Event().wait(0.05)
        self.fail("a execução não terminou a tempo")

    def test_executa_em_modo_encerrar_so_com_os_alvos_aprovados(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._executar_e_esperar([
            {"cnpj": "12.345.678/0001-99", "competencia": "02/2026"},
            {"cnpj": "12345678000199", "competencia": "01/2026"},
            {"cnpj": "98765432000111", "competencia": "02/2026"},
        ])
        self.assertEqual(capturado["modo"], "encerrar")
        self.assertEqual(capturado["competencias"], [date(2026, 1, 1), date(2026, 2, 1)])
        self.assertEqual(
            capturado["alvos"],
            {"12345678000199": {date(2026, 1, 1), date(2026, 2, 1)}, "98765432000111": {date(2026, 2, 1)}},
        )

    def test_itens_invalidos_sao_ignorados(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._executar_e_esperar([
            {"cnpj": "123", "competencia": "01/2026"},  # CNPJ curto
            {"cnpj": "12345678000199", "competencia": "13-2026"},  # competência inválida
            {"cnpj": "12345678000199", "competencia": "03/2026"},  # válido
            None,
        ])
        self.assertEqual(capturado["alvos"], {"12345678000199": {date(2026, 3, 1)}})

    def test_lista_sem_itens_validos_nao_inicia_nada(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self.api.executar_verificadas_encerramento_multi_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), [{"cnpj": "1", "competencia": "x"}],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        self.assertEqual(capturado, {})  # o controlador nem foi criado
        self.assertFalse(self.api._encerramento_multi_periodo_em_execucao)
        self.assertTrue(any("nenhum item válido" in c for c in self.janela.js))

    def test_nao_executa_enquanto_outra_execucao_esta_em_andamento(self):
        self.api._encerramento_multi_periodo_em_execucao = True
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self.api.executar_verificadas_encerramento_multi_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), [{"cnpj": "12345678000199", "competencia": "01/2026"}],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        self.assertEqual(capturado, {})
        self.assertTrue(any("já está em andamento" in c for c in self.janela.js))
        self.api._encerramento_multi_periodo_em_execucao = False


CNPJ_VALIDO_A = "11222333000181"
CNPJ_VALIDO_B = "11444777000161"


class TestOrigemManualNoApp(_Base):
    def _iniciar(self, cnpjs, modo="encerrar", planilha=""):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        retorno = self.api.iniciar_encerramento_multi_periodo_gui(
            planilha, str(self.pasta / "saida"), ["01/2026"],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
            modo, cnpjs,
        )
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.05)
        return capturado, retorno

    def test_cnpjs_digitados_substituem_a_planilha_no_controlador(self):
        capturado, retorno = self._iniciar(["11.222.333/0001-81", CNPJ_VALIDO_B, CNPJ_VALIDO_A])
        self.assertTrue(retorno)
        self.assertIsNone(capturado["caminho_planilha"])  # nenhuma planilha é lida
        self.assertEqual(capturado["cnpjs"], [CNPJ_VALIDO_A, CNPJ_VALIDO_B])  # normalizados e sem duplicatas

    def test_sem_cnpjs_continua_usando_a_planilha(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertEqual(capturado["caminho_planilha"], self.pasta / "planilha.xlsx")
        self.assertIsNone(capturado["cnpjs"])

    def test_cnpj_invalido_nao_inicia_e_mostra_o_erro_na_tela(self):
        capturado, retorno = self._iniciar([CNPJ_VALIDO_A, "12345678000199"])
        self.assertFalse(retorno)
        self.assertEqual(capturado, {})  # o controlador nem foi criado
        self.assertFalse(self.api._encerramento_multi_periodo_em_execucao)
        self.assertTrue(any("CNPJ inválido: 12345678000199" in c for c in self.janela.js))

    def test_resumo_informa_a_origem(self):
        self._iniciar([CNPJ_VALIDO_A])
        self.assertEqual(self.janela.resumo()["origem"], "manual")

    def test_resumo_da_planilha_informa_origem_planilha(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self._iniciar_e_esperar(["01/2026"])
        self.assertEqual(self.janela.resumo()["origem"], "planilha")

    def test_executar_verificadas_no_manual_usa_so_os_cnpjs_aprovados(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self.api.executar_verificadas_encerramento_multi_gui(
            "", str(self.pasta / "saida"),
            [{"cnpj": CNPJ_VALIDO_B, "competencia": "02/2026"}, {"cnpj": CNPJ_VALIDO_A, "competencia": "01/2026"}],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo", "manual",
        )
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.05)
        self.assertIsNone(capturado["caminho_planilha"])
        self.assertEqual(sorted(capturado["cnpjs"]), sorted([CNPJ_VALIDO_A, CNPJ_VALIDO_B]))
        self.assertEqual(capturado["modo"], "encerrar")
        self.assertEqual(set(capturado["alvos"]), {CNPJ_VALIDO_A, CNPJ_VALIDO_B})

    def test_executar_verificadas_da_planilha_continua_sem_cnpjs(self):
        capturado = {}
        self._instalar_controlador_falso(capturado)
        self.api.executar_verificadas_encerramento_multi_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"),
            [{"cnpj": CNPJ_VALIDO_A, "competencia": "01/2026"}],
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", "12345678901", "segredo",
        )
        for _ in range(200):
            if not self.api._encerramento_multi_periodo_em_execucao:
                break
            threading.Event().wait(0.05)
        self.assertIsNone(capturado["cnpjs"])
        self.assertEqual(capturado["caminho_planilha"], self.pasta / "planilha.xlsx")


if __name__ == "__main__":
    unittest.main()
