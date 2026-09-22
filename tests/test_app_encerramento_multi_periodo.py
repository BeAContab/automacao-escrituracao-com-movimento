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

    def _iniciar_e_esperar(self, competencias, cpf="12345678901", senha="segredo"):
        self.api.iniciar_encerramento_multi_periodo_gui(
            str(self.pasta / "planilha.xlsx"), str(self.pasta / "saida"), competencias,
            "chrome.exe", "https://iss.fortaleza.ce.gov.br/grpfor/login.seam", cpf, senha,
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

    def test_pedido_de_cancelamento_faz_o_callback_do_controlador_levantar(self):
        from core.encerramento_iss_sem_movimento.utils.classes import UserStoppedThreadException

        capturado = {}
        chegou_a_chamar_de_novo = []

        def construtor(**kwargs):
            capturado.update(kwargs)

            def executar():
                cb = kwargs["check_thread_stopped_callback"]
                cb()  # 1ª checagem: ainda não foi pedido cancelamento
                self.api.cancelar_encerramento_multi_periodo()  # operador clica em "Cancelar" no meio da execução
                try:
                    cb()  # 2ª checagem: agora precisa levantar
                except UserStoppedThreadException:
                    chegou_a_chamar_de_novo.append(True)
                    raise
                self.fail("a 2ª checagem deveria ter levantado UserStoppedThreadException")

            fake = mock.Mock()
            fake.executar_processo.side_effect = executar
            return fake

        app.ControladorEncerramentoISSMultiPeriodo = construtor
        self._iniciar_e_esperar(["01/2026"])

        self.assertEqual(chegou_a_chamar_de_novo, [True])
        self.assertIn("Erro crítico no encerramento", self.logs())  # a exceção sobe e vira log (mesmo padrão do original)

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


if __name__ == "__main__":
    unittest.main()
