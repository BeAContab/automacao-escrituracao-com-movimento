"""Testes do Encerramento ISS — Múltiplos Meses (encerramento_iss_multi_periodo.py).

REGRA: nenhum teste abre um navegador nem toca o portal ISS. As interações com o portal
(login, busca de inscrição, navegação, processamento de cada competência) são sempre trocadas
por dublês (`unittest.mock`); só a leitura/escrita da planilha de entrada e a geração do
relatório usam arquivos reais (openpyxl), sem qualquer chamada de rede ou Selenium.
"""

import shutil
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook
from pydantic import SecretStr

import core.encerramento_iss_sem_movimento.controladores.encerramento_iss_multi_periodo as m
from core.encerramento_iss_sem_movimento.controladores.encerramento_iss import CodigoEncerramentoISS
from core.encerramento_iss_sem_movimento.iss.credenciais import CredenciaisISSFortaleza
from core.encerramento_iss_sem_movimento.iss.empresa import EmpresaSemMovimentoISSFortaleza

CAMINHO_TEMPLATE = Path("core/templates/template_relatorio_execucao.xlsx")


def _empresa(codigo=1, cnpj="12345678000199", nome="EMPRESA TESTE", linha=2) -> EmpresaSemMovimentoISSFortaleza:
    return EmpresaSemMovimentoISSFortaleza(
        codigo=codigo, nome=nome, cnpj=cnpj, responsavel="SM - FULANO", municipio="FORTALEZA", linha_planilha_fiscal=linha
    )


def _planilha_fiscal(pasta: Path, empresas: list[tuple[int, str, str]], nome: str = "planilha_fiscal.xlsx") -> Path:
    """Cria uma planilha fiscal mínima: cada item é (codigo, nome, cnpj); uma linha por empresa,
    a partir da linha 2 (linha 1 é cabeçalho, igual à planilha real)."""
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["D1"], ws["H1"], ws["W1"], ws["Y1"] = "CODIGO", "NOME", "CNPJ", "RESPONSAVEL", "MUNICIPIO", "RESULTADO"
    for i, (codigo, nome_empresa, cnpj) in enumerate(empresas, start=2):
        ws[f"A{i}"] = codigo
        ws[f"B{i}"] = nome_empresa
        ws[f"D{i}"] = cnpj
        ws[f"H{i}"] = "SM - FULANO"
        ws[f"W{i}"] = "FORTALEZA"
    caminho = pasta / nome
    wb.save(caminho)
    return caminho


def _controlador(caminho_planilha: Path, pasta_saida: Path, competencias: list[date]) -> m.ControladorEncerramentoISSMultiPeriodo:
    return m.ControladorEncerramentoISSMultiPeriodo(
        caminho_planilha=caminho_planilha,
        caminho_saida=pasta_saida,
        caminho_webdriver=Path("chrome.exe"),
        url_iss_fortaleza="https://iss.fortaleza.ce.gov.br/grpfor/login.seam",
        credenciais=CredenciaisISSFortaleza(cpf=SecretStr("12345678901"), senha=SecretStr("segredo")),
        competencias=competencias,
        caminho_template_relatorio=CAMINHO_TEMPLATE,
        check_thread_stopped_callback=lambda: None,
    )


class TestConstrutor(unittest.TestCase):
    def test_sem_competencias_levanta_erro(self):
        with self.assertRaises(ValueError):
            _controlador(Path("planilha.xlsx"), Path("saida"), [])

    def test_competencias_ficam_ordenadas_e_sem_duplicatas(self):
        c = _controlador(
            Path("planilha.xlsx"),
            Path("saida"),
            [date(2026, 3, 1), date(2025, 12, 1), date(2026, 1, 1), date(2026, 3, 1)],
        )
        self.assertEqual(c.competencias, [date(2025, 12, 1), date(2026, 1, 1), date(2026, 3, 1)])


class TestResultados(unittest.TestCase):
    def test_contagens_da_classe_de_resultado(self):
        emp = _empresa()
        resultados = [
            m.ResultadoMesEmpresa(emp, date(2026, 1, 1), encerrada=True),
            m.ResultadoMesEmpresa(emp, date(2026, 2, 1), encerrada=False, problemas=["SERVIÇOS PENDENTES"]),
            m.ResultadoMesEmpresa(emp, date(2026, 3, 1), encerrada=True),
        ]
        resultado = m.ResultadoEncerramentoISSMultiPeriodo(status=None, resultados=resultados, report_paths=[])
        self.assertEqual(resultado.processadas, 3)
        self.assertEqual(resultado.encerradas, 2)
        self.assertEqual(resultado.problemas, 1)

    def test_contagens_com_lista_vazia(self):
        resultado = m.ResultadoEncerramentoISSMultiPeriodo(status=None, resultados=[], report_paths=[])
        self.assertEqual((resultado.processadas, resultado.encerradas, resultado.problemas), (0, 0, 0))


class TestAcrescentarNaPlanilhaFiscal(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.caminho = _planilha_fiscal(self.pasta, [(1, "EMPRESA A", "12345678000199")])
        self.controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])

    def _ler_celula_y2(self) -> str:
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        valor = reader.get_cell("Y2").value
        reader.close_workbook()
        return valor

    def test_primeiro_acrescimo_nao_leva_ponto_e_virgula(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), "05/01/2026", row_number=2)
        reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "01/2026: 05/01/2026")

    def test_varias_competencias_ficam_concatenadas_na_mesma_celula_sem_apagar_a_anterior(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), "05/01/2026", row_number=2)
        reader.close_workbook()

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 2, 1), "SERVIÇOS PENDENTES", row_number=2)
        reader.close_workbook()

        self.assertEqual(self._ler_celula_y2(), "01/2026: 05/01/2026; 02/2026: SERVIÇOS PENDENTES")

    def test_sem_inscricao_nao_leva_rotulo_de_competencia(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, competencia=None, valor="S/ INSCRIÇÃO", row_number=2)
        reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "S/ INSCRIÇÃO")

    def test_nao_apaga_conteudo_que_ja_existia_manualmente_na_celula(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        reader.set_cell_value("Anotação manual do contador", coords="Y2")
        reader.save_spreadsheet(close_workbook=False)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), "05/01/2026", row_number=2)
        reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "Anotação manual do contador; 01/2026: 05/01/2026")

    def test_repetir_a_mesma_informacao_nao_empilha_na_celula(self):
        from python_spreadsheet_reader.readers import XLSXReader

        for _ in range(3):  # três execuções seguidas com o mesmo resultado
            reader = XLSXReader(self.caminho)
            self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), m.MSG_ERRO_PORTAL, row_number=2)
            reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "01/2026: ERRO NO PORTAL")

    def test_resultado_diferente_para_a_mesma_competencia_e_acrescentado(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), m.MSG_ERRO_PORTAL, row_number=2)
        reader.close_workbook()
        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), "05/02/2026", row_number=2)
        reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "01/2026: ERRO NO PORTAL; 01/2026: 05/02/2026")

    def test_competencia_diferente_com_o_mesmo_texto_nao_e_confundida(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 1, 1), m.MSG_ERRO_PORTAL, row_number=2)
        reader.close_workbook()
        reader = XLSXReader(self.caminho)
        self.controlador._acrescentar_na_planilha_fiscal(reader, date(2026, 2, 1), m.MSG_ERRO_PORTAL, row_number=2)
        reader.close_workbook()
        self.assertEqual(self._ler_celula_y2(), "01/2026: ERRO NO PORTAL; 02/2026: ERRO NO PORTAL")


@unittest.skipUnless(CAMINHO_TEMPLATE.exists(), "template do relatório não encontrado")
class TestRelatorioDeExecucao(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_rel_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.controlador = _controlador(self.pasta / "fiscal.xlsx", self.pasta / "saida", [date(2026, 1, 1), date(2026, 2, 1)])

    def test_sem_nenhum_resultado_levanta_excecao_e_nao_gera_arquivo(self):
        with self.assertRaises(m.SemEmpresasProcessadasException):
            self.controlador._gerar_relatorios_de_execucao()

    def test_gera_um_relatorio_por_competencia_na_pasta_ano_mes(self):
        emp_a = _empresa(codigo=1, cnpj="12345678000199", nome="EMPRESA A")
        emp_b = _empresa(codigo=2, cnpj="98765432000111", nome="EMPRESA B", linha=3)
        self.controlador.resultados = [
            m.ResultadoMesEmpresa(emp_a, date(2026, 1, 1), encerrada=True, dt_encerramento=date(2026, 1, 5)),
            m.ResultadoMesEmpresa(emp_b, date(2026, 1, 1), encerrada=False, problemas=["SERVIÇOS PENDENTES"]),
            m.ResultadoMesEmpresa(emp_a, date(2026, 2, 1), encerrada=True, dt_encerramento=date(2026, 2, 6)),
        ]
        caminhos = self.controlador._gerar_relatorios_de_execucao()

        self.assertEqual(len(caminhos), 2)  # uma competência sem resultado (só teria se pedisse 3 meses) não gera nada
        nomes_pastas = sorted(p.parent.name for p in caminhos)
        self.assertEqual(nomes_pastas, ["2026-01", "2026-02"])

        wb = load_workbook(caminhos[0])
        ws = wb.active
        self.assertEqual(ws.cell(row=7, column=2).value, "Competência: 01/2026")
        self.assertEqual(ws.cell(row=5, column=3).value, 2)  # processadas em jan/2026
        self.assertEqual(ws.cell(row=5, column=5).value, 1)  # encerradas em jan/2026
        self.assertEqual(ws.cell(row=5, column=7).value, 1)  # problemas em jan/2026
        nomes_no_relatorio = {ws.cell(row=10, column=2).value, ws.cell(row=11, column=2).value}
        self.assertEqual(nomes_no_relatorio, {"EMPRESA A", "EMPRESA B"})

    def test_relatorio_de_fevereiro_so_tem_a_empresa_daquele_mes(self):
        emp_a = _empresa(codigo=1, cnpj="12345678000199", nome="EMPRESA A")
        self.controlador.resultados = [
            m.ResultadoMesEmpresa(emp_a, date(2026, 1, 1), encerrada=True, dt_encerramento=date(2026, 1, 5)),
            m.ResultadoMesEmpresa(emp_a, date(2026, 2, 1), encerrada=True, dt_encerramento=date(2026, 2, 6)),
        ]
        caminhos = self.controlador._gerar_relatorios_de_execucao()
        caminho_fev = next(p for p in caminhos if p.parent.name == "2026-02")
        wb = load_workbook(caminho_fev)
        ws = wb.active
        self.assertEqual(ws.cell(row=5, column=3).value, 1)


class TestOrquestracaoPorEmpresaEPorMes(unittest.TestCase):
    """Exercita `_processar_linha` com os métodos que tocam o portal totalmente dublados —
    prova a ordem (por empresa, depois por mês) e o isolamento de erros, sem qualquer Selenium."""

    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_orq_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.caminho_planilha = _planilha_fiscal(
            self.pasta, [(1, "EMPRESA A", "12345678000199"), (2, "EMPRESA B", "98765432000111")]
        )
        self.controlador = _controlador(
            self.caminho_planilha, self.pasta / "saida", [date(2026, 1, 1), date(2026, 2, 1)]
        )
        # Planilha com uma única empresa, para os testes que olham o resultado de UMA empresa
        # isoladamente (o comportamento entre empresas diferentes já é coberto pelos testes acima/abaixo).
        self.caminho_planilha_1_empresa = _planilha_fiscal(
            self.pasta, [(1, "EMPRESA A", "12345678000199")], nome="planilha_fiscal_1_empresa.xlsx"
        )
        self.controlador_1_empresa = _controlador(
            self.caminho_planilha_1_empresa, self.pasta / "saida", [date(2026, 1, 1), date(2026, 2, 1)]
        )

    def _reader(self):
        from python_spreadsheet_reader.readers import XLSXReader

        return XLSXReader(self.caminho_planilha)

    def _reader_1_empresa(self):
        from python_spreadsheet_reader.readers import XLSXReader

        return XLSXReader(self.caminho_planilha_1_empresa)

    def test_processa_todos_os_meses_de_uma_empresa_antes_de_abrir_o_modal_de_troca(self):
        chamadas = []
        with mock.patch.object(self.controlador, "_procurar_inscricao_empresa", return_value=True), \
             mock.patch.object(self.controlador, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(self.controlador, "_navegar_tela_escrituracao", lambda driver: chamadas.append("navegar")), \
             mock.patch.object(
                 self.controlador, "_processar_competencia",
                 lambda driver, reader, empresa, competencia: chamadas.append(("competencia", competencia)) or m.ResultadoMesEmpresa(empresa, competencia, encerrada=True),
             ), \
             mock.patch.object(self.controlador, "_abrir_modal_de_alteracao_de_inscricao", lambda driver: chamadas.append("modal")):
            reader = self._reader()
            for row_number, row in reader.lazy_load_sheet():
                self.controlador._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
            reader.close_workbook()

        # 1 navegação e 1 abertura de modal por empresa, com as 2 competências entre elas, na ordem certa
        self.assertEqual(
            chamadas,
            [
                "navegar", ("competencia", date(2026, 1, 1)), ("competencia", date(2026, 2, 1)), "modal",
                "navegar", ("competencia", date(2026, 1, 1)), ("competencia", date(2026, 2, 1)), "modal",
            ],
        )
        self.assertEqual(len(self.controlador.resultados), 4)  # 2 empresas × 2 competências

    def test_empresa_sem_inscricao_marca_todas_as_competencias_e_nao_abre_o_modal(self):
        with mock.patch.object(self.controlador_1_empresa, "_procurar_inscricao_empresa", return_value=False), \
             mock.patch.object(self.controlador_1_empresa, "_abrir_modal_de_alteracao_de_inscricao") as modal, \
             mock.patch.object(self.controlador_1_empresa, "_navegar_tela_escrituracao") as navegar:
            reader = self._reader_1_empresa()
            # A linha 1 (cabeçalho) também é processada, mas é ignorada silenciosamente: os valores de
            # texto ("CODIGO", "NOME"...) falham na validação do Pydantic antes de chegar a qualquer
            # chamada ao portal — só a empresa da linha 2 realmente gera resultado.
            for row_number, row in reader.lazy_load_sheet():
                self.controlador_1_empresa._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
            reader.close_workbook()

        resultados = self.controlador_1_empresa.resultados
        self.assertEqual(len(resultados), 2)  # as 2 competências pedidas, ambas com problema
        self.assertTrue(all(r.problemas == ["S/ INSCRIÇÃO"] and not r.encerrada for r in resultados))
        modal.assert_not_called()
        navegar.assert_not_called()

        # a planilha recebeu o aviso, sem rótulo de mês (não é específico de uma competência)
        reader2 = self._reader_1_empresa()
        self.assertEqual(reader2.get_cell("Y2").value, "S/ INSCRIÇÃO")
        reader2.close_workbook()

    def test_erro_num_mes_nao_impede_os_demais_meses_da_mesma_empresa(self):
        def processar_competencia(driver, reader, empresa, competencia):
            if competencia == date(2026, 1, 1):
                raise m.NoSuchElementException("não achei a linha da competência")
            return m.ResultadoMesEmpresa(empresa, competencia, encerrada=True)

        with mock.patch.object(self.controlador_1_empresa, "_procurar_inscricao_empresa", return_value=True), \
             mock.patch.object(self.controlador_1_empresa, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(self.controlador_1_empresa, "_navegar_tela_escrituracao"), \
             mock.patch.object(self.controlador_1_empresa, "_processar_competencia", side_effect=processar_competencia), \
             mock.patch.object(self.controlador_1_empresa, "_abrir_modal_de_alteracao_de_inscricao") as modal:
            reader = self._reader_1_empresa()
            for row_number, row in reader.lazy_load_sheet():
                self.controlador_1_empresa._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
            reader.close_workbook()

        resultados = self.controlador_1_empresa.resultados
        self.assertEqual(len(resultados), 2)
        self.assertFalse(resultados[0].encerrada)
        self.assertEqual(resultados[0].problemas, [m.MSG_ERRO_PORTAL])  # texto curto, sem nome de exceção do Selenium
        self.assertTrue(resultados[1].encerrada)  # fevereiro seguiu normalmente
        modal.assert_called_once()  # ainda troca de empresa ao final, mesmo com uma falha no meio

        reader2 = self._reader_1_empresa()
        self.assertEqual(reader2.get_cell("Y2").value, "01/2026: ERRO NO PORTAL")
        reader2.close_workbook()

    def test_erro_inesperado_num_mes_refaz_o_caminho_pelo_menu_e_segue_para_o_proximo_mes(self):
        chamadas = []

        def processar_competencia(driver, reader, empresa, competencia):
            if competencia == date(2026, 1, 1):
                raise RuntimeError("falha no meio do fluxo")  # não é NoSuchElement/Timeout
            return m.ResultadoMesEmpresa(empresa, competencia, encerrada=True)

        with mock.patch.object(self.controlador_1_empresa, "_procurar_inscricao_empresa", return_value=True), \
             mock.patch.object(self.controlador_1_empresa, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(self.controlador_1_empresa, "_navegar_tela_escrituracao", lambda driver: chamadas.append("navegar")), \
             mock.patch.object(self.controlador_1_empresa, "_processar_competencia", side_effect=processar_competencia), \
             mock.patch.object(self.controlador_1_empresa, "_abrir_modal_de_alteracao_de_inscricao", lambda driver: chamadas.append("modal")):
            reader = self._reader_1_empresa()
            for row_number, row in reader.lazy_load_sheet():
                self.controlador_1_empresa._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
            reader.close_workbook()

        resultados = self.controlador_1_empresa.resultados
        self.assertEqual([r.encerrada for r in resultados], [False, True])
        self.assertEqual(resultados[0].problemas, [m.MSG_ERRO_PORTAL])
        # 1 navegação inicial + 1 de recuperação depois do erro em janeiro; o modal só no fim
        self.assertEqual(chamadas, ["navegar", "navegar", "modal"])

    def test_navegador_fechado_no_meio_de_um_mes_nao_e_engolido(self):
        def processar_competencia(driver, reader, empresa, competencia):
            raise m.InvalidSessionIdException("invalid session id: session deleted as the browser has closed the connection")

        with mock.patch.object(self.controlador_1_empresa, "_procurar_inscricao_empresa", return_value=True), \
             mock.patch.object(self.controlador_1_empresa, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(self.controlador_1_empresa, "_navegar_tela_escrituracao"), \
             mock.patch.object(self.controlador_1_empresa, "_processar_competencia", side_effect=processar_competencia), \
             mock.patch.object(self.controlador_1_empresa, "_abrir_modal_de_alteracao_de_inscricao"):
            reader = self._reader_1_empresa()
            with self.assertRaises(m.InvalidSessionIdException):
                for row_number, row in reader.lazy_load_sheet():
                    self.controlador_1_empresa._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
            reader.close_workbook()

        self.assertEqual(self.controlador_1_empresa.resultados, [])  # nada foi registrado como "erro de portal"

    def test_erro_inesperado_numa_empresa_nao_impede_a_proxima_empresa(self):
        chamadas_de_empresa = []

        def procurar_inscricao(driver, empresa):
            chamadas_de_empresa.append(empresa.codigo)
            if empresa.codigo == 1:
                raise RuntimeError("falha inesperada e não prevista no portal")
            return True

        with mock.patch.object(self.controlador, "_procurar_inscricao_empresa", side_effect=procurar_inscricao), \
             mock.patch.object(self.controlador, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(self.controlador, "_navegar_tela_escrituracao"), \
             mock.patch.object(
                 self.controlador, "_processar_competencia",
                 lambda driver, reader, empresa, competencia: m.ResultadoMesEmpresa(empresa, competencia, encerrada=True),
             ), \
             mock.patch.object(self.controlador, "_abrir_modal_de_alteracao_de_inscricao"):
            reader = self._reader()
            for row_number, row in reader.lazy_load_sheet():
                try:
                    self.controlador._processar_linha(driver=object(), reader=reader, row_number=row_number, row=row)
                except Exception:  # o executar_processo() de verdade também blinda cada linha; replicamos aqui
                    pass
            reader.close_workbook()

        self.assertEqual(chamadas_de_empresa, [1, 2])  # a empresa 2 foi tentada mesmo com a falha na 1
        self.assertEqual(len(self.controlador.resultados), 2)  # só os 2 resultados da empresa 2 (1 por competência)


class TestNavegadorIndisponivel(unittest.TestCase):
    def test_reconhece_sessao_perdida(self):
        self.assertTrue(m.navegador_indisponivel(m.InvalidSessionIdException("invalid session id")))
        self.assertTrue(m.navegador_indisponivel(m.NoSuchWindowException("no such window")))
        self.assertTrue(m.navegador_indisponivel(
            m.WebDriverException("disconnected: not connected to DevTools (Session info: chrome=153)")
        ))
        self.assertTrue(m.navegador_indisponivel(ConnectionError("conexão recusada")))

    def test_erros_comuns_de_portal_nao_sao_confundidos_com_navegador_morto(self):
        self.assertFalse(m.navegador_indisponivel(m.NoSuchElementException("sem elemento")))
        self.assertFalse(m.navegador_indisponivel(m.TimeoutException("demorou")))
        self.assertFalse(m.navegador_indisponivel(m.StaleElementReferenceException("velho")))
        self.assertFalse(m.navegador_indisponivel(m.WebDriverException("qualquer outro problema")))
        self.assertFalse(m.navegador_indisponivel(RuntimeError("erro de programação")))


class TestCompetenciaNaoEncontrada(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_nao_enc_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.caminho = _planilha_fiscal(self.pasta, [(1, "EMPRESA A", "12345678000199")])
        self.controlador = _controlador(self.caminho, self.pasta / "saida", [date(2025, 5, 1)])

    def test_linha_do_mes_ausente_vira_problema_legivel_sem_nome_de_excecao(self):
        from python_spreadsheet_reader.readers import XLSXReader

        reader = XLSXReader(self.caminho)
        with mock.patch.object(self.controlador, "_preencher_widget_calendario"), \
             mock.patch.object(m, "retry_on_exception", side_effect=m.NoSuchElementException("linha não veio")):
            resultado = self.controlador._processar_competencia(mock.MagicMock(), reader, _empresa(linha=2), date(2025, 5, 1))
        reader.close_workbook()

        self.assertFalse(resultado.encerrada)
        self.assertEqual(resultado.problemas, [m.MSG_COMPETENCIA_NAO_ENCONTRADA])

        reader2 = XLSXReader(self.caminho)
        self.assertEqual(reader2.get_cell("Y2").value, "05/2025: COMPETÊNCIA NÃO ENCONTRADA NO PORTAL")
        reader2.close_workbook()


class TestRecuperacaoDeErros(unittest.TestCase):
    """Comportamento de `executar_processo()` quando o navegador morre ou o portal fica numa tela
    desconhecida — sempre com ChromeDriver falso."""

    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_rec_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.caminho = _planilha_fiscal(self.pasta, [(1, "EMPRESA A", "12345678000199"), (2, "EMPRESA B", "98765432000111")])

    def _rodar(self, controlador, driver_falso, **patches):
        classe = mock.Mock(return_value=driver_falso)
        with mock.patch.object(m, "ChromeDriver", classe), \
             mock.patch.object(controlador, "_login_iss_fortaleza"), \
             mock.patch.object(controlador, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(controlador, "_navegar_tela_escrituracao"), \
             mock.patch.object(controlador, "_abrir_modal_de_alteracao_de_inscricao"):
            with mock.patch.multiple(controlador, **patches):
                return controlador.executar_processo()

    def test_navegador_fechado_para_a_execucao_e_nao_tenta_as_empresas_restantes(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        empresas_tentadas = []

        def procurar(driver, empresa):
            empresas_tentadas.append(empresa.codigo)
            raise m.InvalidSessionIdException("invalid session id: session deleted as the browser has closed the connection")

        driver_falso = mock.Mock(name="ChromeDriverFalso")
        driver_falso.quit_driver.side_effect = m.InvalidSessionIdException("já fechado")  # quit() também falha: não pode estourar
        resultado = self._rodar(controlador, driver_falso, _procurar_inscricao_empresa=mock.Mock(side_effect=procurar))

        self.assertEqual(empresas_tentadas, [1])  # a empresa 2 NÃO foi tentada com o navegador morto
        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION)
        self.assertIn("navegador", resultado.status.message.lower())
        driver_falso.quit_driver.assert_called_once()

    def test_navegador_fechado_ainda_gera_relatorio_do_que_ja_foi_feito(self):
        if not CAMINHO_TEMPLATE.exists():
            self.skipTest("template do relatório não encontrado")
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])

        def procurar(driver, empresa):
            if empresa.codigo == 2:
                raise m.InvalidSessionIdException("invalid session id")
            return True

        resultado = self._rodar(
            controlador,
            mock.Mock(name="ChromeDriverFalso"),
            _procurar_inscricao_empresa=mock.Mock(side_effect=procurar),
            _processar_competencia=mock.Mock(
                side_effect=lambda driver, reader, empresa, competencia: m.ResultadoMesEmpresa(empresa, competencia, encerrada=True)
            ),
        )

        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION)
        self.assertEqual(resultado.processadas, 1)  # só a empresa 1 concluiu
        self.assertEqual(len(resultado.report_paths), 1)

    def test_erro_inesperado_por_empresa_recupera_o_portal_e_segue_para_a_proxima(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        empresas_tentadas = []

        def procurar(driver, empresa):
            empresas_tentadas.append(empresa.codigo)
            if empresa.codigo == 1:
                raise RuntimeError("falha inesperada")
            return True

        recuperar = mock.Mock()
        resultado = self._rodar(
            controlador,
            mock.Mock(name="ChromeDriverFalso"),
            _procurar_inscricao_empresa=mock.Mock(side_effect=procurar),
            _recuperar_para_proxima_empresa=recuperar,
            _processar_competencia=mock.Mock(
                side_effect=lambda driver, reader, empresa, competencia: m.ResultadoMesEmpresa(empresa, competencia, encerrada=True)
            ),
        )

        self.assertEqual(empresas_tentadas, [1, 2])
        recuperar.assert_called_once()
        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.SUCCESS)

    def test_se_nao_da_para_recuperar_o_portal_a_execucao_para_com_mensagem_clara(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        empresas_tentadas = []

        def procurar(driver, empresa):
            empresas_tentadas.append(empresa.codigo)
            raise RuntimeError("falha inesperada")

        resultado = self._rodar(
            controlador,
            mock.Mock(name="ChromeDriverFalso"),
            _procurar_inscricao_empresa=mock.Mock(side_effect=procurar),
            _recuperar_para_proxima_empresa=mock.Mock(side_effect=m.RecuperacaoImpossivelException("não consegui voltar à troca de empresa")),
        )

        self.assertEqual(empresas_tentadas, [1])  # parou em vez de gerar falhas em cascata
        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.ERROR_UNHANDLED_EXCEPTION)
        self.assertIn("não consegui voltar", resultado.status.message)

    def test_recuperacao_nao_faz_nada_se_a_troca_de_empresa_ja_esta_aberta(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        with mock.patch.object(controlador, "_modal_de_troca_de_inscricao_esta_aberto", return_value=True), \
             mock.patch.object(controlador, "_abrir_modal_de_alteracao_de_inscricao") as abrir:
            controlador._recuperar_para_proxima_empresa(mock.Mock())
        abrir.assert_not_called()

    def test_recuperacao_abre_a_troca_de_empresa_quando_ela_nao_esta_aberta(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        with mock.patch.object(controlador, "_modal_de_troca_de_inscricao_esta_aberto", return_value=False), \
             mock.patch.object(controlador, "_abrir_modal_de_alteracao_de_inscricao") as abrir:
            controlador._recuperar_para_proxima_empresa(mock.Mock())
        abrir.assert_called_once()

    def test_recuperacao_que_falha_levanta_recuperacao_impossivel(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        with mock.patch.object(controlador, "_modal_de_troca_de_inscricao_esta_aberto", return_value=False), \
             mock.patch.object(controlador, "_abrir_modal_de_alteracao_de_inscricao", side_effect=m.NoSuchElementException("sem botão")):
            with self.assertRaises(m.RecuperacaoImpossivelException):
                controlador._recuperar_para_proxima_empresa(mock.Mock())

    def test_recuperacao_com_navegador_morto_vira_navegador_indisponivel(self):
        controlador = _controlador(self.caminho, self.pasta / "saida", [date(2026, 1, 1)])
        with mock.patch.object(controlador, "_modal_de_troca_de_inscricao_esta_aberto", side_effect=m.InvalidSessionIdException("invalid session id")):
            with self.assertRaises(m.NavegadorIndisponivelException):
                controlador._recuperar_para_proxima_empresa(mock.Mock())


class TestExecutarProcessoIntegrado(unittest.TestCase):
    """`executar_processo()` de ponta a ponta, com um ChromeDriver totalmente falso (nunca abre
    navegador) e a leitura/escrita real da planilha e do relatório."""

    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="encerr_multi_full_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        self.caminho_planilha = _planilha_fiscal(self.pasta, [(1, "EMPRESA A", "12345678000199")])

    def test_execucao_completa_gera_resultado_e_relatorio_sem_abrir_navegador(self):
        controlador = _controlador(self.caminho_planilha, self.pasta / "saida", [date(2026, 1, 1)])

        driver_falso = mock.Mock(name="ChromeDriverFalso")
        classe_driver_falsa = mock.Mock(return_value=driver_falso)

        with mock.patch.object(m, "ChromeDriver", classe_driver_falsa), \
             mock.patch.object(controlador, "_login_iss_fortaleza"), \
             mock.patch.object(controlador, "_procurar_inscricao_empresa", return_value=True), \
             mock.patch.object(controlador, "_dar_ciencia_nas_mensagens_nao_lidas"), \
             mock.patch.object(controlador, "_navegar_tela_escrituracao"), \
             mock.patch.object(
                 controlador, "_processar_competencia",
                 lambda driver, reader, empresa, competencia: m.ResultadoMesEmpresa(
                     empresa, competencia, encerrada=True, dt_encerramento=date(2026, 1, 5)
                 ),
             ), \
             mock.patch.object(controlador, "_abrir_modal_de_alteracao_de_inscricao"):
            resultado = controlador.executar_processo()

        driver_falso.start_driver.assert_called_once()
        driver_falso.quit_driver.assert_called_once()
        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.SUCCESS)
        self.assertEqual((resultado.processadas, resultado.encerradas, resultado.problemas), (1, 1, 0))
        self.assertEqual(len(resultado.report_paths), 1)
        self.assertTrue(resultado.report_paths[0].exists())

    def test_cancelamento_fecha_o_navegador_e_nao_levanta(self):
        controlador = _controlador(self.caminho_planilha, self.pasta / "saida", [date(2026, 1, 1)])

        driver_falso = mock.Mock(name="ChromeDriverFalso")
        classe_driver_falsa = mock.Mock(return_value=driver_falso)

        def login_e_cancela(driver):
            raise m.UserStoppedThreadException()

        with mock.patch.object(m, "ChromeDriver", classe_driver_falsa), \
             mock.patch.object(controlador, "_login_iss_fortaleza", side_effect=login_e_cancela):
            resultado = controlador.executar_processo()

        driver_falso.quit_driver.assert_called_once()
        self.assertEqual(resultado.status.codigo, CodigoEncerramentoISS.USER_ENDED_PROCESS)
        self.assertEqual(resultado.processadas, 0)


if __name__ == "__main__":
    unittest.main()
