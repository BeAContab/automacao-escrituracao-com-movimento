"""Testes da cópia Multi-CNPJ do robô (escrituracao_multi_robo.py).

REGRA: nenhum teste toca o portal nem grava nota. O driver é sempre falso; a gravação do
módulo fica bloqueada por padrão e só é liberada, com `_clicar_por_id` trocado por um
gravador de chamadas, para provar o comportamento da guarda.
"""

import ast
import contextlib
import io
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

import escrituracao_multi_robo as robo

CAMINHO_MODULO = Path(robo.__file__)


class _DriverProibido:
    """Qualquer uso do driver derruba o teste: prova que nada foi feito no portal."""

    def __getattr__(self, nome):
        raise AssertionError(f"o driver não deveria ser usado (acesso a '{nome}')")


def _documento(**alteracoes) -> robo.DocumentoPortalISS:
    base = dict(
        arquivo_pdf="nota1.xml",
        cnpj_prestador="12345678000195",
        numero_nf="100",
        id_cnae_final="6201501",
        data_emissao="01/09/2026",
        descricao_servico="Desenvolvimento de software",
        uf_local_prestacao="CE",
        cidade_local_prestacao="Fortaleza",
        natureza_operacao="Tributação no município",
        iss_retido="Não",
        valor_servico="1000,00",
        nome_prestador="Prestador Exemplo LTDA",
        uf_prestador="SP",
        cidade_prestador="São Paulo",
        cep_prestador="01001000",
        logradouro_prestador="Rua A",
        bairro_prestador="Centro",
        regime_tributario="OUTROS",
        cnpj_tomador="00584628000181",
    )
    base.update(alteracoes)
    return robo.DocumentoPortalISS(**base)


class TestGravacaoBloqueada(unittest.TestCase):
    def test_bloqueada_por_padrao_e_nao_toca_o_driver(self):
        with self.assertRaises(robo.GravacaoBloqueadaError):
            robo._clicar_gravar_documento(_DriverProibido())

    def test_gravar_com_confirmacoes_tambem_bloqueia_antes_de_qualquer_clique(self):
        with self.assertRaises(robo.GravacaoBloqueadaError):
            robo._gravar_documento_com_confirmacoes(_DriverProibido(), "100")

    def test_liberada_apenas_dentro_do_bloco_e_volta_a_bloquear(self):
        chamadas = []
        with mock.patch.object(robo, "_clicar_por_id", lambda drv, ident, timeout=10: chamadas.append(ident)):
            with robo.gravacao_habilitada(True):
                robo._clicar_gravar_documento(_DriverProibido())
            self.assertEqual(chamadas, ["digitarDocumentoForm:j_id477"])
            with self.assertRaises(robo.GravacaoBloqueadaError):
                robo._clicar_gravar_documento(_DriverProibido())
        self.assertEqual(len(chamadas), 1)

    def test_bloco_com_false_continua_bloqueado(self):
        with robo.gravacao_habilitada(False):
            with self.assertRaises(robo.GravacaoBloqueadaError):
                robo._clicar_gravar_documento(_DriverProibido())

    def test_bloqueia_de_novo_mesmo_se_o_bloco_terminar_com_erro(self):
        with self.assertRaises(ValueError):
            with robo.gravacao_habilitada(True):
                raise ValueError("falha qualquer")
        with self.assertRaises(robo.GravacaoBloqueadaError):
            robo._clicar_gravar_documento(_DriverProibido())

    def test_ha_um_unico_caminho_de_gravacao_no_codigo(self):
        """Varre o fonte: o ID do botão Gravar só aparece dentro de `_clicar_gravar_documento`, e essa
        função só é chamada por `_gravar_documento_com_confirmacoes`. Qualquer novo caminho de gravação
        (por engano) faz este teste falhar."""
        fonte = CAMINHO_MODULO.read_text(encoding="utf-8")
        arvore = ast.parse(fonte)

        def funcao_que_contem(no_alvo):
            for no in ast.walk(arvore):
                if isinstance(no, ast.FunctionDef):
                    for filho in ast.walk(no):
                        if filho is no_alvo:
                            return no.name
            return None

        donos_do_id = set()
        chamadores = set()
        for no in ast.walk(arvore):
            if isinstance(no, ast.Constant) and isinstance(no.value, str) and "j_id477" in no.value:
                donos_do_id.add(funcao_que_contem(no))
            if isinstance(no, ast.Call) and isinstance(no.func, ast.Name) and no.func.id == "_clicar_gravar_documento":
                chamadores.add(funcao_que_contem(no))
        self.assertEqual(donos_do_id, {"_clicar_gravar_documento"})
        self.assertEqual(chamadores, {"_gravar_documento_com_confirmacoes"})
        chamadores_da_gravacao = {
            funcao_que_contem(no)
            for no in ast.walk(arvore)
            if isinstance(no, ast.Call) and isinstance(no.func, ast.Name) and no.func.id == "_gravar_documento_com_confirmacoes"
        }
        # Hoje ninguém chama; a partir do passo 2, só o laço das notas (e ele exige modo real).
        self.assertLessEqual(chamadores_da_gravacao, {"processar_notas"})


class TestIndependenciaDaCopia(unittest.TestCase):
    def test_nao_depende_do_modulo_da_funcao_individual_nem_do_tratamento_de_erros(self):
        fonte = CAMINHO_MODULO.read_text(encoding="utf-8")
        for proibido in ("iss_fortaleza_automacao", "tratamento_erros", "webdriver_manager", "ChromeDriverManager"):
            self.assertNotIn(proibido, "\n".join(l for l in fonte.splitlines() if l.lstrip().startswith(("import", "from"))))

    def test_nao_ha_print_nem_input_no_codigo_de_execucao(self):
        arvore = ast.parse(CAMINHO_MODULO.read_text(encoding="utf-8"))
        funcoes_io = [
            no.func.id
            for no in ast.walk(arvore)
            if isinstance(no, ast.Call) and isinstance(no.func, ast.Name) and no.func.id in ("print",)
        ]
        self.assertEqual(funcoes_io, [])


class TestSaidaEEventos(unittest.TestCase):
    def tearDown(self):
        robo.configurar_saida()

    def test_emitir_vai_para_o_callback_e_nao_para_o_terminal(self):
        recebidas = []
        robo.configurar_saida(callback_saida=recebidas.append)
        captura = io.StringIO()
        with contextlib.redirect_stdout(captura):
            robo._emitir("Processando", "linha", 3)
        self.assertEqual(recebidas, ["Processando linha 3"])
        self.assertEqual(captura.getvalue(), "")

    def test_sem_callback_a_mensagem_e_descartada_sem_erro(self):
        robo.configurar_saida()
        robo._emitir("qualquer coisa")
        robo.registrar_evento_execucao("marco", "ISS Fortaleza")

    def test_evento_vai_para_o_gancho_e_nao_para_o_tratamento_de_erros(self):
        import tratamento_erros

        eventos = []
        robo.configurar_saida(callback_evento=eventos.append)
        antes = list(tratamento_erros.PASTAS_LOGS)
        robo.registrar_evento_execucao("Botão Consultar acionado", "ISS Fortaleza")
        self.assertEqual(eventos, ["ISS Fortaleza: Botão Consultar acionado"])
        self.assertEqual(tratamento_erros.PASTAS_LOGS, antes)

    def test_erro_no_callback_nao_derruba_o_robo(self):
        def quebra(_msg):
            raise RuntimeError("falha no log")

        robo.configurar_saida(callback_saida=quebra, callback_evento=quebra)
        robo._emitir("ok")
        robo.registrar_evento_execucao("ok", "ctx")


class TestControle(unittest.TestCase):
    def tearDown(self):
        robo.configurar_controle()

    def test_cancelamento_levanta_a_excecao_do_modulo(self):
        robo.configurar_controle(callback_cancelamento=lambda: True)
        with self.assertRaises(robo.AutomacaoCanceladaError):
            robo._verificar_cancelamento()

    def test_sem_cancelamento_nao_levanta(self):
        robo.configurar_controle(callback_cancelamento=lambda: False)
        robo._verificar_cancelamento()

    def test_zerar_o_controle_remove_os_callbacks(self):
        robo.configurar_controle(callback_pausa=lambda: None, callback_cancelamento=lambda: True)
        robo.configurar_controle()
        robo._verificar_cancelamento()
        robo._verificar_pausa()

    def test_tempo_pausado_so_conta_bloqueio_de_verdade(self):
        robo.configurar_controle(callback_pausa=lambda: None)
        robo._verificar_pausa()
        self.assertEqual(robo.consumir_tempo_pausado(), 0.0)

    def test_tempo_pausado_acumula_e_zera_ao_consumir(self):
        relogio = iter([100.0, 103.5, 200.0, 201.5])  # 3,5 s + 1,5 s bloqueado
        with mock.patch.object(robo.time, "monotonic", lambda: next(relogio)):
            robo.configurar_controle(callback_pausa=lambda: None)
            robo._verificar_pausa()
            robo._verificar_pausa()
        self.assertAlmostEqual(robo.consumir_tempo_pausado(), 5.0)
        self.assertEqual(robo.consumir_tempo_pausado(), 0.0)


class TestClassificacaoENotas(unittest.TestCase):
    def test_nota_completa_de_fora_e_escriturada(self):
        self.assertEqual(robo.classificar_documento(_documento()), (robo.CATEGORIA_ESCRITURAR, []))

    def test_prestador_de_fortaleza_ce_nao_mei_e_ignorado(self):
        doc = _documento(uf_prestador="CE", cidade_prestador="Fortaleza")
        self.assertEqual(robo.classificar_documento(doc)[0], robo.CATEGORIA_FORTALEZA)

    def test_prestador_de_fortaleza_ce_mei_e_escriturado(self):
        doc = _documento(uf_prestador="ce", cidade_prestador="FORTALEZA", regime_tributario=" mei ")
        self.assertEqual(robo.classificar_documento(doc)[0], robo.CATEGORIA_ESCRITURAR)

    def test_fortaleza_tem_prioridade_sobre_campo_ausente(self):
        doc = _documento(uf_prestador="CE", cidade_prestador="Fortaleza", cep_prestador="")
        self.assertEqual(robo.classificar_documento(doc)[0], robo.CATEGORIA_FORTALEZA)

    def test_campo_obrigatorio_ausente_e_incompleta_e_lista_o_campo(self):
        categoria, ausentes = robo.classificar_documento(_documento(cep_prestador="  "))
        self.assertEqual(categoria, robo.CATEGORIA_INCOMPLETA)
        self.assertEqual(ausentes, ["CEP do Prestador"])

    def test_iss_retido_nao_reconhecido_e_incompleta(self):
        categoria, ausentes = robo.classificar_documento(_documento(iss_retido="talvez"))
        self.assertEqual(categoria, robo.CATEGORIA_INCOMPLETA)
        self.assertTrue(any("não reconhecido" in a for a in ausentes))

    def test_valor_zero_e_incompleta(self):
        self.assertEqual(robo.classificar_documento(_documento(valor_servico="0,00"))[0], robo.CATEGORIA_INCOMPLETA)

    def test_analise_conta_cada_categoria(self):
        docs = [
            _documento(numero_nf="1"),
            _documento(numero_nf="2"),
            _documento(numero_nf="3", uf_prestador="CE", cidade_prestador="Fortaleza"),
            _documento(numero_nf="4", cep_prestador=""),
        ]
        analise = robo.analisar_documentos(docs)
        self.assertEqual((analise.total, analise.a_escriturar, analise.fortaleza, analise.incompletas), (4, 2, 1, 1))
        self.assertIn("4 nota(s): 2 a escriturar", analise.descricao())

    def test_analise_de_lista_vazia(self):
        self.assertEqual(robo.analisar_documentos([]).total, 0)

    def test_chave_da_nota_e_estavel_e_ignora_mascara_do_cnpj(self):
        a = _documento(cnpj_prestador="12.345.678/0001-95")
        b = _documento(cnpj_prestador="12345678000195")
        self.assertEqual(robo.chave_da_nota(a), robo.chave_da_nota(b))
        self.assertNotEqual(robo.chave_da_nota(a), robo.chave_da_nota(_documento(numero_nf="101")))

    def test_agrupa_por_tomador_na_ordem_da_planilha(self):
        docs = [
            _documento(numero_nf="1", cnpj_tomador="00.584.628/0001-81"),
            _documento(numero_nf="2", cnpj_tomador="21345512000160"),
            _documento(numero_nf="3", cnpj_tomador="00584628000181"),
        ]
        grupos = robo.agrupar_por_tomador(docs)
        self.assertEqual(list(grupos), ["00584628000181", "21345512000160"])
        self.assertEqual([d.numero_nf for d in grupos["00584628000181"]], ["1", "3"])


class TestCarregarDocumentos(unittest.TestCase):
    CABECALHOS = [
        "ARQUIVO_XML", "CNPJ_TOMADOR", "CNPJ_PRESTADOR", "NUMERO_NF", "ID_CNAE_FINAL", "DATA_EMISSAO",
        "DESCRICAO_SERVICO", "UF_LOCAL_PRESTACAO", "CIDADE_LOCAL_PRESTACAO", "NATUREZA_OPERACAO",
        "ISS_RETIDO", "VALOR_SERVICO",
    ]

    def _planilha(self, pasta: Path, cabecalhos, linhas) -> Path:
        wb = Workbook()
        ws = wb.active
        ws.append(cabecalhos)
        for linha in linhas:
            ws.append(linha)
        caminho = pasta / "notas.xlsx"
        wb.save(caminho)
        wb.close()
        return caminho

    def test_le_cnpj_tomador_e_repoe_zeros_de_numero(self):
        import shutil
        import tempfile

        pasta = Path(tempfile.mkdtemp(prefix="robo_"))
        try:
            caminho = self._planilha(
                pasta,
                self.CABECALHOS,
                [
                    ["a.xml", 584628000181, "12345678000195", "10", "6201501", "01/09/2026", "x", "CE", "Fortaleza", "Tributação", "Não", 100.5],
                    ["b.xml", "21.345.512/0001-60", "12345678000195", "11", "6201501", "01/09/2026", "x", "CE", "Fortaleza", "Tributação", "Não", 50],
                ],
            )
            docs = robo.carregar_documentos_xlsx(caminho)
        finally:
            shutil.rmtree(pasta, ignore_errors=True)
        self.assertEqual([d.cnpj_tomador for d in docs], ["00584628000181", "21.345.512/0001-60"])
        self.assertEqual(robo.agrupar_por_tomador(docs).keys(), {"00584628000181", "21345512000160"})

    def test_planilha_sem_a_coluna_cnpj_tomador_continua_valida(self):
        import shutil
        import tempfile

        pasta = Path(tempfile.mkdtemp(prefix="robo_"))
        try:
            cab = [c for c in self.CABECALHOS if c != "CNPJ_TOMADOR"]
            caminho = self._planilha(
                pasta, cab, [["a.xml", "12345678000195", "10", "6201501", "01/09/2026", "x", "CE", "Fortaleza", "T", "Não", 1]]
            )
            docs = robo.carregar_documentos_xlsx(caminho)
        finally:
            shutil.rmtree(pasta, ignore_errors=True)
        self.assertEqual(docs[0].cnpj_tomador, "")


if __name__ == "__main__":
    unittest.main()
