"""Testes do passo 2 (escriturar/simular as notas de cada empresa no Multi-CNPJ).

REGRA DESTA SUÍTE: NENHUM teste grava nota nem toca o portal.
- O driver é sempre falso.
- Toda função de preenchimento/gravação do robô é trocada por um dublê (`unittest.mock`).
- Um `setUp` comum troca `_clicar_gravar_documento` por um "detonador": se qualquer teste tentar clicar em
  Gravar de verdade, o teste falha na hora.
"""

import json
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import escrituracao_multi_cnpj as m
import escrituracao_multi_robo as robo
from tests.test_escrituracao_multi_cnpj import AcessoFalso, CNPJ_A, CNPJ_B, _planilha
from tests.test_escrituracao_multi_robo import _documento

COMPETENCIA = robo.construir_competencia(9, 2026)
ROTULO = COMPETENCIA.rotulo  # "09/2026"

CABECALHOS = [
    "ARQUIVO_XML", "CNPJ_TOMADOR", "LOGIN", "SENHA", "CNPJ_PRESTADOR", "NUMERO_NF", "ID_CNAE_FINAL",
    "DATA_EMISSAO", "DESCRICAO_SERVICO", "UF_LOCAL_PRESTACAO", "CIDADE_LOCAL_PRESTACAO", "NATUREZA_OPERACAO",
    "ISS_RETIDO", "VALOR_SERVICO", "NOME_PRESTADOR", "UF_PRESTADOR", "CIDADE_PRESTADOR", "CEP_PRESTADOR",
    "LOGRADOURO_PRESTADOR", "BAIRRO_PRESTADOR", "REGIME_TRIBUTARIO",
]


def _linha(tomador, login, senha, numero, uf="SP", cidade="São Paulo", cep="01001000"):
    return [
        f"n{numero}.xml", tomador, login, senha, "12345678000195", str(numero), "6201501", "01/09/2026", "Serviço",
        "CE", "Fortaleza", "Tributação", "Não", 100.0, "Prestador LTDA", uf, cidade, cep, "Rua A", "Centro", "OUTROS",
    ]


class _Elemento:
    def __init__(self, texto="", visivel=True):
        self.text = texto
        self._visivel = visivel

    def is_displayed(self):
        return self._visivel

    def click(self):
        pass


class _DriverFalso:
    """Só sabe responder à busca de modais de confirmação; qualquer outra coisa derruba o teste."""

    def __init__(self, modais=None):
        self.modais = modais or []

    def find_elements(self, by, seletor):
        return list(self.modais)

    def __getattr__(self, nome):
        raise AssertionError(f"o driver não deveria ser usado (acesso a '{nome}')")


class _Base(unittest.TestCase):
    """Base: pasta temporária, saída do robô capturada e DETONADOR de gravação."""

    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="etapa2_"))
        self.saida = []
        self.marcos = []
        robo.configurar_saida(self.saida.append, self.marcos.append)
        robo.configurar_controle()

        def detonador(*_a, **_k):
            raise AssertionError("TENTATIVA DE CLICAR EM GRAVAR NUM TESTE — isto nunca pode acontecer")

        parar_clique = mock.patch.object(robo, "_clicar_gravar_documento", detonador)
        parar_clique.start()
        self.addCleanup(parar_clique.stop)
        self.addCleanup(shutil.rmtree, self.pasta, True)
        self.addCleanup(robo.configurar_saida)
        self.addCleanup(robo.configurar_controle)

    def registro(self, modo="teste"):
        return robo.RegistroNotas(self.pasta / "logs", ROTULO, modo, agora=lambda: datetime(2026, 9, 21, 12, 0, 0))


class TestRegistroNotas(_Base):
    def test_todos_os_arquivos_ficam_na_mesma_pasta_sem_subpasta(self):
        reg = self.registro()
        doc = _documento()
        reg.sucesso(doc)
        reg.fortaleza(doc)
        reg.incompleta(doc, ["CEP do Prestador"])
        reg.duplicada(doc)
        reg.erro(doc, "ERRO ao processar: x")
        reg.simulada(doc)
        pasta = self.pasta / "logs"
        self.assertEqual(
            sorted(p.name for p in pasta.iterdir()),
            sorted([
                robo.ARQUIVO_SUCESSO, robo.ARQUIVO_FORTALEZA, robo.ARQUIVO_INCOMPLETAS,
                robo.ARQUIVO_DUPLICADAS, robo.ARQUIVO_ERROS, robo.ARQUIVO_SIMULADAS,
            ]),
        )
        self.assertFalse(any(p.is_dir() for p in pasta.iterdir()))  # sem subpastas

    def test_cada_linha_leva_o_cnpj_da_empresa(self):
        reg = self.registro()
        reg.sucesso(_documento(cnpj_tomador="00584628000181"))
        texto = (self.pasta / "logs" / robo.ARQUIVO_SUCESSO).read_text(encoding="utf-8")
        self.assertIn("[00.584.628/0001-81]", texto)
        self.assertIn("[2026-09-21 12:00:00]", texto)

    def test_acrescenta_com_cabecalho_por_execucao_sem_apagar_o_anterior(self):
        doc = _documento()
        self.registro("Escriturar de verdade").sucesso(doc)
        reg2 = self.registro("Escriturar de verdade")
        reg2.sucesso(_documento(numero_nf="200"))
        reg2.sucesso(_documento(numero_nf="201"))
        texto = (self.pasta / "logs" / robo.ARQUIVO_SUCESSO).read_text(encoding="utf-8")
        self.assertEqual(texto.count("===== Execução"), 2)  # 1 cabeçalho por execução, não por linha
        self.assertEqual(texto.count("As seguintes notas fiscais foram escrituradas"), 1)  # introdução só na criação
        for nf in ("- 100", "- 200", "- 201"):
            self.assertIn(nf, texto)
        self.assertIn("competência 09/2026 | modo: Escriturar de verdade", texto)

    def test_incompleta_lista_os_campos_ausentes(self):
        self.registro().incompleta(_documento(), ["CEP do Prestador", "Bairro do Prestador"])
        texto = (self.pasta / "logs" / robo.ARQUIVO_INCOMPLETAS).read_text(encoding="utf-8")
        self.assertIn("-> Ausente(s): CEP do Prestador, Bairro do Prestador", texto)

    def test_falha_de_escrita_nao_levanta_e_avisa(self):
        arquivo_no_lugar_da_pasta = self.pasta / "logs"
        arquivo_no_lugar_da_pasta.write_text("sou um arquivo", encoding="utf-8")
        self.registro().sucesso(_documento())  # não pode levantar
        self.assertTrue(any("Não consegui atualizar o arquivo de log" in s for s in self.saida))


class TestProcessarNotasSimulacao(_Base):
    def _rodar(self, docs, **extra):
        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT) as duplos, \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes",
                               side_effect=AssertionError("gravar não pode ser chamado na simulação")) as gravar, \
             mock.patch.object(robo.time, "sleep"):
            resultado = robo.processar_notas(_DriverFalso(), docs, COMPETENCIA, self.registro("Simulação"),
                                             simulacao=True, **extra)
        return resultado, duplos, gravar

    def test_simulacao_preenche_mas_nunca_grava(self):
        docs = [_documento(numero_nf="1"), _documento(numero_nf="2")]
        resultado, duplos, gravar = self._rodar(docs)
        self.assertEqual((resultado.simuladas, resultado.escrituradas, resultado.erros), (2, 0, 0))
        self.assertEqual(duplos["preencher_dados_prestador"].call_count, 2)
        self.assertEqual(duplos["preencher_documento_servico"].call_count, 2)
        gravar.assert_not_called()
        self.assertEqual(duplos["_clicar_com_espera"].call_count, 2)  # descartou cada nota com "Novo Documento"

    def test_simulacao_mantem_a_gravacao_bloqueada_durante_e_depois(self):
        durante = []

        def espiar(*_a, **_k):
            durante.append(robo._GRAVACAO_PERMITIDA)

        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "preencher_documento_servico", espiar), \
             mock.patch.object(robo.time, "sleep"):
            robo.processar_notas(_DriverFalso(), [_documento()], COMPETENCIA, self.registro(), simulacao=True)
        self.assertEqual(durante, [False])
        self.assertFalse(robo._GRAVACAO_PERMITIDA)

    def test_simulacao_registra_em_simuladas_e_nao_no_log_de_sucesso(self):
        self._rodar([_documento()])
        pasta = self.pasta / "logs"
        self.assertTrue((pasta / robo.ARQUIVO_SIMULADAS).exists())
        self.assertFalse((pasta / robo.ARQUIVO_SUCESSO).exists())
        self.assertFalse((pasta / robo.ARQUIVO_DUPLICADAS).exists())

    def test_simulacao_nao_chama_o_callback_de_persistencia(self):
        chamado = []
        self._rodar([_documento()], ao_escriturar=chamado.append)
        self.assertEqual(chamado, [])

    def test_simulacao_interrompe_se_o_portal_pedir_confirmacao_ao_descartar(self):
        modal = _Elemento("Deseja descartar as alterações?")
        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT), mock.patch.object(robo.time, "sleep"):
            with self.assertRaises(robo.ModalInesperadoError) as ctx:
                robo.processar_notas(_DriverFalso([modal]), [_documento()], COMPETENCIA, self.registro(), simulacao=True)
        self.assertIn("Não cliquei em nada", str(ctx.exception))
        self.assertIn("descartar", str(ctx.exception))

    def test_modal_invisivel_na_simulacao_e_ignorado(self):
        modal_oculto = _Elemento("qualquer", visivel=False)
        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT), mock.patch.object(robo.time, "sleep"):
            resultado = robo.processar_notas(_DriverFalso([modal_oculto]), [_documento()], COMPETENCIA,
                                             self.registro(), simulacao=True)
        self.assertEqual(resultado.simuladas, 1)


class TestProcessarNotasReal(_Base):
    """Modo 'escriturar': a gravação é um DUBLÊ (`_gravar_documento_com_confirmacoes` trocada); nada vai ao portal."""

    def _rodar(self, docs, gravar, **extra):
        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar), \
             mock.patch.object(robo.time, "sleep"):
            return robo.processar_notas(_DriverFalso(), docs, COMPETENCIA, self.registro("Escriturar de verdade"),
                                        simulacao=False, **extra)

    def test_grava_com_a_gravacao_liberada_so_durante_o_laco(self):
        estado = []

        def gravar(driver, nf, timeout=15):
            estado.append(robo._GRAVACAO_PERMITIDA)
            return "sucesso"

        persistidas = []
        resultado = self._rodar([_documento(numero_nf="1"), _documento(numero_nf="2")], gravar,
                                ao_escriturar=persistidas.append)
        self.assertEqual(estado, [True, True])
        self.assertFalse(robo._GRAVACAO_PERMITIDA)  # voltou a bloquear
        self.assertEqual(resultado.escrituradas, 2)
        self.assertEqual(persistidas, [robo.chave_da_nota(_documento(numero_nf="1")), robo.chave_da_nota(_documento(numero_nf="2"))])
        texto = (self.pasta / "logs" / robo.ARQUIVO_SUCESSO).read_text(encoding="utf-8")
        self.assertIn("- 1", texto)

    def test_duplicata_do_portal_e_registrada_e_nao_conta_como_escriturada(self):
        persistidas = []
        resultado = self._rodar([_documento()], lambda d, nf, timeout=15: "duplicata", ao_escriturar=persistidas.append)
        self.assertEqual((resultado.duplicadas, resultado.escrituradas), (1, 0))
        self.assertEqual(persistidas, [])
        self.assertTrue((self.pasta / "logs" / robo.ARQUIVO_DUPLICADAS).exists())

    def test_pula_notas_ja_escrituradas_numa_execucao_anterior(self):
        chamadas = []

        def gravar(driver, nf, timeout=15):
            chamadas.append(nf)
            return "sucesso"

        ja = {robo.chave_da_nota(_documento(numero_nf="1"))}
        resultado = self._rodar([_documento(numero_nf="1"), _documento(numero_nf="2")], gravar, chaves_ja_escrituradas=ja)
        self.assertEqual(chamadas, ["2"])
        self.assertEqual((resultado.ja_escrituradas, resultado.escrituradas), (1, 1))

    def test_fortaleza_e_incompletas_nunca_chegam_ao_portal(self):
        docs = [
            _documento(numero_nf="1", uf_prestador="CE", cidade_prestador="Fortaleza"),
            _documento(numero_nf="2", cep_prestador=""),
            _documento(numero_nf="3"),
        ]
        chamadas = []

        def gravar(driver, nf, timeout=15):
            chamadas.append(nf)
            return "sucesso"

        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT) as duplos, \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar), mock.patch.object(robo.time, "sleep"):
            resultado = robo.processar_notas(_DriverFalso(), docs, COMPETENCIA, self.registro(), simulacao=False)
        self.assertEqual(duplos["preencher_dados_prestador"].call_count, 1)  # só a nota 3
        self.assertEqual(chamadas, ["3"])
        self.assertEqual((resultado.ignoradas_fortaleza, resultado.incompletas, resultado.escrituradas), (1, 1, 1))
        pasta = self.pasta / "logs"
        self.assertTrue((pasta / robo.ARQUIVO_FORTALEZA).exists())
        self.assertIn("CEP do Prestador", (pasta / robo.ARQUIVO_INCOMPLETAS).read_text(encoding="utf-8"))

    def test_erro_em_uma_nota_nao_interrompe_as_outras_e_e_contado(self):
        def gravar(driver, nf, timeout=15):
            if nf == "2":
                raise RuntimeError("portal recusou")
            return "sucesso"

        resultado = self._rodar([_documento(numero_nf=str(i)) for i in (1, 2, 3)], gravar)
        self.assertEqual((resultado.escrituradas, resultado.erros), (2, 1))
        self.assertIn("portal recusou", (self.pasta / "logs" / robo.ARQUIVO_ERROS).read_text(encoding="utf-8"))

    def test_erro_ao_preencher_prestador_conta_como_erro_e_segue(self):
        def prestador(driver, doc, depuracao=False):
            if doc.numero_nf == "1":
                raise RuntimeError("prestador não abriu")

        with mock.patch.multiple(robo, preencher_documento_servico=mock.DEFAULT, _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "preencher_dados_prestador", prestador), \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes", lambda d, nf, timeout=15: "sucesso"), \
             mock.patch.object(robo.time, "sleep"):
            resultado = robo.processar_notas(_DriverFalso(), [_documento(numero_nf="1"), _documento(numero_nf="2")],
                                             COMPETENCIA, self.registro(), simulacao=False)
        self.assertEqual((resultado.erros, resultado.escrituradas), (1, 1))

    def test_cancelamento_interrompe_o_laco(self):
        robo.configurar_controle(callback_cancelamento=lambda: True)
        with self.assertRaises(robo.AutomacaoCanceladaError):
            self._rodar([_documento()], lambda d, nf, timeout=15: "sucesso")
        self.assertFalse(robo._GRAVACAO_PERMITIDA)  # mesmo cancelado, volta a bloquear

    def test_progresso_e_chamado_por_nota(self):
        vistos = []
        self._rodar([_documento(numero_nf="1"), _documento(numero_nf="2")], lambda d, nf, timeout=15: "sucesso",
                    callback_progresso=lambda i, n: vistos.append((i, n)))
        self.assertEqual(vistos, [(1, 2), (2, 2)])

    def test_sessao_expirada_apos_pausa_longa_interrompe_a_empresa(self):
        with mock.patch.object(robo, "consumir_tempo_pausado", return_value=700.0):
            with self.assertRaises(robo.SessaoExpiradaError):
                self._rodar([_documento()], lambda d, nf, timeout=15: "sucesso", sessao_ativa=lambda: False)
        self.assertFalse(robo._GRAVACAO_PERMITIDA)

    def test_sessao_ainda_ativa_apos_pausa_longa_continua(self):
        with mock.patch.object(robo, "consumir_tempo_pausado", return_value=700.0):
            resultado = self._rodar([_documento()], lambda d, nf, timeout=15: "sucesso", sessao_ativa=lambda: True)
        self.assertEqual(resultado.escrituradas, 1)

    def test_pausa_curta_nao_consulta_a_sessao(self):
        consultas = []
        with mock.patch.object(robo, "consumir_tempo_pausado", return_value=30.0):
            self._rodar([_documento()], lambda d, nf, timeout=15: "sucesso", sessao_ativa=lambda: consultas.append(1) or True)
        self.assertEqual(consultas, [])

    def test_descricao_do_resultado(self):
        r = robo.ResultadoNotas(escrituradas=5, duplicadas=1, ignoradas_fortaleza=2, incompletas=1, ja_escrituradas=3, erros=1)
        self.assertEqual(
            r.descricao(False),
            "5 escriturada(s), 1 duplicada(s) (o portal já tinha), 2 ignorada(s) (prestador de Fortaleza/CE não MEI), "
            "1 incompleta(s), 3 já escriturada(s) antes, 1 com erro",
        )
        self.assertTrue(robo.ResultadoNotas(simuladas=4).descricao(True).startswith("4 simulada(s) (nada foi gravado)"))


class TestPrepararTela(_Base):
    """Sequência de cliques até 'Digitar Documento' — sem portal: todos os passos são dublês."""

    class _Espera:
        def __init__(self, *_a, **_k):
            pass

        def until(self, _condicao):
            return _Elemento()

    def _cliques_e_patches(self, botao_escriturar=True):
        cliques = []
        patches = [
            mock.patch.object(robo, "WebDriverWait", self._Espera),
            mock.patch.object(robo, "_clicar_por_id", lambda drv, ident, timeout=10: cliques.append(ident)),
            mock.patch.object(robo, "selecionar_competencia_na_tela_richfaces"),
            mock.patch.object(robo, "aguardar_botao_escriturar", lambda drv, timeout=120: botao_escriturar),
            mock.patch.object(robo, "aguardar_tela_escrituracao_fiscal"),
            mock.patch.object(robo, "clicar_aba_servicos_tomados"),
            mock.patch.object(robo, "aguardar_tela_digitar_documento"),
            mock.patch.object(robo.time, "sleep"),
        ]
        return cliques, patches

    def _driver(self):
        class D:
            def find_elements(self, *_a):
                return []

        return D()

    def test_sequencia_de_cliques_ate_digitar_documento_sem_nunca_gravar(self):
        cliques, patches = self._cliques_e_patches()
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        robo.preparar_tela_digitar_documento(self._driver(), COMPETENCIA)
        self.assertEqual(
            cliques,
            [
                "formMenuTopo:menuEscrituracao:j_id80",
                "manterEscrituracaoForm:btnConsultar",
                "manterEscrituracaoForm:dataTable:0:linkEscriturar",
                "servico_tomado_form:seamj_id849",
            ],
        )
        self.assertNotIn("digitarDocumentoForm:j_id477", cliques)

    def test_competencia_bloqueada_levanta_erro_e_nao_abre_a_escrituracao(self):
        cliques, patches = self._cliques_e_patches(botao_escriturar=False)
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        with self.assertRaises(robo.CompetenciaSemEscriturarDisponivelError) as ctx:
            robo.preparar_tela_digitar_documento(self._driver(), COMPETENCIA)
        self.assertIn("09/2026", str(ctx.exception))
        self.assertNotIn("manterEscrituracaoForm:dataTable:0:linkEscriturar", cliques)

    def test_cancelamento_antes_de_navegar(self):
        cliques, patches = self._cliques_e_patches()
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        robo.configurar_controle(callback_cancelamento=lambda: True)
        with self.assertRaises(robo.AutomacaoCanceladaError):
            robo.preparar_tela_digitar_documento(self._driver(), COMPETENCIA)
        self.assertEqual(cliques, [])


class TestEstadoPorCompetencia(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="estado_"))
        self.planilha = self.pasta / "notas.xlsx"
        self.addCleanup(shutil.rmtree, self.pasta, True)

    def test_competencias_diferentes_nao_se_misturam(self):
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "09/2026")
        m.registrar_nota_escriturada(self.planilha, CNPJ_A, "chave1", "09/2026")
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "10/2026"), set())
        self.assertEqual(m.ler_notas_escrituradas(self.planilha, CNPJ_A, "10/2026"), set())
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "09/2026"), {CNPJ_A})
        self.assertEqual(m.ler_notas_escrituradas(self.planilha, CNPJ_A, "09/2026"), {"chave1"})

    def test_notas_sao_por_empresa_e_sem_duplicar(self):
        m.registrar_nota_escriturada(self.planilha, CNPJ_A, "k", "09/2026")
        m.registrar_nota_escriturada(self.planilha, CNPJ_A, "k", "09/2026")
        m.registrar_nota_escriturada(self.planilha, CNPJ_B, "k2", "09/2026")
        self.assertEqual(m.ler_notas_escrituradas(self.planilha, CNPJ_A, "09/2026"), {"k"})
        self.assertEqual(m.ler_notas_escrituradas(self.planilha, CNPJ_B, "09/2026"), {"k2"})
        dados = json.loads(m.caminho_estado(self.planilha).read_text(encoding="utf-8"))
        self.assertEqual(dados["competencias"]["09/2026"]["notas"][CNPJ_A], ["k"])

    def test_arquivo_corrompido_vira_estado_vazio_sem_erro(self):
        m.caminho_estado(self.planilha).write_text("{isto não é json", encoding="utf-8")
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "09/2026"), set())
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "09/2026")  # sobrescreve com estado válido
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "09/2026"), {CNPJ_A})

    def test_formato_inesperado_vira_estado_vazio(self):
        m.caminho_estado(self.planilha).write_text(json.dumps(["lista", "solta"]), encoding="utf-8")
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "09/2026"), set())

    def test_gravacao_atomica_nao_deixa_arquivo_temporario(self):
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "09/2026")
        self.assertEqual([p.name for p in self.pasta.iterdir()], [m.caminho_estado(self.planilha).name])

    def test_limpar_estado_de_uma_competencia_preserva_as_outras_e_apaga_o_arquivo_se_vazio(self):
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "09/2026")
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "10/2026")
        m.limpar_estado(self.planilha, "09/2026")
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "09/2026"), set())
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, "10/2026"), {CNPJ_A})
        m.limpar_estado(self.planilha, "10/2026")
        self.assertFalse(m.caminho_estado(self.planilha).exists())

    def test_limpar_tudo(self):
        m.marcar_empresa_concluida(self.planilha, CNPJ_A, "09/2026")
        m.limpar_estado(self.planilha)
        self.assertFalse(m.caminho_estado(self.planilha).exists())


class TestOrquestradorComEscrituracao(_Base):
    """Orquestrador + gancho real (`criar_escriturador_de_empresa`), com o robô inteiro dublado."""

    def setUp(self):
        super().setUp()
        self.planilha = _planilha(
            self.pasta,
            [
                _linha(CNPJ_A, "11111111111", "SenhaSecretaA", 1),
                _linha(CNPJ_A, "11111111111", "SenhaSecretaA", 2, uf="CE", cidade="Fortaleza"),
                _linha(CNPJ_A, "11111111111", "SenhaSecretaA", 3, cep=""),
                _linha(CNPJ_B, "22222222222", "SenhaSecretaB", 4),
                _linha(CNPJ_B, "22222222222", "SenhaSecretaB", 5),
            ],
            cabecalhos=CABECALHOS,
        )
        self.logs = []
        self.log = lambda msg, erro=False: self.logs.append((msg, erro))
        self.progresso = []
        self.gravadas = []

    def _executar(self, *, simulacao, falhar_nf=(), acesso=None, competencia_rotulo=ROTULO, docs=None):
        docs = docs if docs is not None else m.carregar_documentos_por_empresa(self.planilha)
        registro = robo.RegistroNotas(self.pasta, ROTULO, "Simulação" if simulacao else "Escriturar de verdade")
        escriturador = m.criar_escriturador_de_empresa(
            docs, COMPETENCIA, self.planilha, registro, simulacao=simulacao,
            callback_log=self.log, callback_progresso=lambda p, s: self.progresso.append((p, s)),
        )

        def gravar(driver, nf, timeout=15):
            if nf in falhar_nf:
                raise RuntimeError("portal recusou")
            self.gravadas.append(nf)
            return "sucesso"

        acesso = acesso or AcessoFalso()
        acesso.driver = _DriverFalso()
        with mock.patch.multiple(robo, preparar_tela_digitar_documento=mock.DEFAULT, preencher_dados_prestador=mock.DEFAULT,
                                 preencher_documento_servico=mock.DEFAULT, _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar), mock.patch.object(robo.time, "sleep"):
            return m.executar_escrituracao_multi_cnpj(
                self.planilha, lambda: acesso, processar_empresa=escriturador, callback_log=self.log,
                callback_progresso=lambda p, s: self.progresso.append((p, s)), simulacao=simulacao,
                competencia=ROTULO, analisar_empresa=m.criar_analisador_de_empresa(docs),
            )

    def test_agrupamento_por_tomador_usa_a_mesma_normalizacao_das_empresas(self):
        docs = m.carregar_documentos_por_empresa(self.planilha)
        self.assertEqual(sorted(docs), sorted([CNPJ_A, CNPJ_B]))
        self.assertEqual([d.numero_nf for d in docs[CNPJ_A]], ["1", "2", "3"])

    def test_pre_analise_offline_aparece_no_log_antes_de_abrir_o_portal(self):
        self._executar(simulacao=True)
        textos = [msg for msg, _ in self.logs]
        pre = [t for t in textos if "Pré-análise da planilha" in t]
        self.assertEqual(len(pre), 2)
        self.assertIn("3 nota(s): 1 a escriturar, 1 ignorada(s)", pre[0])
        self.assertIn("1 incompleta(s)", pre[0])
        self.assertLess(textos.index(pre[0]), next(i for i, t in enumerate(textos) if "entrando no portal" in t))

    def test_simulacao_ponta_a_ponta_nao_grava_nem_cria_estado(self):
        resultado = self._executar(simulacao=True)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_SIMULADA, m.STATUS_SIMULADA])
        self.assertEqual(self.gravadas, [])  # `gravar` nunca foi chamado
        self.assertFalse(m.caminho_estado(self.planilha).exists())
        self.assertTrue((self.pasta / robo.ARQUIVO_SIMULADAS).exists())
        self.assertFalse((self.pasta / robo.ARQUIVO_SUCESSO).exists())
        self.assertIn("SIMULAÇÃO", "\n".join(msg for msg, _ in self.logs))

    def test_simulacao_nao_faz_uma_execucao_real_pular_empresas(self):
        self._executar(simulacao=True)
        real = self._executar(simulacao=False)
        self.assertEqual([r.status for r in real.empresas], [m.STATUS_CONCLUIDA, m.STATUS_CONCLUIDA])
        self.assertEqual(sorted(self.gravadas), ["1", "4", "5"])

    def test_escriturar_marca_notas_e_empresas_no_estado(self):
        resultado = self._executar(simulacao=False)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_CONCLUIDA, m.STATUS_CONCLUIDA])
        self.assertEqual(sorted(self.gravadas), ["1", "4", "5"])  # Fortaleza (2) e incompleta (3) não vão ao portal
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, ROTULO), {CNPJ_A, CNPJ_B})
        self.assertEqual(len(m.ler_notas_escrituradas(self.planilha, CNPJ_B, ROTULO)), 2)
        self.assertIn("1 escriturada(s)", resultado.empresas[0].mensagem)

    def test_empresa_com_nota_que_falhou_fica_com_pendencias_e_so_ela_e_refeita(self):
        r1 = self._executar(simulacao=False, falhar_nf=("5",))
        self.assertEqual([r.status for r in r1.empresas], [m.STATUS_CONCLUIDA, m.STATUS_COM_PENDENCIAS])
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, ROTULO), {CNPJ_A})  # B NÃO é marcada
        self.assertEqual(len(m.ler_notas_escrituradas(self.planilha, CNPJ_B, ROTULO)), 1)  # a nota 4 ficou salva

        self.gravadas.clear()
        r2 = self._executar(simulacao=False)
        self.assertEqual([r.status for r in r2.empresas], [m.STATUS_JA_CONCLUIDA, m.STATUS_CONCLUIDA])
        self.assertEqual(self.gravadas, ["5"])  # só a nota que falhou foi refeita; a 4 não se repete

    def test_outra_competencia_nao_herda_o_estado(self):
        self._executar(simulacao=False)
        self.gravadas.clear()
        docs = m.carregar_documentos_por_empresa(self.planilha)
        registro = robo.RegistroNotas(self.pasta, "10/2026", "Escriturar de verdade")
        outra = robo.construir_competencia(10, 2026)
        escriturador = m.criar_escriturador_de_empresa(
            docs, outra, self.planilha, registro, simulacao=False, callback_log=self.log)
        acesso = AcessoFalso()
        acesso.driver = _DriverFalso()
        with mock.patch.multiple(robo, preparar_tela_digitar_documento=mock.DEFAULT, preencher_dados_prestador=mock.DEFAULT,
                                 preencher_documento_servico=mock.DEFAULT, _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes",
                               lambda d, nf, timeout=15: self.gravadas.append(nf) or "sucesso"), \
             mock.patch.object(robo.time, "sleep"):
            r = m.executar_escrituracao_multi_cnpj(
                self.planilha, lambda: acesso, processar_empresa=escriturador, callback_log=self.log,
                competencia="10/2026")
        self.assertEqual([x.status for x in r.empresas], [m.STATUS_CONCLUIDA] * 2)
        self.assertEqual(sorted(self.gravadas), ["1", "4", "5"])

    def test_progresso_dentro_da_empresa_e_monotonico_e_termina_em_100(self):
        self._executar(simulacao=True)
        pcts = [p for p, _ in self.progresso]
        self.assertEqual(pcts, sorted(pcts))
        self.assertEqual(pcts[-1], 100)
        self.assertTrue(any("nota 1/" in s for _, s in self.progresso))

    def test_linhas_do_robo_no_log_levam_o_cnpj_da_empresa(self):
        self._executar(simulacao=True)
        do_robo = [msg for msg, _ in self.logs if "Processando linha" in msg or "SIMULADO" in msg]
        self.assertTrue(do_robo)
        self.assertTrue(all(msg.startswith("[00.584.628/0001-81]") or msg.startswith("[21.345.512/0001-60]") for msg in do_robo))

    def test_saida_do_robo_e_zerada_ao_terminar(self):
        self._executar(simulacao=True)
        robo._emitir("depois da execução")
        self.assertFalse(any("depois da execução" in msg for msg, _ in self.logs))

    def test_competencia_indisponivel_pula_a_empresa_e_segue(self):
        docs = m.carregar_documentos_por_empresa(self.planilha)
        escriturador = m.criar_escriturador_de_empresa(
            docs, COMPETENCIA, self.planilha, self.registro(), simulacao=True, callback_log=self.log)
        acesso = AcessoFalso()
        acesso.driver = _DriverFalso()

        def preparar(driver, competencia):
            raise robo.CompetenciaSemEscriturarDisponivelError("A competência 09/2026 está desabilitada")

        with mock.patch.object(robo, "preparar_tela_digitar_documento", preparar):
            r = m.executar_escrituracao_multi_cnpj(
                self.planilha, lambda: acesso, processar_empresa=escriturador, callback_log=self.log,
                simulacao=True, competencia=ROTULO)
        self.assertEqual([x.status for x in r.empresas], [m.STATUS_COMPETENCIA_INDISPONIVEL] * 2)
        self.assertEqual(sum(1 for c in acesso.chamadas if c[0] == "logout"), 2)  # logout sempre

    def test_sessao_expirada_e_traduzida_e_a_empresa_nao_e_marcada(self):
        docs = m.carregar_documentos_por_empresa(self.planilha)
        escriturador = m.criar_escriturador_de_empresa(
            docs, COMPETENCIA, self.planilha, self.registro(), simulacao=False, callback_log=self.log)
        acesso = AcessoFalso()
        acesso.driver = _DriverFalso()
        with mock.patch.object(robo, "preparar_tela_digitar_documento",
                               side_effect=robo.SessaoExpiradaError("a sessão do portal expirou")):
            r = m.executar_escrituracao_multi_cnpj(
                self.planilha, lambda: acesso, processar_empresa=escriturador, callback_log=self.log, competencia=ROTULO)
        self.assertEqual([x.status for x in r.empresas], [m.STATUS_SESSAO_EXPIRADA] * 2)
        self.assertEqual(m.ler_empresas_concluidas(self.planilha, ROTULO), set())

    def test_cancelamento_dentro_da_empresa_marca_cancelada_e_faz_logout(self):
        docs = m.carregar_documentos_por_empresa(self.planilha)
        escriturador = m.criar_escriturador_de_empresa(
            docs, COMPETENCIA, self.planilha, self.registro(), simulacao=False, callback_log=self.log)
        acesso = AcessoFalso()
        acesso.driver = _DriverFalso()
        with mock.patch.object(robo, "preparar_tela_digitar_documento",
                               side_effect=robo.AutomacaoCanceladaError("cancelado")):
            r = m.executar_escrituracao_multi_cnpj(
                self.planilha, lambda: acesso, processar_empresa=escriturador, callback_log=self.log, competencia=ROTULO)
        self.assertTrue(r.cancelado)
        self.assertEqual(r.empresas[0].status, m.STATUS_CANCELADA)
        self.assertIn(("logout", "11111111111"), acesso.chamadas)

    def test_senha_e_login_nao_vazam_no_log_da_escrituracao(self):
        self._executar(simulacao=False)
        tudo = "\n".join(msg for msg, _ in self.logs)
        for segredo in ("SenhaSecretaA", "SenhaSecretaB", "11111111111", "22222222222"):
            self.assertNotIn(segredo, tudo)
        for arquivo in self.pasta.iterdir():
            if arquivo.suffix in (".txt", ".json"):
                conteudo = arquivo.read_text(encoding="utf-8")
                for segredo in ("SenhaSecretaA", "SenhaSecretaB", "11111111111", "22222222222"):
                    self.assertNotIn(segredo, conteudo, arquivo.name)

    def test_empresa_sem_notas_na_planilha_de_robo_devolve_resumo_simples(self):
        escriturador = m.criar_escriturador_de_empresa(
            {}, COMPETENCIA, self.planilha, self.registro(), simulacao=True, callback_log=self.log)
        resumo = escriturador(m.EmpresaMultiCnpj(cnpj=CNPJ_A), AcessoFalso())
        self.assertEqual((resumo.qtd, resumo.pendencias), (0, 0))

    def test_gancho_que_devolve_apenas_inteiro_continua_funcionando(self):
        r = m.executar_escrituracao_multi_cnpj(
            self.planilha, AcessoFalso, processar_empresa=lambda e, a: 3, callback_log=self.log, competencia=ROTULO)
        self.assertEqual([x.status for x in r.empresas], [m.STATUS_CONCLUIDA] * 2)
        self.assertEqual(r.empresas[0].qtd_notas, 3)

    def test_resumo_da_gui_marca_simulada_como_ok_e_pendencias_como_nao_ok(self):
        r = m.ResultadoMultiCnpj(empresas=[
            m.ResultadoEmpresa(CNPJ_A, m.STATUS_SIMULADA), m.ResultadoEmpresa(CNPJ_B, m.STATUS_COM_PENDENCIAS)])
        oks = [e["ok"] for e in r.para_json()["empresas"]]
        self.assertEqual(oks, [True, False])


class TestSemCaminhoDeGravacaoNoModoSimulacao(_Base):
    """Prova final: rodando uma simulação completa, NENHUMA função de gravação é sequer chamada."""

    def test_nenhuma_funcao_de_gravacao_e_chamada_na_simulacao(self):
        gravar = mock.Mock(side_effect=AssertionError("gravação chamada!"))
        confirmar = mock.Mock(side_effect=AssertionError("confirmação chamada!"))
        with mock.patch.multiple(robo, preencher_dados_prestador=mock.DEFAULT, preencher_documento_servico=mock.DEFAULT,
                                 _clicar_com_espera=mock.DEFAULT), \
             mock.patch.object(robo, "_gravar_documento_com_confirmacoes", gravar), \
             mock.patch.object(robo, "_clicar_botao_modal_confirmacao", confirmar), \
             mock.patch.object(robo.time, "sleep"):
            docs = [_documento(numero_nf=str(i)) for i in range(20)]
            resultado = robo.processar_notas(_DriverFalso(), docs, COMPETENCIA, self.registro("Simulação"), simulacao=True)
        self.assertEqual(resultado.simuladas, 20)
        gravar.assert_not_called()
        confirmar.assert_not_called()


if __name__ == "__main__":
    unittest.main()
