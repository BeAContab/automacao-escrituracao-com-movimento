"""Testes de `iss_athenas_core.py` (migração do standalone "ISS Fortaleza → Athenas").

Cobre os casos unitários sugeridos no `MIGRACAO.md` §11, mais o teste de regressão
comparando com o arquivo de referência validado pela equipe fiscal, se disponível
localmente em `importar/`.
"""

import shutil
import tempfile
import unittest
from collections import Counter
from pathlib import Path

import pandas as pd

import iss_athenas_core as m

CAMINHO_ISS_REFERENCIA = Path("importar/NF-15350-082026.xlsx")
CAMINHO_ATHENAS_REFERENCIA = Path("importar/Athenas_Tomados_NF-15350-082026.xlsx")

# Colunas do arquivo ISS Fortaleza na ordem posicional que `processar_aba` espera
# (43 colunas — ver MIGRACAO.md §7). Só as usadas nos testes têm valor não-trivial;
# as demais ficam com um valor neutro qualquer.
COLUNAS_ISS = 43


def _linha_iss(**valores) -> list:
    """Monta uma linha de 43 colunas no layout do ISS Fortaleza, com defaults neutros
    e sobrescrevendo pelos índices/nomes passados em `valores` (aceita COL_* por nome)."""
    linha = [None] * COLUNAS_ISS
    linha[1] = 100  # Número
    linha[3] = "09/2026"  # Competência
    linha[4] = "15/09/2026"  # Data
    linha[13] = "10.10"  # Item da Lista
    linha[19] = 1000.0  # Valor dos Serviços
    linha[23] = 0.0  # Retenções Federais
    linha[25] = 0.0  # PIS
    linha[26] = 0.0  # COFINS
    linha[27] = 0.0  # IRRF
    linha[28] = 0.0  # CSLL
    linha[29] = 0.0  # INSS
    linha[33] = "Não"  # ISS Retido
    linha[34] = 0.0  # Valor do ISS
    linha[36] = "NORMAL"  # Status Doc.
    linha[38] = "12345678000195"  # CPF/CNPJ
    linha[39] = "Prestador Teste"  # Razão Social
    for nome_ou_indice, valor in valores.items():
        indice = getattr(m, nome_ou_indice) if isinstance(nome_ou_indice, str) and nome_ou_indice.startswith("COL_") else nome_ou_indice
        linha[indice] = valor
    return linha


class _DfFalso:
    """Substitui um DataFrame só para `processar_aba`, que usa apenas `.values`."""

    def __init__(self, linhas: list[list]):
        self.values = linhas


class TestItemLista(unittest.TestCase):
    def test_float_recupera_o_zero_perdido(self):
        self.assertEqual(m.item_lista_int(17.1), 1710)  # 17.10 perdeu o zero final no Excel

    def test_string_com_ponto(self):
        self.assertEqual(m.item_lista_int("31.01"), 3101)

    def test_string_curta(self):
        self.assertEqual(m.item_lista_int("1.05"), 105)

    def test_none_vira_string_vazia(self):
        self.assertEqual(m.item_lista_int(None), "")


class TestCnpjCpf(unittest.TestCase):
    def test_cnpj_float_com_ponto_zero(self):
        self.assertEqual(m.formatar_cnpj_cpf(50336159000150.0), "50336159000150")

    def test_cpf_11_digitos(self):
        self.assertEqual(m.formatar_cnpj_cpf(12345678901), "12345678901")

    def test_cpf_repoe_zero_a_esquerda(self):
        self.assertEqual(m.formatar_cnpj_cpf(1234567890), "01234567890")

    def test_vazio(self):
        self.assertEqual(m.formatar_cnpj_cpf(None), "")


class TestAjustarData(unittest.TestCase):
    def test_data_fora_da_competencia_vira_ultimo_dia(self):
        self.assertEqual(m.ajustar_data("31/03/2026", "04/2026"), "30/04/2026")

    def test_data_dentro_da_competencia_mantem(self):
        self.assertEqual(m.ajustar_data("15/04/2026", "04/2026"), "15/04/2026")


class TestRetencoesFederais(unittest.TestCase):
    """Espelha a árvore de decisão do §6.3 do MIGRACAO.md."""

    def _processar(self, **kwargs):
        linha = _linha_iss(**kwargs)
        return m.processar_aba(_DfFalso([linha]), "tomados")[0]

    def test_pis_cofins_zero_explicito_com_csll_lump(self):
        # CSLL = 4,65% do valor: PIS/COFINS/CSLL vêm agrupados num único valor
        r = self._processar(COL_RET_FED=1000.0, COL_PIS=0.0, COL_COFINS=0.0, COL_CSLL=46.5)
        pis, cofins, irrf, csll = r["valores"][7:11]
        self.assertAlmostEqual(pis, round(1000 * m.TAXA_PIS, 5))
        self.assertAlmostEqual(cofins, round(1000 * m.TAXA_COFINS, 5))
        self.assertAlmostEqual(csll, round(1000 * m.TAXA_CSLL, 5))
        self.assertEqual(r["valores"][16], "")  # sem observação de divergência

    def test_lump_sobre_base_reduzida(self):
        # Exemplo real do MIGRACAO.md §6.3, caso 2
        r = self._processar(COL_VALOR=56696.51, COL_RET_FED=1.0, COL_CSLL=205.34, COL_PIS=0.0, COL_COFINS=0.0)
        pis, cofins, _irrf, csll = r["valores"][7:11]
        self.assertAlmostEqual(pis, 28.70, places=1)
        self.assertAlmostEqual(cofins, 132.48, places=1)
        self.assertAlmostEqual(csll, 44.16, places=1)

    def test_tres_impostos_individualizados_sem_divergencia(self):
        r = self._processar(
            COL_VALOR=4200.0, COL_RET_FED=1.0, COL_PIS=27.30, COL_COFINS=126.01, COL_CSLL=42.01,
        )
        pis, cofins, _irrf, csll = r["valores"][7:11]
        self.assertEqual((pis, cofins, csll), (27.30, 126.01, 42.01))
        self.assertEqual(r["valores"][16], "")  # dentro da tolerância: sem observação

    def test_tres_impostos_com_divergencia_registra_observacao(self):
        r = self._processar(
            COL_VALOR=4200.0, COL_RET_FED=1.0, COL_PIS=100.0, COL_COFINS=126.01, COL_CSLL=42.01,
        )
        self.assertIn("Percentual divergente", r["valores"][16])
        self.assertEqual(r["valores"][7], 100.0)  # valor recebido é mantido mesmo divergente

    def test_so_csll_um_por_cento_fica_individual(self):
        r = self._processar(COL_VALOR=1000.0, COL_RET_FED=1.0, COL_CSLL=10.0, COL_PIS=0.0, COL_COFINS=0.0)
        pis, cofins, _irrf, csll = r["valores"][7:11]
        self.assertIsNone(pis)
        self.assertIsNone(cofins)
        self.assertEqual(csll, 10.0)

    def test_so_pis_e_cofins_sem_csll_ficam_vazios_e_marcam_amarelo(self):
        r = self._processar(COL_VALOR=1000.0, COL_RET_FED=1.0, COL_PIS=10.0, COL_COFINS=30.0, COL_CSLL=0.0)
        pis, cofins, _irrf, csll = r["valores"][7:11]
        self.assertIsNone(pis)
        self.assertIsNone(cofins)
        self.assertIsNone(csll)
        self.assertTrue(r["amarelo"])
        self.assertIn("nao sao retidos", r["valores"][16])

    def test_irrf_nunca_e_calculado_so_da_coluna(self):
        r = self._processar(COL_RET_FED=1.0, COL_IRRF=15.0, COL_CSLL=0.0, COL_PIS=0.0, COL_COFINS=0.0)
        self.assertEqual(r["valores"][9], 15.0)

    def test_sem_retencao_federal_impostos_ficam_vazios(self):
        r = self._processar(COL_RET_FED=0.0)
        self.assertEqual(r["valores"][7:11], [None, None, None, None])


class TestCodClienteEStatus(unittest.TestCase):
    def test_cod_cliente_1_quando_tomador_sem_cnpj_prestados(self):
        linha = _linha_iss(COL_CNPJ=None)
        r = m.processar_aba(_DfFalso([linha]), "prestados", regime="normal")[0]
        # Prestados insere Valor do ISS Próprio e Cod Cliente após "Valor do ISS Retido" (índice 12)
        self.assertEqual(r["valores"][14], 1)

    def test_cod_cliente_vazio_quando_ha_cnpj(self):
        linha = _linha_iss(COL_CNPJ="12345678000195")
        r = m.processar_aba(_DfFalso([linha]), "prestados", regime="normal")[0]
        self.assertIsNone(r["valores"][14])

    def test_tomados_nao_tem_coluna_cod_cliente(self):
        linha = _linha_iss()
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertEqual(len(r["valores"]), len(m.HDR_TOMADOS))

    def test_status_diferente_de_normal_marca_cancelada(self):
        linha = _linha_iss(COL_STATUS="CANCELADA")
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertTrue(r["cancelada"])
        self.assertEqual(r["valores"][15], "CANCELADA")

    def test_status_vazio_vira_normal(self):
        linha = _linha_iss(COL_STATUS="")
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertFalse(r["cancelada"])


class TestCfop(unittest.TestCase):
    def test_tomados_com_retencao(self):
        linha = _linha_iss(COL_RET_FED=1.0)
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertEqual(r["valores"][3], 8013)

    def test_tomados_sem_retencao(self):
        linha = _linha_iss()
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertEqual(r["valores"][3], 8012)

    def test_prestados_normal_com_retencao_via_iss_retido(self):
        linha = _linha_iss(COL_ISS_RETIDO="Sim", COL_VALOR_ISS=50.0)
        r = m.processar_aba(_DfFalso([linha]), "prestados", regime="normal")[0]
        self.assertEqual(r["valores"][3], 8011)

    def test_prestados_simples_sem_retencao(self):
        linha = _linha_iss()
        r = m.processar_aba(_DfFalso([linha]), "prestados", regime="simples")[0]
        self.assertEqual(r["valores"][3], 8001)

    def test_retencao_via_inss(self):
        linha = _linha_iss(COL_INSS=20.0)
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertEqual(r["valores"][3], 8013)


class TestIss(unittest.TestCase):
    def test_iss_retido_preenchido_quando_sim(self):
        linha = _linha_iss(COL_ISS_RETIDO="Sim", COL_VALOR_ISS=42.0)
        r = m.processar_aba(_DfFalso([linha]), "tomados")[0]
        self.assertEqual(r["valores"][12], 42.0)

    def test_iss_proprio_preenchido_so_em_prestados_quando_nao_retido(self):
        linha = _linha_iss(COL_ISS_RETIDO="Não", COL_VALOR_ISS=42.0)
        r = m.processar_aba(_DfFalso([linha]), "prestados", regime="normal")[0]
        self.assertEqual(r["valores"][13], 42.0)  # Valor do ISS Próprio
        self.assertIsNone(r["valores"][12])  # Valor do ISS Retido


class TestProcessarArquivoESalvarLogErros(unittest.TestCase):
    """`processar_arquivo`/`salvar_log_erros` com arquivos reais (openpyxl), sem rede."""

    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="athenas_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)

    def _criar_planilha_iss(self, nome, abas):
        caminho = self.pasta / nome
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            for nome_aba, linhas in abas.items():
                df = pd.DataFrame(linhas, columns=[f"col{i}" for i in range(COLUNAS_ISS)])
                df.to_excel(writer, sheet_name=nome_aba, header=True, index=False)
        return caminho

    def test_gera_os_dois_arquivos_quando_ha_as_duas_abas(self):
        linha = _linha_iss()
        caminho = self._criar_planilha_iss(
            "empresa1.xlsx", {"Serviços Tomados": [linha], "Serviços Prestados": [linha]}
        )
        logs = []
        resultado = m.processar_arquivo(caminho, "normal", self.pasta, log=logs.append)
        self.assertEqual(resultado, {"tomados": 1, "prestados": 1})
        self.assertTrue((self.pasta / "Athenas_Tomados_empresa1.xlsx").exists())
        self.assertTrue((self.pasta / "Athenas_Prestados_empresa1.xlsx").exists())
        self.assertTrue(any("Tomados" in l for l in logs))

    def test_gera_so_o_arquivo_da_aba_presente(self):
        linha = _linha_iss()
        caminho = self._criar_planilha_iss("empresa2.xlsx", {"Serviços Tomados": [linha]})
        resultado = m.processar_arquivo(caminho, "normal", self.pasta, log=lambda _m: None)
        self.assertEqual(resultado, {"tomados": 1, "prestados": 0})
        self.assertFalse((self.pasta / "Athenas_Prestados_empresa2.xlsx").exists())

    def test_sem_nenhuma_aba_conhecida_levanta_erro(self):
        caminho = self._criar_planilha_iss("empresa3.xlsx", {"Outra Aba": [_linha_iss()]})
        with self.assertRaises(ValueError):
            m.processar_arquivo(caminho, "normal", self.pasta, log=lambda _m: None)

    def test_salvar_log_erros_grava_uma_linha_por_arquivo(self):
        erros = [
            {"Arquivo": "a.xlsx", "Caminho": "C:/a.xlsx", "Regime": "normal", "Erro": "falhou", "Detalhes": "traceback"},
        ]
        caminho = m.salvar_log_erros(erros, self.pasta)
        self.assertTrue(caminho.exists())
        wb_dados = pd.read_excel(caminho, sheet_name="Erros")
        self.assertEqual(len(wb_dados), 1)
        self.assertEqual(wb_dados.iloc[0]["Arquivo"], "a.xlsx")


@unittest.skipUnless(
    CAMINHO_ISS_REFERENCIA.exists() and CAMINHO_ATHENAS_REFERENCIA.exists(),
    "arquivos de referência não encontrados em importar/ — copie-os para rodar esta regressão",
)
class TestRegressaoContraReferenciaValidada(unittest.TestCase):
    """Compara com o arquivo de referência validado pela equipe fiscal (MIGRACAO.md §11).

    Divergência CONHECIDA e aceita pelo usuário: a referência exclui 4 notas com a coluna
    "Status Aceite" = "Recusada" (89599097, 19, 10520, 21); o código (fiel ao original) não
    filtra por essa coluna — só por "Status Doc." (cancelada). Este teste, portanto, NÃO
    compara a contagem total de linhas; compara os valores campo a campo só das notas que
    aparecem nos dois conjuntos (interseção por Número), que é o que importa para garantir
    que a lógica de cálculo continua correta.
    """

    CAMPOS = [
        "Valor dos Serviços", "PIS", "COFINS", "IRRF", "CSLL", "INSS",
        "Valor do ISS Retido", "CFOP", "Item da Lista", "Valor Líquido",
    ]

    def test_valores_batem_para_as_notas_presentes_nos_dois_conjuntos(self):
        # Chave (Número, CNPJ do prestador): "Número" sozinho NÃO é único — prestadores
        # diferentes reutilizam a mesma numeração de nota (confirmado no arquivo real:
        # números como 8, 13, 17, 19, 70... aparecem 2x, de empresas diferentes).
        df = pd.read_excel(CAMINHO_ISS_REFERENCIA, sheet_name="Serviços Tomados", header=0)
        linhas = m.processar_aba(df, "tomados", col_valor_liq=m.encontrar_col_valor_liq(df))
        novo = pd.DataFrame([l["valores"] for l in linhas if not l["cancelada"]], columns=m.HDR_TOMADOS)

        ref = pd.read_excel(CAMINHO_ATHENAS_REFERENCIA, sheet_name="Importação (2)")
        ref = ref.drop(columns=[c for c in ref.columns if str(c).startswith("Unnamed")])
        ref = ref[ref["Número"].notna()]  # remove as linhas de totalização do rodapé

        def chave_novo(row):
            return (int(row["Número"]), str(row["CPF/CNPJ Prestador"]).strip())

        def chave_ref(row):
            # a referência guarda o CNPJ cru (ex.: "8324965000141.0"); normaliza do mesmo
            # jeito que o código faz para poder casar com a chave de "novo".
            return (int(row["Número"]), m.formatar_cnpj_cpf(row["CPF/CNPJ Prestador"]))

        novo_por_chave = {chave_novo(v): v for _, v in novo.iterrows()}
        ref_por_chave = {chave_ref(v): v for _, v in ref.iterrows()}
        chaves_em_comum = set(novo_por_chave) & set(ref_por_chave)

        # A referência inteira precisa estar contida em "novo" (a única diferença esperada
        # são as 4 notas "Recusada" que só existem em "novo" — ver teste seguinte).
        self.assertEqual(set(ref_por_chave) - chaves_em_comum, set())
        self.assertEqual(len(chaves_em_comum), len(ref_por_chave))

        divergencias = []
        for chave in chaves_em_comum:
            linha_novo, linha_ref = novo_por_chave[chave], ref_por_chave[chave]
            for campo in self.CAMPOS:
                a = m.nz(linha_novo[campo]) if pd.notna(linha_novo[campo]) else 0.0
                b = m.nz(linha_ref[campo]) if pd.notna(linha_ref[campo]) else 0.0
                if abs(a - b) >= 0.02:
                    divergencias.append((chave, campo, a, b))

        self.assertEqual(divergencias, [])

    def test_diferenca_de_contagem_e_exatamente_as_4_notas_recusadas_conhecidas(self):
        df = pd.read_excel(CAMINHO_ISS_REFERENCIA, sheet_name="Serviços Tomados", header=0)
        linhas = m.processar_aba(df, "tomados", col_valor_liq=m.encontrar_col_valor_liq(df))
        numeros_novo = Counter(int(l["valores"][0]) for l in linhas if not l["cancelada"])

        ref = pd.read_excel(CAMINHO_ATHENAS_REFERENCIA, sheet_name="Importação (2)")
        ref = ref[ref["Número"].notna()]
        numeros_ref = Counter(int(v) for v in ref["Número"])

        sobrando = numeros_novo - numeros_ref
        faltando = numeros_ref - numeros_novo
        self.assertEqual(faltando, Counter())  # a referência é subconjunto do que geramos
        self.assertEqual(sum(sobrando.values()), 4)  # exatamente as 4 notas "Recusada" conhecidas


if __name__ == "__main__":
    unittest.main()
