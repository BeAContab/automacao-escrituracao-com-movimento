import unittest

from iss_fortaleza_automacao import (
    DocumentoPortalISS,
    _aplicar_override_iss_retido_por_id_cnae,
    _aplicar_override_iss_retido_por_tipo_tributacao,
    _eh_prestador_fortaleza_ce,
    _interpretar_iss_retido,
)
from processamento_xml import _aplicar_regra_iss_retido_por_id_cnae, _eh_nome_prestador_padrao_mei


def _documento(uf_prestador="", cidade_prestador=""):
    """Cria um DocumentoPortalISS mínimo só com os campos relevantes para
    _eh_prestador_fortaleza_ce."""
    return DocumentoPortalISS(
        arquivo_pdf="nota.xml",
        cnpj_prestador="12345678000199",
        numero_nf="1",
        id_cnae_final="0101",
        data_emissao="01/01/2026",
        descricao_servico="Serviço teste",
        uf_local_prestacao="CE",
        cidade_local_prestacao="FORTALEZA",
        natureza_operacao="Tributação no Município",
        iss_retido="Não",
        valor_servico="100,00",
        uf_prestador=uf_prestador,
        cidade_prestador=cidade_prestador,
    )


class TestInterpretarIssRetido(unittest.TestCase):
    def test_sim_variacoes(self):
        for valor in ("SIM", "Sim", "sim", " Sim "):
            with self.subTest(valor=valor):
                self.assertEqual(_interpretar_iss_retido(valor), (True, True))

    def test_nao_variacoes(self):
        for valor in ("NÃO", "Não", "nao", "N", "n"):
            with self.subTest(valor=valor):
                self.assertEqual(_interpretar_iss_retido(valor), (False, True))

    def test_valores_nao_reconhecidos_tratados_como_nao_retido(self):
        for valor in ("", "talvez", "S", "x", None):
            with self.subTest(valor=valor):
                deve_marcar, reconhecido = _interpretar_iss_retido(valor)
                self.assertFalse(deve_marcar)
                self.assertFalse(reconhecido)


class TestOverrideIssRetidoPorIdCnaeAutomacao(unittest.TestCase):
    """Override aplicado no momento de decidir o checkbox durante a automação
    (iss_fortaleza_automacao.py) — tem prioridade sobre o valor de ISS_RETIDO
    já lido da planilha."""

    def test_1207_forca_marcar_mesmo_partindo_de_nao(self):
        self.assertTrue(_aplicar_override_iss_retido_por_id_cnae(False, "1207"))

    def test_1213_forca_desmarcar_mesmo_partindo_de_sim(self):
        self.assertFalse(_aplicar_override_iss_retido_por_id_cnae(True, "1213"))

    def test_outro_id_cnae_nao_altera_o_valor(self):
        self.assertTrue(_aplicar_override_iss_retido_por_id_cnae(True, "0107"))
        self.assertFalse(_aplicar_override_iss_retido_por_id_cnae(False, "0107"))

    def test_id_cnae_vazio_nao_altera_o_valor(self):
        self.assertTrue(_aplicar_override_iss_retido_por_id_cnae(True, ""))
        self.assertFalse(_aplicar_override_iss_retido_por_id_cnae(False, None))

    def test_id_cnae_com_mascara_e_normalizado(self):
        self.assertTrue(_aplicar_override_iss_retido_por_id_cnae(False, "12.07"))
        self.assertFalse(_aplicar_override_iss_retido_por_id_cnae(True, "12-13"))


class TestRegraIssRetidoPorIdCnaeProcessamento(unittest.TestCase):
    """Mesma regra, aplicada na geração da planilha (processamento_xml.py)."""

    def test_1207_forca_sim(self):
        self.assertEqual(_aplicar_regra_iss_retido_por_id_cnae("NÃO", "1207"), "SIM")

    def test_1213_forca_nao(self):
        self.assertEqual(_aplicar_regra_iss_retido_por_id_cnae("SIM", "1213"), "NÃO")

    def test_outro_id_cnae_preserva_valor_recebido(self):
        self.assertEqual(_aplicar_regra_iss_retido_por_id_cnae("SIM", "0107"), "SIM")
        self.assertEqual(_aplicar_regra_iss_retido_por_id_cnae("NÃO", "0107"), "NÃO")


class TestOverrideIssRetidoPorTipoTributacao(unittest.TestCase):
    """Override de Simples Nacional MEI — prioridade máxima, aplicado depois
    do override de ID_CNAE."""

    def test_simples_nacional_mei_forca_desmarcar(self):
        self.assertFalse(_aplicar_override_iss_retido_por_tipo_tributacao(True, "Simples Nacional MEI"))

    def test_normal_nao_altera_o_valor(self):
        self.assertTrue(_aplicar_override_iss_retido_por_tipo_tributacao(True, "Normal"))
        self.assertFalse(_aplicar_override_iss_retido_por_tipo_tributacao(False, "Normal"))

    def test_simples_nacional_me_epp_nao_altera_o_valor(self):
        self.assertTrue(_aplicar_override_iss_retido_por_tipo_tributacao(True, "Simples Nacional ME-EPP"))

    def test_prioridade_sobre_id_cnae_1207(self):
        # ID_CNAE 1207 forçaria "marcado", mas MEI vence por ser aplicado depois.
        deve_marcar = _aplicar_override_iss_retido_por_id_cnae(False, "1207")
        deve_marcar = _aplicar_override_iss_retido_por_tipo_tributacao(deve_marcar, "Simples Nacional MEI")
        self.assertFalse(deve_marcar)


class TestEhPrestadorFortalezaCe(unittest.TestCase):
    def test_fortaleza_ce_e_reconhecido(self):
        self.assertTrue(_eh_prestador_fortaleza_ce(_documento(uf_prestador="CE", cidade_prestador="FORTALEZA")))

    def test_outra_cidade_no_ce_nao_e_fortaleza(self):
        self.assertFalse(_eh_prestador_fortaleza_ce(_documento(uf_prestador="CE", cidade_prestador="SOBRAL")))

    def test_fortaleza_em_outra_uf_nao_conta(self):
        self.assertFalse(_eh_prestador_fortaleza_ce(_documento(uf_prestador="PB", cidade_prestador="CAMPINA GRANDE")))


class TestEhNomePrestadorPadraoMei(unittest.TestCase):
    def test_padrao_mei_reconhecido(self):
        self.assertTrue(_eh_nome_prestador_padrao_mei("40.386.363 JOSE CARLOS GOMES DA SILVA"))
        self.assertTrue(_eh_nome_prestador_padrao_mei("36.592.831 JOSE ANDREWY DA SILVA MELO"))

    def test_nome_comum_nao_bate(self):
        self.assertFalse(_eh_nome_prestador_padrao_mei("JOSE CARLOS GOMES DA SILVA"))

    def test_razao_social_ltda_nao_bate(self):
        self.assertFalse(_eh_nome_prestador_padrao_mei("X1 LIKE PRODUCOES E EVENTOS LTDA"))

    def test_vazio_nao_bate(self):
        self.assertFalse(_eh_nome_prestador_padrao_mei(""))
        self.assertFalse(_eh_nome_prestador_padrao_mei(None))


if __name__ == "__main__":
    unittest.main()
