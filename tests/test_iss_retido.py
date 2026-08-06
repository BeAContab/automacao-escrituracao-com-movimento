import unittest

from iss_fortaleza_automacao import _aplicar_override_iss_retido_por_id_cnae, _interpretar_iss_retido
from processamento_xml import _aplicar_regra_iss_retido_por_id_cnae


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


if __name__ == "__main__":
    unittest.main()
