import unittest

from iss_fortaleza_automacao import limpar_cnpj_para_digitacao, _normalizar_celula_cnpj


class TestLimparCnpjParaDigitacao(unittest.TestCase):
    def test_cnpj_numerico_legado_com_mascara(self):
        self.assertEqual(
            limpar_cnpj_para_digitacao("12.345.678/0001-95"), "12345678000195"
        )

    def test_cnpj_numerico_legado_sem_mascara(self):
        self.assertEqual(limpar_cnpj_para_digitacao("12345678000195"), "12345678000195")

    def test_cnpj_alfanumerico_com_mascara(self):
        self.assertEqual(
            limpar_cnpj_para_digitacao("12.ABC.345/01DE-35"), "12ABC34501DE35"
        )

    def test_cnpj_alfanumerico_minusculo_e_uppercased(self):
        self.assertEqual(
            limpar_cnpj_para_digitacao("12abc34501de35"), "12ABC34501DE35"
        )

    def test_none_e_vazio(self):
        self.assertEqual(limpar_cnpj_para_digitacao(None), "")
        self.assertEqual(limpar_cnpj_para_digitacao(""), "")

    def test_espacos_sao_removidos(self):
        self.assertEqual(limpar_cnpj_para_digitacao("  12 ABC 345 01DE 35 "), "12ABC34501DE35")


class TestNormalizarCelulaCnpjXlsx(unittest.TestCase):
    def test_celula_string_alfanumerica(self):
        self.assertEqual(_normalizar_celula_cnpj("12ABC34501DE35"), "12ABC34501DE35")

    def test_celula_numerica_int_com_zero_a_esquerda_perdido(self):
        # Excel armazenou como número; zero à esquerda foi perdido na conversão numérica.
        self.assertEqual(_normalizar_celula_cnpj(1234567800195), "01234567800195")

    def test_celula_numerica_float(self):
        self.assertEqual(_normalizar_celula_cnpj(12345678000195.0), "12345678000195")

    def test_celula_numerica_curta_e_preenchida_ate_14(self):
        # CNPJs legítimos podem começar com muitos zeros (ex.: 00.000.000/0001-91).
        self.assertEqual(_normalizar_celula_cnpj(191), "00000000000191")

    def test_celula_none(self):
        self.assertEqual(_normalizar_celula_cnpj(None), "")

    def test_valores_impossiveis_viram_vazio(self):
        # Precisam retornar "" para que validar_campos_obrigatorios() acuse o
        # CNPJ ausente, em vez de digitar um CNPJ inventado no portal.
        for valor in (0, 0.0, -5, False, True, "", "   "):
            with self.subTest(valor=valor):
                self.assertEqual(_normalizar_celula_cnpj(valor), "")


if __name__ == "__main__":
    unittest.main()
