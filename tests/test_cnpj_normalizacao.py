import unittest

from iss_fortaleza_automacao import (
    DocumentoPortalISS,
    _eh_cpf_prestador,
    _normalizar_celula_cnpj,
    limpar_cnpj_para_digitacao,
)


def _documento(cnpj_prestador="", tipo_cliente_prestador=""):
    """Cria um DocumentoPortalISS mínimo só com os campos relevantes para
    _eh_cpf_prestador, evitando repetir os demais campos obrigatórios do
    dataclass em cada teste."""
    return DocumentoPortalISS(
        arquivo_pdf="nota.xml",
        cnpj_prestador=cnpj_prestador,
        numero_nf="1",
        id_cnae_final="0101",
        data_emissao="01/01/2026",
        descricao_servico="Serviço teste",
        uf_local_prestacao="CE",
        cidade_local_prestacao="FORTALEZA",
        natureza_operacao="Tributação no Município",
        iss_retido="Não",
        valor_servico="100,00",
        tipo_cliente_prestador=tipo_cliente_prestador,
    )


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


class TestEhCpfPrestador(unittest.TestCase):
    def test_coluna_tipo_cliente_pessoa_fisica_manda(self):
        doc = _documento(cnpj_prestador="12345678000199", tipo_cliente_prestador="Pessoa Física")
        self.assertTrue(_eh_cpf_prestador(doc))

    def test_coluna_tipo_cliente_pessoa_juridica_manda(self):
        doc = _documento(cnpj_prestador="01336415460", tipo_cliente_prestador="Pessoa Jurídica")
        self.assertFalse(_eh_cpf_prestador(doc))

    def test_fallback_por_comprimento_quando_coluna_vazia_11_digitos_e_cpf(self):
        doc = _documento(cnpj_prestador="01336415460", tipo_cliente_prestador="")
        self.assertTrue(_eh_cpf_prestador(doc))

    def test_fallback_por_comprimento_quando_coluna_vazia_14_digitos_e_cnpj(self):
        doc = _documento(cnpj_prestador="12345678000199", tipo_cliente_prestador="")
        self.assertFalse(_eh_cpf_prestador(doc))

    def test_fallback_por_comprimento_cnpj_alfanumerico(self):
        doc = _documento(cnpj_prestador="12ABC34501DE35", tipo_cliente_prestador="")
        self.assertFalse(_eh_cpf_prestador(doc))


if __name__ == "__main__":
    unittest.main()
