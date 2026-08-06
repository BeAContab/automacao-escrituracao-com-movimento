import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from processamento_xml import (
    _detectar_layout_xml,
    _extrair_dados_xml_paraiba,
    _nome_municipio_ibge,
    extrair_dados_xml,
)

PASTA_AMOSTRAS = Path(__file__).resolve().parent.parent / "Layout Portal Da Paraiba"

# Prestador pessoa física (CPF), Campina Grande/PB, não cancelada.
AMOSTRA_CPF = PASTA_AMOSTRAS / "NFSe_2026000000001_2WRB-S3DR.xml"
# Prestador pessoa jurídica (CNPJ), Campina Grande/PB, não cancelada.
AMOSTRA_CNPJ = PASTA_AMOSTRAS / "NFSe_2026000000001_QRJF-LJEP.xml"
# Nota cancelada e substituída por outra (tem NfseCancelamento e NfseSubstituicao).
AMOSTRA_CANCELADA = PASTA_AMOSTRAS / "NFSe_2026000008959_EC5V-SNUD.xml"

_XML_NACIONAL_SINTETICO = """<?xml version="1.0" encoding="utf-8"?>
<NFSe>
  <infNFSe>
    <nNFSe>123</nNFSe>
    <dhEmi>2026-01-15T10:00:00</dhEmi>
  </infNFSe>
</NFSe>
"""


@unittest.skipUnless(PASTA_AMOSTRAS.exists(), "Pasta de amostras 'Layout Portal Da Paraiba' não encontrada")
class TestDetectarLayoutXml(unittest.TestCase):
    def test_amostra_paraiba_detectada_como_paraiba(self):
        root = ET.parse(AMOSTRA_CPF).getroot()
        self.assertEqual(_detectar_layout_xml(root), "paraiba")

    def test_xml_nacional_sintetico_detectado_como_nacional(self):
        root = ET.fromstring(_XML_NACIONAL_SINTETICO)
        self.assertEqual(_detectar_layout_xml(root), "nacional")

    def test_xml_sem_tag_conhecida_e_desconhecido(self):
        root = ET.fromstring("<algumaCoisa><outraTag/></algumaCoisa>")
        self.assertEqual(_detectar_layout_xml(root), "desconhecido")


@unittest.skipUnless(PASTA_AMOSTRAS.exists(), "Pasta de amostras 'Layout Portal Da Paraiba' não encontrada")
class TestExtrairDadosXmlParaibaPrestadorCpf(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root = ET.parse(AMOSTRA_CPF).getroot()
        cls.dados = _extrair_dados_xml_paraiba(root, AMOSTRA_CPF)

    def test_extracao_nao_e_none(self):
        self.assertIsNotNone(self.dados)

    def test_prestador_cpf(self):
        self.assertEqual(self.dados["cnpj_prestador"], "01336415460")
        self.assertEqual(self.dados["tipo_cliente_prestador"], "Pessoa Física")

    def test_dados_basicos(self):
        self.assertEqual(self.dados["numero_nf"], "2026000000001")
        self.assertEqual(self.dados["data_emissao"], "03/07/2026")
        self.assertEqual(self.dados["nome_prestador"], "Milton Fernando Campos de Melo")

    def test_endereco_prestador_resolvido_por_nome_de_cidade(self):
        self.assertEqual(self.dados["cidade_prestador"], "Campina Grande")
        self.assertEqual(self.dados["uf_prestador"], "PB")
        self.assertEqual(self.dados["cep_prestador"], "58446000")
        self.assertEqual(self.dados["logradouro_prestador"], "SITIO VARZEA DO ARROZ")
        self.assertEqual(self.dados["bairro_prestador"], "Não informado")

    def test_prefeitura(self):
        self.assertEqual(self.dados["prefeitura"], "PREFEITURA MUNICIPAL DE CAMPINA GRANDE")

    def test_valores(self):
        self.assertEqual(self.dados["valor_servico"], "4.750,00")
        self.assertEqual(self.dados["aliquota"], "5,00")

    def test_iss_retido_polaridade_abrasf(self):
        # IssRetido=2 no padrão ABRASF significa NÃO retido (polaridade
        # invertida em relação ao layout Nacional).
        self.assertEqual(self.dados["iss_retido"], "NÃO")

    def test_natureza_operacao_fora_do_municipio(self):
        self.assertEqual(self.dados["natureza_operacao"], "Tributação Fora do Município")

    def test_regime_tributario_outros_quando_nao_e_codigo_5(self):
        # Esta amostra não tem RegimeEspecialTributacao nenhum.
        self.assertEqual(self.dados["regime_tributario"], "OUTROS")

    def test_descricao_servico_sem_artefato_de_escape(self):
        self.assertNotIn("\\s\\n", self.dados["descricao_servico"])
        self.assertIn("SERVIÇOS DE PRODUÇÃO MUSICAL", self.dados["descricao_servico"])


@unittest.skipUnless(PASTA_AMOSTRAS.exists(), "Pasta de amostras 'Layout Portal Da Paraiba' não encontrada")
class TestExtrairDadosXmlParaibaPrestadorCnpj(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root = ET.parse(AMOSTRA_CNPJ).getroot()
        cls.dados = _extrair_dados_xml_paraiba(root, AMOSTRA_CNPJ)

    def test_prestador_cnpj(self):
        self.assertEqual(self.dados["cnpj_prestador"], "23239797000180")
        self.assertEqual(self.dados["tipo_cliente_prestador"], "Pessoa Jurídica")

    def test_id_cnae_usa_codigo_cnae_quando_presente(self):
        # Esta amostra tem <CodigoCnae>5620102</CodigoCnae>, que deve ter
        # prioridade sobre ItemListaServico (1711).
        self.assertEqual(self.dados["id_cnae"], "5620102")

    def test_regime_tributario_outros_para_codigo_6(self):
        # RegimeEspecialTributacao=6 (ME/EPP Simples Nacional) NÃO é MEI —
        # só o código 5 é tratado como MEI.
        self.assertEqual(self.dados["regime_tributario"], "OUTROS")


@unittest.skipUnless(PASTA_AMOSTRAS.exists(), "Pasta de amostras 'Layout Portal Da Paraiba' não encontrada")
class TestExtrairDadosXmlNotaCancelada(unittest.TestCase):
    def test_nota_cancelada_e_ignorada_com_motivo_especifico(self):
        dados, motivo = extrair_dados_xml(AMOSTRA_CANCELADA)
        self.assertIsNone(dados)
        self.assertEqual(motivo, "nota cancelada")


class TestRegimeTributarioMei(unittest.TestCase):
    """RegimeEspecialTributacao == '5' é o único código ABRASF oficial para MEI;
    nenhuma amostra real tem esse código, então este caso é sintético."""

    _XML_MEI = """<?xml version="1.0" encoding="utf-8"?>
<CompNfse xmlns="http://www.abrasf.org.br/nfse.xsd">
  <Nfse versao="2.02">
    <InfNfse Id="x">
      <Numero>1</Numero>
      <DataEmissao>2026-01-01T00:00:00</DataEmissao>
      <ValoresNfse><Aliquota>5</Aliquota></ValoresNfse>
      <PrestadorServico>
        <IdentificacaoPrestador><CpfCnpj><Cnpj>12345678000199</Cnpj></CpfCnpj></IdentificacaoPrestador>
        <RazaoSocial>Empresa Teste</RazaoSocial>
        <Endereco><Endereco>Rua X</Endereco><CodigoMunicipio>2504009</CodigoMunicipio><Uf>PB</Uf><Cep>58000000</Cep></Endereco>
      </PrestadorServico>
      <OrgaoGerador><CodigoMunicipio>2504009</CodigoMunicipio><Uf>PB</Uf></OrgaoGerador>
      <DeclaracaoPrestacaoServico>
        <InfDeclaracaoPrestacaoServico>
          <Servico>
            <Valores><ValorServicos>100</ValorServicos></Valores>
            <IssRetido>2</IssRetido>
            <ItemListaServico>0101</ItemListaServico>
          </Servico>
          <RegimeEspecialTributacao>5</RegimeEspecialTributacao>
        </InfDeclaracaoPrestacaoServico>
      </DeclaracaoPrestacaoServico>
    </InfNfse>
  </Nfse>
</CompNfse>
"""

    def test_codigo_5_e_mei(self):
        with tempfile.TemporaryDirectory() as tmp:
            caminho = Path(tmp, "nota_mei.xml")
            caminho.write_text(self._XML_MEI, encoding="utf-8")
            root = ET.parse(caminho).getroot()
            dados = _extrair_dados_xml_paraiba(root, caminho)
            self.assertEqual(dados["regime_tributario"], "MEI")


class TestNomeMunicipioIbge(unittest.TestCase):
    def test_codigo_conhecido_resolve_para_nome(self):
        self.assertEqual(_nome_municipio_ibge("2504009"), "Campina Grande")

    def test_codigo_desconhecido_retorna_o_proprio_codigo(self):
        self.assertEqual(_nome_municipio_ibge("9999999"), "9999999")

    def test_codigo_vazio_retorna_vazio(self):
        self.assertEqual(_nome_municipio_ibge(""), "")


if __name__ == "__main__":
    unittest.main()
