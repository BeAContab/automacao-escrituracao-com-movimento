import tempfile
import unittest
from pathlib import Path
from unittest import mock

import processamento_xml
from processamento_xml import processar_pasta_xmls_auto


class TestProcessarPastaXmlsAuto(unittest.TestCase):
    def test_pasta_sem_xml_e_sem_subpastas_levanta_erro(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(FileNotFoundError):
                processar_pasta_xmls_auto(tmp)

    def test_subpastas_todas_vazias_levanta_erro_citando_os_nomes(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "ClienteA").mkdir()
            Path(tmp, "ClienteB").mkdir()
            with self.assertRaises(FileNotFoundError) as ctx:
                processar_pasta_xmls_auto(tmp)
            self.assertIn("ClienteA", str(ctx.exception))
            self.assertIn("ClienteB", str(ctx.exception))

    def test_xml_direto_delega_uma_vez_escopado_apenas_aos_arquivos_diretos(self):
        with tempfile.TemporaryDirectory() as tmp:
            xml_path = Path(tmp, "nota1.xml")
            xml_path.write_text("<x/>", encoding="utf-8")
            with mock.patch.object(
                processamento_xml, "processar_pasta_xmls",
                return_value=f"{tmp}/notas_xml_processadas.xlsx",
            ) as mocked:
                resultado = processar_pasta_xmls_auto(tmp)

            mocked.assert_called_once()
            self.assertEqual(mocked.call_args.kwargs["pasta_origem"], tmp)
            # Não pode delegar sem arquivos_xml_forcados: sem isso, processar_pasta_xmls
            # faria seu próprio rglob recursivo e engoliria o conteúdo de subpastas.
            self.assertEqual(mocked.call_args.kwargs["arquivos_xml_forcados"], [xml_path])
            self.assertEqual(resultado, [f"{tmp}/notas_xml_processadas.xlsx"])

    def test_subpastas_com_xml_geram_uma_chamada_cada_e_pulam_vazias(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "ClienteA").mkdir()
            Path(tmp, "ClienteA", "nota1.xml").write_text("<x/>", encoding="utf-8")
            Path(tmp, "ClienteB").mkdir()
            Path(tmp, "ClienteB", "nota2.xml").write_text("<x/>", encoding="utf-8")
            Path(tmp, "ClienteC_vazia").mkdir()

            def fake(pasta_origem, **kwargs):
                return f"{pasta_origem}/notas_xml_processadas.xlsx"

            with mock.patch.object(processamento_xml, "processar_pasta_xmls", side_effect=fake) as mocked:
                resultado = processar_pasta_xmls_auto(tmp)

            self.assertEqual(mocked.call_count, 2)
            self.assertEqual(len(resultado), 2)

    def test_pasta_pai_com_xml_direto_e_subpastas_gera_uma_planilha_para_cada(self):
        """Regressão: reproduz o cenário relatado pelo usuário — pasta selecionada com
        XML direto NELA MESMA e também com subpastas com XML. Antes da correção, isso
        caía no modo "1 planilha" (que usa rglob e engolia as subpastas junto)."""
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "solto.xml").write_text("<x/>", encoding="utf-8")
            Path(tmp, "ClienteA").mkdir()
            Path(tmp, "ClienteA", "nota1.xml").write_text("<x/>", encoding="utf-8")
            Path(tmp, "ClienteB").mkdir()
            Path(tmp, "ClienteB", "nota2.xml").write_text("<x/>", encoding="utf-8")

            def fake(pasta_origem, **kwargs):
                return f"{pasta_origem}/notas_xml_processadas.xlsx"

            with mock.patch.object(processamento_xml, "processar_pasta_xmls", side_effect=fake) as mocked:
                resultado = processar_pasta_xmls_auto(tmp)

            # 1 chamada para os XMLs soltos na raiz + 1 por subpasta = 3 no total
            self.assertEqual(mocked.call_count, 3)
            self.assertEqual(len(resultado), 3)
            primeira_chamada = mocked.call_args_list[0]
            self.assertEqual(primeira_chamada.kwargs["pasta_origem"], tmp)
            self.assertEqual(
                [p.name for p in primeira_chamada.kwargs["arquivos_xml_forcados"]],
                ["solto.xml"],
            )


if __name__ == "__main__":
    unittest.main()
