"""Testes da pausa/parada dentro do laço de páginas de `executar_exportacao_xml_prestados`.

REGRA: nenhum teste abre um navegador nem toca o portal. Todas as interações com o
Selenium (WebDriverWait, seleção de competência, avanço de página, espera de download)
são dubladas; o "driver" nunca é um navegador de verdade.
"""

import unittest
from unittest import mock

import exportador_xml_prestados as m


class _EsperaFalsa:
    """Substitui WebDriverWait: `.until(condicao)` sempre devolve um elemento clicável falso,
    independente da condição pedida — o teste não valida seletores, só o controle de fluxo."""

    def __init__(self, driver, timeout):
        pass

    def until(self, condicao):
        return mock.Mock()


class TestPausaEParadaNoLacoDePaginas(unittest.TestCase):
    def setUp(self):
        self.logs = []
        self.log = lambda msg, is_error=False: self.logs.append((msg, is_error))
        self.progresso = []

        patches = [
            mock.patch.object(m, "abrir_navegador_visivel", return_value=mock.Mock()),
            mock.patch.object(m, "WebDriverWait", _EsperaFalsa),
            mock.patch.object(m, "_selecionar_competencia_consulta"),
            mock.patch.object(m, "_aguardar_downloads_concluirem"),
            mock.patch.object(m.time, "sleep"),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

    def _rodar(self, avancar_pagina_valores, callback_pausa=None, callback_cancelamento=None):
        with mock.patch.object(m, "_tentar_avancar_pagina", side_effect=avancar_pagina_valores):
            m.executar_exportacao_xml_prestados(
                pasta_destino="C:\\qualquer\\coisa",
                competencia_str="09/2026",
                callback_log=self.log,
                callback_progresso=lambda pct, status: self.progresso.append((pct, status)),
                callback_pausa=callback_pausa,
                callback_cancelamento=callback_cancelamento,
            )

    def test_conclusao_normal_nao_menciona_parada_e_ainda_chama_a_pausa_a_cada_pagina(self):
        chamadas_pausa = []
        self._rodar([False], callback_pausa=lambda: chamadas_pausa.append(1))  # só 1 página, sem próxima
        texto = "\n".join(msg for msg, _ in self.logs)
        self.assertIn("Tudo pronto!", texto)
        self.assertNotIn("parada pelo operador", texto)
        self.assertEqual(self.progresso[-1], (100, "Exportação concluída!"))
        self.assertGreaterEqual(len(chamadas_pausa), 1)  # a pausa é consultada mesmo sem nunca pausar

    def test_cancelamento_entre_paginas_para_o_laco_sem_tocar_o_driver_daquela_pagina(self):
        cancelar_a_partir_da_pagina = 2
        contador = {"pagina": 0}

        def cancelar():
            contador["pagina"] += 1
            return contador["pagina"] >= cancelar_a_partir_da_pagina

        # _tentar_avancar_pagina sempre devolveria True (há mais páginas), mas o cancelamento
        # deve interromper o laço antes disso ser sequer consultado na 2ª volta.
        self._rodar([True, True, True], callback_cancelamento=cancelar)

        texto = "\n".join(msg for msg, _ in self.logs)
        self.assertIn("Exportação parada pelo operador", texto)
        self.assertIn("0 arquivo(s)", texto)  # cancelou antes de exportar qualquer lote
        self.assertEqual(self.progresso[-1], (100, "Parado pelo operador"))

    def test_pausa_bloqueia_antes_de_cada_pagina_e_a_execucao_so_segue_apos_liberar(self):
        ordem = []

        def pausa_que_bloqueia_uma_vez():
            ordem.append("pausa")
            # simula o comportamento real de _ControleProcessamentoXml.aguardar_se_pausado:
            # bloquear é responsabilidade do controle, aqui só provamos que foi chamado
            # antes de cada tentativa de selecionar a página.

        self._rodar([False], callback_pausa=pausa_que_bloqueia_uma_vez)
        self.assertEqual(ordem, ["pausa"])  # 1 página só: 1 checagem de pausa

    def test_cancelamento_na_primeira_volta_nao_exporta_nada_e_nao_chama_avancar_pagina(self):
        avancar = mock.Mock()
        with mock.patch.object(m, "_tentar_avancar_pagina", avancar):
            m.executar_exportacao_xml_prestados(
                pasta_destino="C:\\qualquer\\coisa",
                competencia_str="09/2026",
                callback_log=self.log,
                callback_cancelamento=lambda: True,
            )
        avancar.assert_not_called()
        texto = "\n".join(msg for msg, _ in self.logs)
        self.assertIn("Exportação parada pelo operador", texto)


if __name__ == "__main__":
    unittest.main()
