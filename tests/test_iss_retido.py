import unittest

from iss_fortaleza_automacao import _interpretar_iss_retido


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


if __name__ == "__main__":
    unittest.main()
