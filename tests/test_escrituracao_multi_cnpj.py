import json
import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import escrituracao_multi_cnpj as m

CNPJ_A = "00584628000181"
CNPJ_B = "21345512000160"
CNPJ_C = "44679408000107"
CABECALHOS = ["ARQUIVO_XML", "NUMERO_NF", "CNPJ_TOMADOR", "LOGIN", "SENHA"]


def _planilha(pasta: Path, linhas, cabecalhos=None, nome="notas.xlsx") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.append(cabecalhos or CABECALHOS)
    for linha in linhas:
        ws.append(linha)
    caminho = pasta / nome
    wb.save(caminho)
    wb.close()
    return caminho


def _linhas(cnpj, login, senha, quantidade=3):
    return [[f"{cnpj}_{i}.xml", i, cnpj, login, senha] for i in range(quantidade)]


class AcessoFalso:
    """Implementação em memória de AcessoIss: registra as chamadas e falha sob demanda."""

    def __init__(self, erros_por_login=None, logout_ok=True, ao_logar=None):
        self.chamadas = []
        self.erros_por_login = erros_por_login or {}
        self.logout_ok = logout_ok
        self.ao_logar = ao_logar
        self._login_atual = None

    def login(self, login, senha):
        self._login_atual = login
        self.chamadas.append(("login", login))
        if self.ao_logar:
            self.ao_logar(login)
        erro = self.erros_por_login.get(login)
        if erro and erro[0] == "login":
            raise erro[1]

    def garantir_empresa(self, cnpj):
        self.chamadas.append(("garantir_empresa", cnpj))
        erro = self.erros_por_login.get(self._login_atual)
        if erro and erro[0] == "empresa":
            raise erro[1]

    def logout(self):
        self.chamadas.append(("logout", self._login_atual))
        return self.logout_ok

    def encerrar(self):
        self.chamadas.append(("encerrar", None))


class TestUtilitarios(unittest.TestCase):
    def test_normalizar_login(self):
        self.assertEqual(m.normalizar_login(5142974336), "05142974336")  # Excel perdeu o zero
        self.assertEqual(m.normalizar_login(5142974336.0), "05142974336")
        self.assertEqual(m.normalizar_login("051.429.743-36"), "05142974336")
        self.assertEqual(m.normalizar_login("  64345612345 "), "64345612345")
        self.assertEqual(m.normalizar_login("usuario.teste"), "usuario.teste")
        self.assertEqual(m.normalizar_login(None), "")

    def test_normalizar_senha_nao_altera_texto(self):
        self.assertEqual(m.normalizar_senha(" Ab#1 x "), " Ab#1 x ")
        self.assertEqual(m.normalizar_senha(123456), "123456")
        self.assertEqual(m.normalizar_senha(123456.0), "123456")
        self.assertEqual(m.normalizar_senha(None), "")

    def test_cnpj(self):
        self.assertEqual(m.normalizar_cnpj_celula(584628000181), "00584628000181")
        self.assertEqual(m.normalizar_cnpj_celula("00.584.628/0001-81"), CNPJ_A)
        self.assertEqual(m.formatar_cnpj(CNPJ_A), "00.584.628/0001-81")

    def test_mascarar_login_nunca_mostra_o_valor_todo(self):
        self.assertEqual(m.mascarar_login("05142974336"), "***.***.***-36")
        self.assertNotIn("0514297", m.mascarar_login("05142974336"))


class TestCarregarPlanilha(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="escmulti_"))

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_agrupa_por_cnpj_na_ordem_e_normaliza_credenciais(self):
        linhas = _linhas(CNPJ_A, 64345612345, "Senha#A", 10) + _linhas(CNPJ_B, "051.429.743-36", "Senha#B", 4)
        planilha = m.carregar_empresas_xlsx(_planilha(self.pasta, linhas))
        self.assertEqual([e.cnpj for e in planilha.empresas], [CNPJ_A, CNPJ_B])
        a, b = planilha.empresas
        self.assertTrue(a.valida and b.valida)
        self.assertEqual((a.qtd_notas, b.qtd_notas), (10, 4))
        self.assertEqual(a.login, "64345612345")  # número do Excel virou texto de 11 dígitos
        self.assertEqual(b.login, "05142974336")
        self.assertEqual(b.senha, "Senha#B")

    def test_senha_diferente_entre_linhas_do_mesmo_cnpj_e_invalida_sem_vazar_valores(self):
        # o autopreenchimento do Excel incrementa o número final: Senha#51, Senha#52...
        linhas = [[f"a{i}.xml", i, CNPJ_A, "05142974336", f"Segredo@{51 + i}"] for i in range(5)]
        empresa = m.carregar_empresas_xlsx(_planilha(self.pasta, linhas)).empresas[0]
        self.assertFalse(empresa.valida)
        texto = " ".join(empresa.problemas)
        self.assertIn("SENHA diferente", texto)
        self.assertNotIn("Segredo", texto)
        self.assertNotIn("0514297", texto)

    def test_login_ou_senha_em_branco_ou_parcial(self):
        linhas = (
            _linhas(CNPJ_A, None, "x", 2)
            + [[f"b{i}.xml", i, CNPJ_B, "05142974336", "" if i == 1 else "s"] for i in range(3)]
        )
        a, b = m.carregar_empresas_xlsx(_planilha(self.pasta, linhas)).empresas
        self.assertIn("LOGIN em branco", a.problemas)
        self.assertTrue(any("SENHA preenchido só em algumas linhas" in p for p in b.problemas))

    def test_cnpj_invalido(self):
        empresa = m.carregar_empresas_xlsx(_planilha(self.pasta, _linhas("12AB", "05142974336", "s", 1))).empresas[0]
        self.assertFalse(empresa.valida)
        self.assertTrue(any("14 caracteres" in p for p in empresa.problemas))

    def test_ignora_linhas_vazias_e_avisa_sobre_linhas_sem_cnpj(self):
        linhas = _linhas(CNPJ_A, "05142974336", "s", 2) + [[None] * 5] + [["x.xml", 1, None, "1", "2"]]
        planilha = m.carregar_empresas_xlsx(_planilha(self.pasta, linhas))
        self.assertEqual(len(planilha.empresas), 1)
        self.assertEqual(planilha.empresas[0].qtd_notas, 2)
        self.assertEqual(len(planilha.avisos_gerais), 1)

    def test_falta_de_colunas(self):
        with self.assertRaises(ValueError) as ctx:
            m.carregar_empresas_xlsx(_planilha(self.pasta, [], cabecalhos=["ARQUIVO_XML", "NUMERO_NF"]))
        self.assertIn("CNPJ_TOMADOR", str(ctx.exception))

    def test_senha_nao_aparece_no_repr(self):
        empresa = m.carregar_empresas_xlsx(_planilha(self.pasta, _linhas(CNPJ_A, "05142974336", "SegredoAbsoluto", 1))).empresas[0]
        self.assertNotIn("SegredoAbsoluto", repr(empresa))


class TestOrquestracao(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="escmulti_"))
        self.logs = []
        self.log = lambda msg, erro=False: self.logs.append((msg, erro))

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def _duas_empresas(self):
        return _planilha(
            self.pasta,
            _linhas(CNPJ_A, "11111111111", "SenhaSecretaA", 3) + _linhas(CNPJ_B, "22222222222", "SenhaSecretaB", 2),
        )

    def _executar(self, planilha, acesso, **kwargs):
        criados = kwargs.pop("criados", None)

        def criar():
            if criados is not None:
                criados.append(1)
            return acesso

        return m.executar_escrituracao_multi_cnpj(planilha, criar, callback_log=self.log, **kwargs)

    def test_validacao_de_acessos_faz_login_confere_e_faz_logout_de_cada_empresa(self):
        acesso = AcessoFalso()
        resultado = self._executar(self._duas_empresas(), acesso)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_ACESSO_VALIDADO] * 2)
        self.assertEqual(
            acesso.chamadas,
            [
                ("login", "11111111111"), ("garantir_empresa", CNPJ_A), ("logout", "11111111111"),
                ("login", "22222222222"), ("garantir_empresa", CNPJ_B), ("logout", "22222222222"),
                ("encerrar", None),
            ],
        )

    def test_senha_nunca_vai_para_o_log_e_login_sai_mascarado(self):
        self._executar(self._duas_empresas(), AcessoFalso())
        todo_o_log = "\n".join(msg for msg, _ in self.logs)
        self.assertNotIn("SenhaSecreta", todo_o_log)
        self.assertNotIn("11111111111", todo_o_log)
        self.assertIn("***.***.***-11", todo_o_log)

    def test_falha_em_uma_empresa_nao_derruba_as_outras_e_o_logout_acontece_mesmo_assim(self):
        acesso = AcessoFalso({"11111111111": ("login", m.LoginRecusadoError("o portal recusou o login"))})
        resultado = self._executar(self._duas_empresas(), acesso)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_LOGIN_RECUSADO, m.STATUS_ACESSO_VALIDADO])
        self.assertEqual([c for c in acesso.chamadas if c[0] == "logout"], [("logout", "11111111111"), ("logout", "22222222222")])
        self.assertIn(("login", "22222222222"), acesso.chamadas)

    def test_empresa_nao_encontrada_e_cnpj_divergente(self):
        acesso = AcessoFalso({
            "11111111111": ("empresa", m.EmpresaNaoEncontradaError("não aparece na lista")),
            "22222222222": ("empresa", m.CnpjDivergenteError("inscrição de outro CNPJ")),
        })
        resultado = self._executar(self._duas_empresas(), acesso)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_EMPRESA_NAO_ENCONTRADA, m.STATUS_CNPJ_DIVERGENTE])

    def test_erro_inesperado_e_sanitizado(self):
        acesso = AcessoFalso({"11111111111": ("login", RuntimeError("falha ao digitar SenhaSecretaA no campo"))})
        resultado = self._executar(self._duas_empresas(), acesso)
        self.assertEqual(resultado.empresas[0].status, m.STATUS_ERRO)
        self.assertNotIn("SenhaSecretaA", resultado.empresas[0].mensagem)
        self.assertNotIn("SenhaSecretaA", "\n".join(msg for msg, _ in self.logs))
        self.assertEqual(resultado.empresas[1].status, m.STATUS_ACESSO_VALIDADO)

    def test_logout_nao_confirmado_reabre_o_navegador_para_a_proxima_empresa(self):
        criados = []
        acessos = [AcessoFalso(logout_ok=False), AcessoFalso()]

        def criar():
            criados.append(1)
            return acessos[len(criados) - 1]

        m.executar_escrituracao_multi_cnpj(self._duas_empresas(), criar, callback_log=self.log)
        self.assertEqual(len(criados), 2)
        self.assertIn(("encerrar", None), acessos[0].chamadas)  # sessão suja descartada
        self.assertIn(("login", "22222222222"), acessos[1].chamadas)

    def test_logout_confirmado_e_registrado_no_log(self):
        self._executar(self._duas_empresas(), AcessoFalso())
        self.assertEqual(sum(1 for msg, erro in self.logs if "Logout confirmado" in msg and not erro), 2)

    def test_logout_falho_na_ultima_empresa_so_fecha_o_navegador_e_informa_o_detalhe(self):
        acesso = AcessoFalso(logout_ok=False)
        acesso.detalhe_logout = "tentativa 1; clique em Sair bloqueado; página: https://iss/x SenhaSecretaB"
        criados = []
        planilha = _planilha(self.pasta, _linhas(CNPJ_B, "22222222222", "SenhaSecretaB", 2))
        self._executar(planilha, acesso, criados=criados)
        aviso = next(msg for msg, erro in self.logs if "Não consegui confirmar o logout" in msg)
        self.assertIn("apenas fechar o navegador", aviso)
        self.assertNotIn("reabrir", aviso)
        self.assertIn("clique em Sair bloqueado", aviso)
        self.assertNotIn("SenhaSecretaB", aviso)
        self.assertEqual(len(criados), 1)  # não abriu outro navegador

    def test_logout_falho_com_mais_empresas_avisa_que_vai_reabrir(self):
        self._executar(self._duas_empresas(), AcessoFalso(logout_ok=False))
        aviso = next(msg for msg, erro in self.logs if "Não consegui confirmar o logout" in msg)
        self.assertIn("reabrir o navegador", aviso)

    def test_dados_invalidos_da_planilha_pulam_a_empresa_e_nao_abrem_navegador_se_nao_ha_nada_a_fazer(self):
        planilha = _planilha(self.pasta, _linhas(CNPJ_A, None, "s", 2))
        criados = []
        resultado = self._executar(planilha, AcessoFalso(), criados=criados)
        self.assertEqual(resultado.empresas[0].status, m.STATUS_PLANILHA_INVALIDA)
        self.assertEqual(criados, [])

    def test_cancelar_entre_empresas_marca_as_restantes_e_faz_logout_da_atual(self):
        acesso = AcessoFalso()
        cancelar = {"v": False}
        acesso.ao_logar = lambda login: cancelar.__setitem__("v", True)  # operador cancela durante a 1ª empresa
        resultado = self._executar(self._duas_empresas(), acesso, callback_cancelamento=lambda: cancelar["v"])
        self.assertTrue(resultado.cancelado)
        self.assertEqual(resultado.empresas[0].status, m.STATUS_ACESSO_VALIDADO)  # terminou a atual
        self.assertEqual(resultado.empresas[1].status, m.STATUS_CANCELADA)
        self.assertIn(("logout", "11111111111"), acesso.chamadas)
        self.assertNotIn(("login", "22222222222"), acesso.chamadas)
        self.assertEqual(acesso.chamadas[-1], ("encerrar", None))

    def test_cancelamento_durante_o_login_faz_logout_e_encerra(self):
        acesso = AcessoFalso({"11111111111": ("login", m.OperacaoCanceladaError())})
        resultado = self._executar(self._duas_empresas(), acesso)
        self.assertTrue(resultado.cancelado)
        self.assertEqual([r.status for r in resultado.empresas], [m.STATUS_CANCELADA, m.STATUS_CANCELADA])
        self.assertIn(("logout", "11111111111"), acesso.chamadas)

    def test_pausa_e_consultada_a_cada_empresa(self):
        chamadas = []
        self._executar(self._duas_empresas(), AcessoFalso(), callback_pausa=lambda: chamadas.append(1))
        self.assertEqual(len(chamadas), 2)

    def test_resumo_para_a_gui_nao_tem_dados_sensiveis(self):
        resultado = self._executar(self._duas_empresas(), AcessoFalso())
        texto = json.dumps(resultado.para_json(), ensure_ascii=False)
        self.assertIn("00.584.628/0001-81", texto)
        self.assertNotIn("SenhaSecreta", texto)
        self.assertEqual(resultado.para_json()["contagem"], {"Acesso validado": 2})

    def test_modo_validacao_nao_cria_arquivo_de_retomada(self):
        planilha = self._duas_empresas()
        self._executar(planilha, AcessoFalso())
        self.assertFalse(m.caminho_estado(planilha).exists())


class _Elemento:
    def __init__(self, ao_clicar=None):
        self._ao_clicar = ao_clicar

    def is_displayed(self):
        return True

    def click(self):
        if self._ao_clicar:
            self._ao_clicar()


class _NavegadorFalso:
    """Navegador mínimo para exercitar o logout. Telas: `logado` (ISS com link Sair), `confirmacao`
    (login único pedindo confirmação), `home` (portal deslogado), `iss_logado_sem_idp` (ISS ainda
    logado depois que o login único já saiu)."""

    def __init__(self, sair_bloqueado=False, confirma_logout=True, alerta_pendente=False, iss_sobrevive=False):
        self.tela = "logado"
        self.sair_bloqueado = sair_bloqueado
        self.confirma_logout = confirma_logout
        self.alerta_pendente = alerta_pendente
        self.iss_sobrevive = iss_sobrevive
        self.scripts = []
        self.alertas_aceitos = 0
        self.cliques_por_script = 0
        self.current_url = "https://iss.fortaleza.ce.gov.br/grpfor/home.seam?x=1"
        self.switch_to = self

    # switch_to.alert.accept()
    @property
    def alert(self):
        if not self.alerta_pendente:
            raise RuntimeError("sem alerta")
        return self

    def accept(self):
        self.alerta_pendente = False
        self.alertas_aceitos += 1

    def execute_script(self, script, *args):
        self.scripts.append(script)
        if "arguments[0].click()" in script:
            self.cliques_por_script += 1
            self._sair_iss(forcado=True)
            return None
        return 1 if "hideModalPanel" in script else None

    def get(self, url):
        self.tela = "confirmacao" if "logout" in url else "home"

    def _sair_iss(self, forcado=False):
        if self.sair_bloqueado and not forcado:
            raise RuntimeError("element click intercepted: Other element would receive the click: <div id=modal>")
        self.tela = "home" if self.tela == "iss_logado_sem_idp" else "confirmacao"

    def _confirmar(self):
        self.tela = "iss_logado_sem_idp" if self.iss_sobrevive else "home"
        self.iss_sobrevive = False

    def find_elements(self, by, seletor):
        if self.alerta_pendente:
            raise RuntimeError("unexpected alert open")
        if "identity.logout" in seletor:
            return [_Elemento(self._sair_iss)] if self.tela in ("logado", "iss_logado_sem_idp") else []
        if seletor == "kc-logout":
            return [_Elemento(self._confirmar)] if self.tela == "confirmacao" and self.confirma_logout else []
        if seletor == 'a[href="/grpfor/oauth2/login"]':
            return [_Elemento()] if self.tela == "home" else []
        return []

    def find_element(self, by, seletor):
        encontrados = self.find_elements(by, seletor)
        if not encontrados:
            raise RuntimeError("no such element")
        return encontrados[0]

    def quit(self):
        pass


class TestLogoutSelenium(unittest.TestCase):
    def _acesso(self, navegador):
        acesso = m.AcessoIssSelenium(abrir_navegador=lambda: navegador, timeout=1)
        acesso.tempo_maximo_logout = 1.0  # acelera os testes
        acesso.pausa_logout = 0.02
        acesso._navegador()  # logout() não faz nada se o navegador nunca foi aberto
        return acesso

    def test_logout_normal(self):
        acesso = self._acesso(_NavegadorFalso())
        self.assertTrue(acesso.logout())
        self.assertEqual(acesso.detalhe_logout, "")

    def test_botao_sair_bloqueado_por_modal_fecha_os_modais_e_forca_o_clique(self):
        navegador = _NavegadorFalso(sair_bloqueado=True)
        acesso = self._acesso(navegador)
        self.assertTrue(acesso.logout())
        self.assertTrue(any("hideModalPanel" in s for s in navegador.scripts))
        self.assertTrue(any("onbeforeunload" in s for s in navegador.scripts))
        self.assertEqual(navegador.cliques_por_script, 1)

    def test_login_unico_sai_mas_o_iss_continua_logado_e_o_robo_sai_do_iss_de_novo(self):
        # caso do teste real: home.seam aparecia sem o link "Fazer login" porque o ISS ainda estava logado
        navegador = _NavegadorFalso(sair_bloqueado=True, iss_sobrevive=True)
        acesso = self._acesso(navegador)
        self.assertTrue(acesso.logout())
        self.assertEqual(navegador.tela, "home")

    def test_alerta_de_sair_da_pagina_e_aceito(self):
        navegador = _NavegadorFalso(alerta_pendente=True)
        acesso = self._acesso(navegador)
        self.assertTrue(acesso.logout())
        self.assertEqual(navegador.alertas_aceitos, 1)

    def test_falha_devolve_false_com_diagnostico_sem_query_string(self):
        navegador = _NavegadorFalso(confirma_logout=False)
        acesso = self._acesso(navegador)
        self.assertFalse(acesso.logout())
        self.assertIn("tempo esgotado", acesso.detalhe_logout)
        self.assertIn("página: https://iss.fortaleza.ce.gov.br/grpfor/home.seam", acesso.detalhe_logout)
        self.assertNotIn("x=1", acesso.detalhe_logout)

    def test_diagnostico_registra_o_motivo_do_clique_bloqueado(self):
        navegador = _NavegadorFalso(sair_bloqueado=True, confirma_logout=False)
        acesso = self._acesso(navegador)
        self.assertFalse(acesso.logout())
        self.assertIn("clique em Sair bloqueado (RuntimeError: element click intercepted", acesso.detalhe_logout)


class TestRetomada(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="escmulti_"))
        self.planilha = _planilha(
            self.pasta, _linhas(CNPJ_A, "11111111111", "a", 2) + _linhas(CNPJ_B, "22222222222", "b", 2)
        )

    def tearDown(self):
        shutil.rmtree(self.pasta, ignore_errors=True)

    def test_empresas_concluidas_sao_puladas_e_o_estado_some_quando_tudo_termina(self):
        processadas = []
        falhar_b = {"v": True}

        def processar(empresa, acesso):
            if empresa.cnpj == CNPJ_B and falhar_b["v"]:
                falhar_b["v"] = False
                raise RuntimeError("caiu no meio da escrituração")
            processadas.append(empresa.cnpj)
            return empresa.qtd_notas

        # 1ª execução: a empresa B falha
        r1 = m.executar_escrituracao_multi_cnpj(self.planilha, AcessoFalso, processar_empresa=processar, callback_log=lambda *_: None)
        self.assertEqual([r.status for r in r1.empresas], [m.STATUS_CONCLUIDA, m.STATUS_ERRO])
        self.assertEqual(m.ler_empresas_concluidas(self.planilha), {CNPJ_A})
        self.assertNotIn("a", m.caminho_estado(self.planilha).read_text(encoding="utf-8").replace("empresas_concluidas", ""))  # sem senha

        # 2ª execução: A é pulada, B é tentada de novo e conclui
        r2 = m.executar_escrituracao_multi_cnpj(self.planilha, AcessoFalso, processar_empresa=processar, callback_log=lambda *_: None)
        self.assertEqual([r.status for r in r2.empresas], [m.STATUS_JA_CONCLUIDA, m.STATUS_CONCLUIDA])
        self.assertEqual(processadas, [CNPJ_A, CNPJ_B])
        self.assertFalse(m.caminho_estado(self.planilha).exists())  # terminou tudo: nada a retomar


if __name__ == "__main__":
    unittest.main()
