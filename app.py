import os
import sys
import json
import threading
from pathlib import Path
import traceback

import webview

import re

# Importar lógicas do projeto
from iss_fortaleza_automacao import executar_automacao_iss, interpretar_competencia, AutomacaoCanceladaError
from exportador_xml_prestados import executar_exportacao_xml_prestados
from tratamento_erros import registrar_erro, registrar_evento_execucao, configurar_pastas_logs
from processamento_xml import processar_pasta_xmls_auto

class GUIStdoutWrapper:
    def __init__(self, original_stdout, api):
        self.original_stdout = original_stdout
        self.api = api

    def write(self, text):
        self.original_stdout.write(text)
        self.original_stdout.flush()
        
        if hasattr(self.api, "_gui_log_callback") and self.api._gui_log_callback:
            lines = text.splitlines()
            for line in lines:
                stripped = line.strip()
                if stripped:
                    try:
                        self.api._gui_log_callback(stripped)
                    except Exception:
                        pass

    def flush(self):
        self.original_stdout.flush()

class BeAContabAPI:
    def __init__(self, window):
        self.window = window
        self._pausa_event = threading.Event()
        self._pausa_event.set()
        self._pausa_automacao_event = threading.Event()
        self._pausa_automacao_event.set()
        self._cancelar_automacao_event = threading.Event()
        self._cancelar_automacao_event.clear()

        # Flags e lock para evitar a execução concorrente de múltiplos robôs de automação
        # O Lock garante atomicidade na verificação e alteração da flag (FALHA-03)
        self._automacao_em_execucao = False
        self._lock_automacao = threading.Lock()
        
        # Carrega chaves salvas para as variáveis de ambiente na inicialização
        try:
            chaves = self.obter_chaves_salvas()
            if chaves.get("tess_key"):
                os.environ["TESS_API_KEY"] = chaves["tess_key"]
            if chaves.get("tess_agent_id"):
                os.environ["TESS_AGENT_ID"] = chaves["tess_agent_id"]
        except Exception as e:
            print(f"Erro ao carregar chaves na inicialização: {e}")

    def pausar_automacao(self):
        self._pausa_automacao_event.clear()

    def retomar_automacao(self):
        self._pausa_automacao_event.set()

    def cancelar_automacao(self):
        """Sinaliza o cancelamento da automação em execução.

        Também libera uma pausa em andamento (`.set()` do evento de pausa), senão
        o cancelamento ficaria preso esperando o operador retomar antes de poder
        ser sentido pelo robô.
        """
        self._cancelar_automacao_event.set()
        self._pausa_automacao_event.set()

    def selecionar_pasta(self):
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            return result[0]
        return None

    def selecionar_arquivo(self):
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=('Excel Files (*.xlsx)', 'All files (*.*)')
        )
        if result:
            # MELHORIA-06: Sinaliza à GUI que o arquivo foi selecionado nesta sessão,
            # permitindo que a validação visual só exiba aviso de arquivo movido em
            # reutilizações de sessão anteriores.
            try:
                self.window.evaluate_js("window._arquivoExcelSelecionadoNestaSessao = true;")
            except Exception:
                pass
            return result[0]
        return None

    def abrir_link(self, url):
        import webbrowser
        webbrowser.open(url)
        return True

    def fechar_app(self):
        self.window.destroy()

    def obter_caminho_config(self):
        """Retorna o caminho do arquivo .env na pasta AppData do usuário para garantir gravação estável.
        
        Isso previne erros de permissão de escrita após a instalação na pasta Program Files do Windows.
        """
        appdata = os.environ.get("APPDATA")
        if appdata:
            pasta_config = Path(appdata) / "BeAContab"
        else:
            pasta_config = Path.home() / ".beacontab"
        
        try:
            pasta_config.mkdir(parents=True, exist_ok=True)
        except Exception:
            # Fallback em caso de erro na criação do diretório AppData
            return Path(__file__).parent / ".env"
            
        return pasta_config / ".env"

    def obter_chaves_salvas(self):
        """Retorna as chaves do Tess AI salvas no arquivo .env do usuário."""
        chaves = {"tess_key": "", "tess_agent_id": ""}
        
        # 1. Tenta obter da pasta AppData do usuário
        env_path = self.obter_caminho_config()
        
        # 2. Fallback para o diretório do script (útil em modo de desenvolvimento)
        if not env_path.exists():
            env_desenv = Path(__file__).parent / ".env"
            if env_desenv.exists():
                env_path = env_desenv
 
        if env_path.exists():
            try:
                conteudo = env_path.read_text(encoding="utf-8")
                match_tess_key = re.search(r"^TESS_API_KEY\s*=\s*(.*)$", conteudo, re.MULTILINE)
                match_tess_agent = re.search(r"^TESS_AGENT_ID\s*=\s*(.*)$", conteudo, re.MULTILINE)
                if match_tess_key:
                    chaves["tess_key"] = match_tess_key.group(1).strip()
                if match_tess_agent:
                    chaves["tess_agent_id"] = match_tess_agent.group(1).strip()
            except Exception as e:
                print(f"Erro ao ler .env: {e}")
                
        # 3. Fallback: Se não encontrou chaves da Tess no .env (ou o .env não existe/está vazio),
        # tenta carregar as chaves padrões injetadas durante o processo de build.
        if not chaves.get("tess_key") or not chaves.get("tess_agent_id"):
            try:
                # BUG-01 corrigido: removido espaço indevido entre 'default_' e 'keys'
                import default_keys
                if not chaves.get("tess_key") and getattr(default_keys, "TESS_API_KEY", ""):
                    chaves["tess_key"] = default_keys.TESS_API_KEY
                if not chaves.get("tess_agent_id") and getattr(default_keys, "TESS_AGENT_ID", ""):
                    chaves["tess_agent_id"] = default_keys.TESS_AGENT_ID
            except ImportError:
                pass
                
        return chaves

    def salvar_chaves(self, tess_key="", tess_agent_id=""):
        """Salva as chaves no arquivo .env do AppData do usuário e atualiza no ambiente local."""
        env_path = self.obter_caminho_config()
        template_path = Path(__file__).parent / ".env.example"
        
        conteudo = ""
        if env_path.exists():
            try:
                conteudo = env_path.read_text(encoding="utf-8")
            except Exception:
                pass
        
        # Se o .env do usuário não existe, tenta carregar como base o .env local ou o .env.example
        if not conteudo:
            env_desenv = Path(__file__).parent / ".env"
            if env_desenv.exists():
                try:
                    conteudo = env_desenv.read_text(encoding="utf-8")
                except Exception:
                    pass
            elif template_path.exists():
                try:
                    conteudo = template_path.read_text(encoding="utf-8")
                except Exception:
                    pass
                
        if not conteudo:
            conteudo = "TESS_API_KEY=\nTESS_AGENT_ID=\n"
            
        if re.search(r"^TESS_API_KEY\s*=", conteudo, re.MULTILINE):
            conteudo = re.sub(r"^TESS_API_KEY\s*=.*$", f"TESS_API_KEY={tess_key}", conteudo, flags=re.MULTILINE)
        else:
            conteudo += f"\nTESS_API_KEY={tess_key}"

        if re.search(r"^TESS_AGENT_ID\s*=", conteudo, re.MULTILINE):
            conteudo = re.sub(r"^TESS_AGENT_ID\s*=.*$", f"TESS_AGENT_ID={tess_agent_id}", conteudo, flags=re.MULTILINE)
        else:
            conteudo += f"\nTESS_AGENT_ID={tess_agent_id}"
            
        try:
            env_path.write_text(conteudo, encoding="utf-8")
            os.environ["TESS_API_KEY"] = tess_key
            os.environ["TESS_AGENT_ID"] = tess_agent_id
            return True
        except Exception as e:
            print(f"Erro ao salvar .env em {env_path}: {e}")
            return False

    def confirmar_login_feito(self):
        """Cria o arquivo de flag indicando que o login manual foi concluído pelo operador."""
        try:
            appdata = os.environ.get("APPDATA", str(Path.home()))
            flag_path = Path(appdata) / "BeAContab" / "brain" / "funcao2_login_ok.flag"
            flag_path.parent.mkdir(parents=True, exist_ok=True)
            flag_path.write_text("ok", encoding="utf-8")
            return True
        except Exception as e:
            print(f"Erro ao confirmar login manual: {e}")
            return False

    def iniciar_automacao(self, excel_path, competencia_str):
        # Verificação e ativação da flag em bloco atômico para evitar race condition (FALHA-03)
        with self._lock_automacao:
            if self._automacao_em_execucao:
                self.window.evaluate_js("window.log_automacao('Erro: A automação de escrituração já está em andamento no momento!', true)")
                return
            self._automacao_em_execucao = True

        self._pausa_automacao_event.set()
        self._cancelar_automacao_event.clear()
        threading.Thread(target=self._processar_automacao, args=(excel_path, competencia_str), daemon=True).start()

    def _processar_automacao(self, excel_path, competencia_str):
        # Define o callback para logs de automação na GUI
        def gui_log_callback(mensagem):
            # Usa json.dumps para escapar corretamente strings com aspas, barras e quebras de linha
            mensagem_js = json.dumps(mensagem)
            self.window.evaluate_js(f"window.log_automacao({mensagem_js})")
        self._gui_log_callback = gui_log_callback

        try:
            planilha = Path(excel_path)
            if not planilha.exists():
                self.window.evaluate_js(f"window.log_automacao('Erro: Planilha não encontrada.', true)")
                return

            # Configura diretório de log para a pasta log da planilha
            configurar_pastas_logs([planilha.parent])

            # Sanitização via json.dumps() para evitar injeção de JS caso o nome da planilha
            # contenha aspas, barras invertidas ou outros caracteres especiais (BUG-03)
            msg_inicio = json.dumps(f"Iniciando automação com a planilha {planilha.name}...")
            self.window.evaluate_js(f"window.log_automacao({msg_inicio})")

            try:
                competencia = interpretar_competencia(competencia_str)
            except Exception as e:
                mensagem_js = json.dumps(f"Erro na competência: {str(e)}")
                try:
                    self.window.evaluate_js(f"window.log_automacao({mensagem_js}, true)")
                except Exception:
                    pass
                return

            msg_competencia = json.dumps(f"Competência confirmada: {competencia.mes}/{competencia.ano}")
            self.window.evaluate_js(f"window.log_automacao({msg_competencia})")
            self.window.evaluate_js("window.log_automacao('Iniciando navegador Chrome (Selenium). Não feche a janela do robô!')")
            
            # Remove flag de login anterior se existir
            # Usa AppData para garantir permissão de escrita no executável instalado
            appdata = os.environ.get("APPDATA", str(Path.home()))
            flag_path = Path(appdata) / "BeAContab" / "brain" / "funcao2_login_ok.flag"
            if flag_path.exists():
                try:
                    flag_path.unlink()
                except Exception:
                    pass

            # Define o callback para atualizar a barra de progresso da automação
            def callback_progresso(indice, total):
                # Aguarda se a automação estiver pausada pelo usuário
                self._pausa_automacao_event.wait()
                
                porcentagem = int((indice / total) * 100)
                self.window.evaluate_js(f"window.update_progresso_automacao({porcentagem}, 'Escriturando {indice} de {total} notas')")

            # Define o callback para verificar a pausa em pontos de baixa latência do Selenium
            def callback_pausa():
                self._pausa_automacao_event.wait()

            # Define o callback para verificar o cancelamento solicitado pelo operador
            def callback_cancelamento():
                return self._cancelar_automacao_event.is_set()

            # FALHA-08: Exibe o card de confirmação de login AQUI, depois dos checks de
            # competência/planilha, garantindo que erros precoces não deixem o card preso visível.
            # O card é ocultado pelo bloco finally abaixo, em qualquer cenário de saída.
            try:
                self.window.evaluate_js("document.getElementById('card-confirmar-login').classList.remove('hidden')")
            except Exception:
                pass

            # Chama a execução principal do robô
            executar_automacao_iss(
                planilha,
                competencia,
                depuracao=False,
                callback_progresso=callback_progresso,
                aguardar_login_por_arquivo=True,
                caminho_confirmacao_login=flag_path,
                executando_em_gui=True,
                callback_pausa=callback_pausa,
                callback_cancelamento=callback_cancelamento
            )
            self.window.evaluate_js("window.update_progresso_automacao(100, 'Escrituração concluída!')")
            msg_fim = json.dumps("Execução do robô concluída.")
            self.window.evaluate_js(f"window.log_automacao({msg_fim})")

        except AutomacaoCanceladaError:
            msg_cancel = json.dumps(
                "Automação cancelada pelo operador. O navegador foi encerrado — "
                "selecione a planilha/competência e inicie novamente quando quiser."
            )
            try:
                self.window.evaluate_js(f"window.log_automacao({msg_cancel})")
            except Exception:
                pass
        except Exception as e:
            tb = traceback.format_exc()
            # Usa json.dumps para escapar corretamente caracteres especiais (aspas,
            # quebras de linha, etc.) presentes nas mensagens de exceção do Selenium
            # Registra o traceback completo no log da GUI para facilitar o diagnóstico
            print(f"[ERRO INTERNO] {tb}")
            mensagem_js = json.dumps(f"Erro crítico: {str(e)}")
            try:
                self.window.evaluate_js(f"window.log_automacao({mensagem_js}, true)")
            except Exception:
                pass
        finally:
            # Libera a flag com proteção do lock para consistência com a aquisição (FALHA-03)
            with self._lock_automacao:
                self._automacao_em_execucao = False
            self._gui_log_callback = None
            try:
                # FALHA-08: O card de login é SEMPRE ocultado no finally do Python,
                # garantindo que o card não fique preso visivel em nenhum cenário de erro
                self.window.evaluate_js("document.getElementById('card-confirmar-login').classList.add('hidden')")
                self.window.evaluate_js("document.getElementById('btn-pausar-retomar-automacao').classList.add('hidden')")
                self.window.evaluate_js("document.getElementById('btn-cancelar-automacao').classList.add('hidden')")
            except Exception:
                pass

    def processar_xmls_gui(self, pasta_origem: str, tess_key: str = "", tess_agent_id: str = ""):
        """
        Inicia o processamento de XMLs com classificação de CNAE via Tess AI
        e gravação na planilha a partir de uma thread.
        """
        threading.Thread(
            target=self._processar_xmls,
            args=(pasta_origem, tess_key, tess_agent_id),
            daemon=True
        ).start()

    def iniciar_exportacao_xml_gui(self, pasta_destino: str, competencia_str: str):
        """Dispara a automação de exportação de XMLs de Serviços Prestados em uma thread separada."""
        threading.Thread(
            target=self._executar_exportacao_xml,
            args=(pasta_destino, competencia_str),
            daemon=True
        ).start()

    def _processar_xmls(self, pasta_origem: str, tess_key: str = "", tess_agent_id: str = ""):
        """Thread que executa o processamento de XMLs."""
        def log_gui(msg: str, is_error: bool = False) -> None:
            try:
                mensagem_js = json.dumps(msg)
                erro_js = "true" if is_error else "false"
                self.window.evaluate_js(f"window.log_xml({mensagem_js}, {erro_js})")
            except Exception:
                pass

        def callback_progresso(porcentagem: int, status: str) -> None:
            try:
                status_js = json.dumps(status)
                self.window.evaluate_js(f"window.update_progresso_xml({porcentagem}, {status_js})")
            except Exception:
                pass

        try:
            log_gui("Iniciando varredura e importação de XMLs...")

            caminhos_resultado = processar_pasta_xmls_auto(
                pasta_origem=pasta_origem,
                tess_key=tess_key,
                tess_agent_id=tess_agent_id,
                callback_log=log_gui,
                callback_progresso=callback_progresso
            )
            log_gui("Processamento concluído com sucesso!")
            for caminho in caminhos_resultado:
                log_gui(f"Planilha gerada em: {caminho}")
            self.window.evaluate_js("window.update_progresso_xml(100, 'Concluído!')")

            if len(caminhos_resultado) == 1:
                # MELHORIA-05: Envia o caminho da planilha para a GUI exibir o botão "Abrir Planilha Gerada"
                caminho_js = json.dumps(str(caminhos_resultado[0]))
                try:
                    self.window.evaluate_js(f"window.mostrar_botao_planilha({caminho_js})")
                except Exception:
                    pass
            else:
                # Múltiplas subpastas processadas: não há um único arquivo para abrir,
                # então oferece abrir a pasta-pai com todas as planilhas geradas.
                pasta_js = json.dumps(str(pasta_origem))
                try:
                    self.window.evaluate_js(
                        f"window.mostrar_botao_pasta_resultados({pasta_js}, {len(caminhos_resultado)})"
                    )
                except Exception:
                    pass
        except Exception as e:
            tb = traceback.format_exc()
            print(f"[ERRO INTERNO XML] {tb}")
            log_gui(f"Erro crítico no processamento de XMLs: {e}", is_error=True)
            self.window.evaluate_js("window.update_progresso_xml(100, 'Falha no processamento')")

    def _executar_exportacao_xml(self, pasta_destino: str, competencia_str: str):
        """Thread que executa a exportação de XMLs de Serviços Prestados."""
        def log_gui(msg: str, is_error: bool = False) -> None:
            try:
                mensagem_js = json.dumps(msg)
                erro_js = "true" if is_error else "false"
                self.window.evaluate_js(f"window.log_exportar({mensagem_js}, {erro_js})")
            except Exception:
                pass

        def callback_progresso(pct: int, status: str) -> None:
            try:
                status_js = json.dumps(status)
                self.window.evaluate_js(f"window.update_progresso_exportar({pct}, {status_js})")
            except Exception:
                pass

        # Gera o caminho do arquivo de flag de confirmação de login
        appdata = os.environ.get("APPDATA", str(Path.home()))
        flag_path = Path(appdata) / "BeAContab" / "brain" / "funcao3_login_ok.flag"
        flag_path.parent.mkdir(parents=True, exist_ok=True)
        if flag_path.exists():
            try:
                flag_path.unlink()
            except Exception:
                pass

        # Exibe o card de confirmação de login na aba de exportação
        try:
            self.window.evaluate_js("document.getElementById('card-confirmar-login-exportar').classList.remove('hidden')")
            self.window.evaluate_js("document.getElementById('btn-iniciar-exportacao').disabled = true")
        except Exception:
            pass

        try:
            log_gui(f"Iniciando exportação de XMLs - Competência: {competencia_str}")
            executar_exportacao_xml_prestados(
                pasta_destino=pasta_destino,
                competencia_str=competencia_str,
                callback_log=log_gui,
                callback_progresso=callback_progresso,
                caminho_confirmacao_login=flag_path,
            )
        except Exception as e:
            tb = traceback.format_exc()
            print(f"[ERRO INTERNO EXPORTAÇÃO] {tb}")
            log_gui(f"Erro crítico na exportação de XMLs: {e}", is_error=True)
            self.window.evaluate_js("window.update_progresso_exportar(100, 'Falha na exportação')")
        finally:
            try:
                self.window.evaluate_js("document.getElementById('card-confirmar-login-exportar').classList.add('hidden')")
                self.window.evaluate_js("document.getElementById('btn-iniciar-exportacao').disabled = false")
            except Exception:
                pass

    def confirmar_login_exportacao(self):
        """Cria o arquivo de flag indicando que o login manual foi concluído pelo operador na aba de exportação."""
        try:
            appdata = os.environ.get("APPDATA", str(Path.home()))
            flag_path = Path(appdata) / "BeAContab" / "brain" / "funcao3_login_ok.flag"
            flag_path.parent.mkdir(parents=True, exist_ok=True)
            flag_path.write_text("ok", encoding="utf-8")
            return True
        except Exception as e:
            print(f"Erro ao confirmar login de exportação: {e}")
            return False

def obter_caminho_recurso(caminho_relativo: str) -> str:
    """Retorna o caminho absoluto do recurso, funcionando no desenvolvimento ou no executável empacotado."""
    if hasattr(sys, '_MEIPASS'):
        return str(Path(sys._MEIPASS) / caminho_relativo)
    return str((Path(__file__).parent / caminho_relativo).resolve())

if __name__ == '__main__':
    # Cria a janela do pywebview
    # A interface gráfica fica em gui/index.html
    html_path = obter_caminho_recurso("gui/index.html")
    
    window = webview.create_window('Automação: ISS Fortaleza', html_path, width=1280, height=800)
    api = BeAContabAPI(window)
    
    # Redireciona stdout e stderr para capturar prints do robô em tempo real na GUI
    sys.stdout = GUIStdoutWrapper(sys.stdout, api)
    sys.stderr = GUIStdoutWrapper(sys.stderr, api)

    window.expose(
        api.iniciar_automacao,
        api.selecionar_pasta,
        api.selecionar_arquivo,
        api.fechar_app,
        api.abrir_link,
        api.obter_chaves_salvas,
        api.salvar_chaves,
        api.confirmar_login_feito,
        api.pausar_automacao,
        api.retomar_automacao,
        api.cancelar_automacao,
        api.processar_xmls_gui,
        api.iniciar_exportacao_xml_gui,
        api.confirmar_login_exportacao,
    )
    
    webview.start(debug=False)
