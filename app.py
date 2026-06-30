import os
import sys
import threading
from pathlib import Path
import traceback

import webview

import re

# Importar lógicas do projeto
from extrair_nf_pdfs import listar_pdfs, extrair_nota_fiscal, extrair_nota_fiscal_somente_texto, gerar_xlsx, registrar_log_incompletos, registrar_incompleto_realtime, NotaFiscalExtraida
from gemini_extracao import ErroCotaGemini, ErroCotaGroq
from iss_fortaleza_automacao import executar_automacao_iss, solicitar_competencia, interpretar_competencia
from tratamento_erros import registrar_erro, registrar_evento_execucao, configurar_pasta_logs, configurar_pastas_logs
from extracao_prefeituras import processar_pasta_prefeitura, PREFEITURAS_SUPORTADAS, organizar_pasta_pdfs
from unir_planilhas import unir_xlsx

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

    def pausar_extracao(self):
        self._pausa_event.clear()

    def retomar_extracao(self):
        self._pausa_event.set()

    def pausar_automacao(self):
        self._pausa_automacao_event.clear()

    def retomar_automacao(self):
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
            return result[0]
        return None

    def selecionar_multiplos_xlsx(self):
        """Abre o seletor de múltiplos arquivos XLSX para a funcionalidade de união."""
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=('Excel Files (*.xlsx)', 'All files (*.*)')
        )
        if result:
            return list(result)
        return []

    def exportar_por_prefeitura(self, mapa_prefeituras_pastas: dict):
        """
        Recebe um dicionário {nome_prefeitura: caminho_pasta} e processa
        cada pasta em thread separada, emitindo logs em tempo real na GUI.
        """
        # Garante que a extração começa não pausada
        self._pausa_event.set()

        threading.Thread(
            target=self._processar_exportar_prefeituras,
            args=(mapa_prefeituras_pastas,),
            daemon=True
        ).start()

    def _processar_exportar_prefeituras(self, mapa_prefeituras_pastas: dict):
        """Thread de processamento da exportação estruturada por prefeitura."""
        def log_gui(msg: str) -> None:
            msg_safe = msg.replace("'", "\\'").replace('"', '\\"')
            try:
                self.window.evaluate_js(f"window.log_prefeitura('{msg_safe}')")
            except Exception:
                pass

        total_xlsx = []
        total_notas = 0

        # Primeiro conta todos os PDFs para cálculo global do progresso
        total_pdfs_global = 0
        pastas_validas = []

        for nome_prefeitura, caminho_pasta in mapa_prefeituras_pastas.items():
            if not caminho_pasta:
                continue
            pasta = Path(caminho_pasta)
            if pasta.exists() and pasta.is_dir():
                pdfs_da_pasta = [p for p in pasta.glob("*.pdf") if p.parent == pasta]
                if pdfs_da_pasta:
                    total_pdfs_global += len(pdfs_da_pasta)
                    pastas_validas.append((nome_prefeitura, pasta, len(pdfs_da_pasta)))

        if total_pdfs_global == 0:
            log_gui("Nenhum arquivo PDF encontrado para processar.")
            try:
                self.window.evaluate_js(f"window.update_progresso_prefeitura(100, 'Concluído!')")
            except Exception:
                pass
            return

        processados_global = 0

        for nome_prefeitura, pasta, qtd_pdfs in pastas_validas:
            log_gui(f"=== Iniciando: {nome_prefeitura} ===")

            def callback_progresso(indice, total):
                nao_local = processados_global + indice
                porcentagem = int((nao_local / total_pdfs_global) * 100)
                try:
                    self.window.evaluate_js(
                        f"window.update_progresso_prefeitura({porcentagem}, 'Processando {nao_local} de {total_pdfs_global}')"
                    )
                except Exception:
                    pass

            try:
                registros, xlsx = processar_pasta_prefeitura(
                    pasta,
                    nome_prefeitura,
                    callback_log=log_gui,
                    callback_progresso=callback_progresso,
                )
                total_notas += len(registros)
                if xlsx:
                    total_xlsx.append(str(xlsx))
            except Exception as e:
                log_gui(f"[ERRO] {nome_prefeitura}: {e}")

            processados_global += qtd_pdfs

        msg_final = f"Exportação concluída. {total_notas} nota(s) extraída(s) em {len(total_xlsx)} planilha(s)."
        log_gui(msg_final)
        try:
            self.window.evaluate_js(f"window.update_progresso_prefeitura(100, 'Concluído!')")
        except Exception:
            pass

    def unir_planilhas_gui(self, lista_xlsx: list[str], pasta_destino: str, remover_duplicatas: bool = True):
        """
        Consolida múltiplos arquivos XLSX ou pastas de origem em um único arquivo de saída.
        Executado em thread separada.
        """
        # Garante que a extração começa não pausada
        self._pausa_event.set()

        threading.Thread(
            target=self._processar_uniao_planilhas,
            args=(lista_xlsx, pasta_destino, remover_duplicatas),
            daemon=True
        ).start()

    def _processar_uniao_planilhas(self, lista_xlsx: list[str], pasta_destino: str, remover_duplicatas: bool):
        """Thread de execução da união de planilhas."""
        def log_gui(msg: str) -> None:
            msg_safe = msg.replace("'", "\\'").replace('"', '\\"')
            try:
                self.window.evaluate_js(f"window.log_uniao('{msg_safe}')")
            except Exception:
                pass

        try:
            if not lista_xlsx:
                log_gui("[ERRO] Nenhum arquivo ou pasta XLSX selecionada.")
                return
            if not pasta_destino:
                log_gui("[ERRO] Nenhuma pasta de destino selecionada.")
                return

            log_gui("Iniciando consolidação das planilhas...")
            caminho_resultado = unir_xlsx(
                lista_caminhos=lista_xlsx,
                caminho_destino=pasta_destino,
                remover_duplicatas=remover_duplicatas,
                callback_log=log_gui,
            )
            log_gui(f"Concluído! Arquivo unificado salvo em: {caminho_resultado}")
            try:
                self.window.evaluate_js("window.update_progresso_uniao(100, 'Concluído!')")
            except Exception:
                pass
        except Exception as e:
            log_gui(f"[ERRO] Falha ao unir planilhas: {e}")

    def obter_prefeituras_suportadas(self):
        """Retorna a lista de prefeituras suportadas para popular a GUI."""
        return list(PREFEITURAS_SUPORTADAS.keys())

    def organizar_pdfs_por_prefeitura_gui(self, pasta_origem: str):
        """
        Organiza os PDFs da pasta_origem movendo cada um para a subpasta da respectiva prefeitura.
        Executado em thread separada.
        """
        # Garante que a extração começa não pausada
        self._pausa_event.set()

        threading.Thread(
            target=self._processar_organizar_pdfs,
            args=(pasta_origem,),
            daemon=True
        ).start()

    def _processar_organizar_pdfs(self, pasta_origem: str):
        """Thread de execução da organização de PDFs."""
        def log_gui(msg: str) -> None:
            msg_safe = msg.replace("'", "\\'").replace('"', '\\"')
            try:
                self.window.evaluate_js(f"window.log_organizador('{msg_safe}')")
            except Exception:
                pass

        try:
            if not pasta_origem:
                log_gui("[ERRO] Nenhuma pasta de origem selecionada.")
                return
            pasta = Path(pasta_origem)
            if not pasta.exists() or not pasta.is_dir():
                log_gui(f"[ERRO] Pasta inválida ou inexistente: {pasta_origem}")
                return

            def callback_progresso(indice, total):
                porcentagem = int((indice / total) * 100)
                try:
                    self.window.evaluate_js(
                        f"window.update_progresso_organizador({porcentagem}, 'Organizando {indice} de {total}')"
                    )
                except Exception:
                    pass

            sucessos, ignorados = organizar_pasta_pdfs(
                pasta,
                callback_log=log_gui,
                callback_progresso=callback_progresso,
            )
            log_gui(f"Organização concluída! {sucessos} PDFs organizados, {ignorados} não movidos.")
            try:
                self.window.evaluate_js("window.update_progresso_organizador(100, 'Concluído!')")
            except Exception:
                pass
        except Exception as e:
            log_gui(f"[ERRO] Falha durante a organização: {e}")

    def abrir_link(self, url):
        import webbrowser
        webbrowser.open(url)
        return True

    def fechar_app(self):
        self.window.destroy()

    def obter_chaves_salvas(self):
        """Retorna as chaves salvas no arquivo .env."""
        chaves = {"gemini": "", "groq": ""}
        env_path = Path(__file__).parent / ".env"
        if env_path.exists():
            try:
                conteudo = env_path.read_text(encoding="utf-8")
                match_gemini = re.search(r"^GEMINI_API_KEY\s*=\s*(.*)$", conteudo, re.MULTILINE)
                match_groq = re.search(r"^GROQ_API_KEY\s*=\s*(.*)$", conteudo, re.MULTILINE)
                if match_gemini:
                    chaves["gemini"] = match_gemini.group(1).strip()
                if match_groq:
                    chaves["groq"] = match_groq.group(1).strip()
            except Exception as e:
                print(f"Erro ao ler .env: {e}")
        return chaves

    def salvar_chaves(self, gemini_key, groq_key):
        """Salva as chaves no arquivo .env e atualiza no ambiente local."""
        env_path = Path(__file__).parent / ".env"
        template_path = Path(__file__).parent / ".env.example"
        
        conteudo = ""
        if env_path.exists():
            try:
                conteudo = env_path.read_text(encoding="utf-8")
            except Exception:
                pass
        
        if not conteudo and template_path.exists():
            try:
                conteudo = template_path.read_text(encoding="utf-8")
            except Exception:
                pass
                
        if not conteudo:
            conteudo = "GEMINI_API_KEY=\nGROQ_API_KEY=\n"
            
        if re.search(r"^GEMINI_API_KEY\s*=", conteudo, re.MULTILINE):
            conteudo = re.sub(r"^GEMINI_API_KEY\s*=.*$", f"GEMINI_API_KEY={gemini_key}", conteudo, flags=re.MULTILINE)
        else:
            conteudo += f"\nGEMINI_API_KEY={gemini_key}"
            
        if re.search(r"^GROQ_API_KEY\s*=", conteudo, re.MULTILINE):
            conteudo = re.sub(r"^GROQ_API_KEY\s*=.*$", f"GROQ_API_KEY={groq_key}", conteudo, flags=re.MULTILINE)
        else:
            conteudo += f"\nGROQ_API_KEY={groq_key}"
            
        try:
            env_path.write_text(conteudo, encoding="utf-8")
            os.environ["GEMINI_API_KEY"] = gemini_key
            os.environ["GROQ_API_KEY"] = groq_key
            return True
        except Exception as e:
            print(f"Erro ao salvar .env: {e}")
            return False

    def contar_pdfs_origem(self, pasta_caminho):
        """Retorna a contagem de PDFs na pasta de origem informada."""
        if not pasta_caminho:
            return 0
        try:
            pasta = Path(pasta_caminho)
            if pasta.exists() and pasta.is_dir():
                return len(list(listar_pdfs(pasta)))
        except Exception:
            pass
        return 0

    def confirmar_login_feito(self):
        """Cria o arquivo flag sinalizando que o login manual foi concluído no navegador."""
        try:
            flag_path = Path("brain/funcao2_login_ok.flag")
            flag_path.parent.mkdir(parents=True, exist_ok=True)
            flag_path.write_text("ok", encoding="utf-8")
            return True
        except Exception as e:
            print(f"Erro ao confirmar login por arquivo: {e}")
            return False

    def iniciar_extracao(self, origem, destino, gemini_key, groq_key, usar_gemini, usar_groq):
        # Configurar chaves no processo
        if gemini_key:
            os.environ["GEMINI_API_KEY"] = gemini_key
        if groq_key:
            os.environ["GROQ_API_KEY"] = groq_key
        
        # Garante que a extração comece não pausada
        self._pausa_event.set()
        
        # Iniciar em thread separada para não travar a UI
        threading.Thread(
            target=self._processar_extracao,
            args=(origem, destino, gemini_key, groq_key, usar_gemini, usar_groq),
            daemon=True
        ).start()

    def _processar_extracao(self, origem, destino, gemini_key, groq_key, usar_gemini, usar_groq):
        # Define o callback para logs de extração na GUI
        def gui_log_callback(mensagem):
            msg_safe = mensagem.replace("'", "\\'").replace('"', '\\"')
            self.window.evaluate_js(f"window.log_extracao('{msg_safe}')")
        self._gui_log_callback = gui_log_callback

        try:
            pasta = Path(origem)
            if not pasta.exists() or not pasta.is_dir():
                self.window.evaluate_js(f"window.log_extracao('Erro: Pasta de origem inválida.', true)")
                return
            
            caminho_destino = Path(destino)
            if caminho_destino.is_dir():
                caminho_destino = caminho_destino / "nf_compilado.xlsx"
            
            # Configura diretórios de logs simultâneos
            configurar_pastas_logs([pasta, caminho_destino.parent])

            # Limpa logs antigos de extração
            for p_log in [pasta, caminho_destino.parent]:
                if p_log:
                    caminho_log_antigo = p_log / "log" / f"{caminho_destino.stem}_log_extracao.txt"
                    if caminho_log_antigo.exists():
                        try:
                            caminho_log_antigo.unlink()
                        except Exception:
                            pass

            pdfs = list(listar_pdfs(pasta))
            if not pdfs:
                self.window.evaluate_js(f"window.log_extracao('Nenhum PDF encontrado na pasta.', true)")
                return
            
            self.window.evaluate_js(f"window.log_extracao('Encontrados {len(pdfs)} PDFs.')")
            
            registros = []
            notas_nao_analisadas = []
            
            for indice, pdf in enumerate(pdfs, start=1):
                # Aguarda se a extração estiver pausada pelo usuário
                self._pausa_event.wait()
                
                porcentagem = int((indice / len(pdfs)) * 100)
                self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Processando PDF: {pdf.name}...')")
                try:
                    registro = extrair_nota_fiscal(
                        pdf,
                        usar_gemini=usar_gemini,
                        usar_groq=usar_groq,
                        gemini_key=gemini_key,
                        groq_key=groq_key
                    )
                    registros.append(registro)
                    registrar_incompleto_realtime(registro, caminho_destino, origem_dir=pasta)
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] {pdf.name} extraído com sucesso.')")
                except (ErroCotaGemini, ErroCotaGroq) as exc:
                    self.window.evaluate_js(f"window.log_extracao('Erro: Suas cotas de IA atingiram o limite! Por favor, tente no dia seguinte ou troque sua chave de API.', true)")
                    self.window.evaluate_js("window.sinalizar_limite_cota()")
                    
                    # Salva log de limite de cota com as notas restantes não processadas nas duas pastas
                    nao_processados = [p.name for p in pdfs[indice - 1:]]
                    for p_log in [pasta, caminho_destino.parent]:
                        if p_log:
                            caminho_log_cota = p_log / "log" / "log_limite_cota.txt"
                            try:
                                caminho_log_cota.parent.mkdir(parents=True, exist_ok=True)
                                with open(caminho_log_cota, "w", encoding="utf-8") as f:
                                    f.write("LOG DE LIMITE DE COTA DE IA EXCEDIDO\n")
                                    f.write("====================================\n")
                                    f.write("O processamento foi interrompido porque o limite de cotas (rate limit / requests limit) da API do Gemini ou do Groq foi atingido.\n")
                                    f.write("Você pode retomar a execução amanhã ou atualizar suas chaves de API nas configurações do aplicativo.\n")
                                    f.write("Os dados das notas processadas até o momento foram salvos com sucesso na planilha Excel.\n\n")
                                    f.write("As seguintes notas fiscais NÃO foram processadas:\n")
                                    for nome_pdf in nao_processados:
                                        f.write(f"- {nome_pdf}\n")
                            except Exception:
                                pass
                    break
                except ValueError as exc:
                    notas_nao_analisadas.append(pdf.name)
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}.', true)")
                except Exception as exc:
                    registros.append(NotaFiscalExtraida(arquivo_pdf=pdf.name, descricao_servico=f"ERRO NA EXTRAÇÃO: {exc}"))
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}.', true)")

                self.window.evaluate_js(f"window.update_progresso({porcentagem}, 'Processados {indice} de {len(pdfs)} PDFs')")

            if notas_nao_analisadas:
                for p_log in [pasta, caminho_destino.parent]:
                    if p_log:
                        caminho_log_na = p_log / "log" / "log_nao_analisados.txt"
                        try:
                            caminho_log_na.parent.mkdir(parents=True, exist_ok=True)
                            with open(caminho_log_na, "w", encoding="utf-8") as f:
                                f.write("Ocorreram falhas ao tentar extrair dados com as IAs (Gemini e Groq) para as seguintes notas fiscais:\n")
                                for nome_pdf in notas_nao_analisadas:
                                    f.write(f"- {nome_pdf}\n")
                        except Exception:
                            pass

            registros_completos = [r for r in registros if not r.campos_vazios()]
            gerar_xlsx(registros_completos, caminho_destino)
            registrar_log_incompletos(registros, caminho_destino, origem_dir=pasta)
            
            msg_final = f"Extração concluída. {len(registros_completos)} linhas exportadas em {caminho_destino.name}."
            self.window.evaluate_js(f"window.log_extracao('{msg_final}')")

        except Exception as e:
            tb = traceback.format_exc()
            self.window.evaluate_js(f"window.log_extracao('Erro crítico: {str(e)}', true)")
        finally:
            self._gui_log_callback = None
            try:
                self.window.evaluate_js("document.getElementById('btn-pausar-retomar').classList.add('hidden')")
            except Exception:
                pass

    def iniciar_extracao_otimizada(self, origem, destino, gemini_key, groq_key, usar_gemini, usar_groq):
        # Configurar chaves no processo
        if gemini_key:
            os.environ["GEMINI_API_KEY"] = gemini_key
        if groq_key:
            os.environ["GROQ_API_KEY"] = groq_key
        
        # Garante que a extração comece não pausada
        self._pausa_event.set()
        
        # Iniciar em thread separada para não travar a UI
        threading.Thread(
            target=self._processar_extracao_otimizada,
            args=(origem, destino, gemini_key, groq_key, usar_gemini, usar_groq),
            daemon=True
        ).start()

    def _processar_extracao_otimizada(self, origem, destino, gemini_key, groq_key, usar_gemini, usar_groq):
        # Define o callback para logs de extração na GUI
        def gui_log_callback(mensagem):
            msg_safe = message_to_js = mensagem.replace("'", "\\'").replace('"', '\\"')
            self.window.evaluate_js(f"window.log_extracao('{msg_safe}')")
        self._gui_log_callback = gui_log_callback

        try:
            pasta = Path(origem)
            if not pasta.exists() or not pasta.is_dir():
                self.window.evaluate_js(f"window.log_extracao('Erro: Pasta de origem inválida.', true)")
                return
            
            caminho_destino = Path(destino)
            if caminho_destino.is_dir():
                caminho_destino = caminho_destino / "nf_compilado.xlsx"
            
            # Configura diretórios de logs simultâneos
            configurar_pastas_logs([pasta, caminho_destino.parent])

            # Limpa logs antigos de extração
            for p_log in [pasta, caminho_destino.parent]:
                if p_log:
                    caminho_log_antigo = p_log / "log" / f"{caminho_destino.stem}_log_extracao.txt"
                    if caminho_log_antigo.exists():
                        try:
                            caminho_log_antigo.unlink()
                        except Exception:
                            pass

            pdfs = list(listar_pdfs(pasta))
            if not pdfs:
                self.window.evaluate_js(f"window.log_extracao('Nenhum PDF encontrado na pasta.', true)")
                return
            
            self.window.evaluate_js(f"window.log_extracao('Encontrados {len(pdfs)} PDFs.')")
            
            registros = []
            notas_nao_analisadas = []
            
            for indice, pdf in enumerate(pdfs, start=1):
                # Aguarda se a extração estiver pausada pelo usuário
                self._pausa_event.wait()
                
                porcentagem = int((indice / len(pdfs)) * 100)
                self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Processando PDF: {pdf.name} (Modo Otimizado - Texto)...')")
                try:
                    registro = extrair_nota_fiscal_somente_texto(
                        pdf,
                        usar_gemini=usar_gemini,
                        usar_groq=usar_groq,
                        gemini_key=gemini_key,
                        groq_key=groq_key
                    )
                    registros.append(registro)
                    registrar_incompleto_realtime(registro, caminho_destino, origem_dir=pasta)
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] {pdf.name} extraído com sucesso.')")
                except (ErroCotaGemini, ErroCotaGroq) as exc:
                    self.window.evaluate_js(f"window.log_extracao('Erro: Suas cotas de IA atingiram o limite! Por favor, tente no dia seguinte ou troque sua chave de API.', true)")
                    self.window.evaluate_js("window.sinalizar_limite_cota()")
                    
                    # Salva log de limite de cota com as notas restantes não processadas nas duas pastas
                    nao_processados = [p.name for p in pdfs[indice - 1:]]
                    for p_log in [pasta, caminho_destino.parent]:
                        if p_log:
                            caminho_log_cota = p_log / "log" / "log_limite_cota.txt"
                            try:
                                caminho_log_cota.parent.mkdir(parents=True, exist_ok=True)
                                with open(caminho_log_cota, "w", encoding="utf-8") as f:
                                    f.write("LOG DE LIMITE DE COTA DE IA EXCEDIDO\n")
                                    f.write("====================================\n")
                                    f.write("O processamento foi interrompido porque o limite de cotas (rate limit / requests limit) da API do Gemini ou do Groq foi atingido.\n")
                                    f.write("Você pode retomar a execução amanhã ou atualizar suas chaves de API nas configurações do aplicativo.\n")
                                    f.write("Os dados das notas processadas até o momento foram salvos com sucesso na planilha Excel.\n\n")
                                    f.write("As seguintes notas fiscais NÃO foram processadas:\n")
                                    for nome_pdf in nao_processados:
                                        f.write(f"- {nome_pdf}\n")
                            except Exception:
                                pass
                    break
                except ValueError as exc:
                    notas_nao_analisadas.append(pdf.name)
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}.', true)")
                except Exception as exc:
                    registros.append(NotaFiscalExtraida(arquivo_pdf=pdf.name, descricao_servico=f"ERRO NA EXTRAÇÃO: {exc}"))
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}.', true)")

                self.window.evaluate_js(f"window.update_progresso({porcentagem}, 'Processados {indice} de {len(pdfs)} PDFs')")

            if notas_nao_analisadas:
                for p_log in [pasta, caminho_destino.parent]:
                    if p_log:
                        caminho_log_na = p_log / "log" / "log_nao_analisados.txt"
                        try:
                            caminho_log_na.parent.mkdir(parents=True, exist_ok=True)
                            with open(caminho_log_na, "w", encoding="utf-8") as f:
                                f.write("Ocorreram falhas ao tentar extrair dados com as IAs (Gemini e Groq) para as seguintes notas fiscais:\n")
                                for nome_pdf in notas_nao_analisadas:
                                    f.write(f"- {nome_pdf}\n")
                        except Exception:
                            pass

            registros_completos = [r for r in registros if not r.campos_vazios()]
            gerar_xlsx(registros_completos, caminho_destino)
            registrar_log_incompletos(registros, caminho_destino, origem_dir=pasta)
            
            msg_final = f"Extração concluída (Otimizada). {len(registros_completos)} linhas exportadas em {caminho_destino.name}."
            self.window.evaluate_js(f"window.log_extracao('{msg_final}')")

        except Exception as e:
            tb = traceback.format_exc()
            self.window.evaluate_js(f"window.log_extracao('Erro crítico: {str(e)}', true)")
        finally:
            self._gui_log_callback = None
            try:
                self.window.evaluate_js("document.getElementById('btn-pausar-retomar').classList.add('hidden')")
            except Exception:
                pass

    def iniciar_automacao(self, excel_path, competencia_str):
        self._pausa_automacao_event.set()
        threading.Thread(target=self._processar_automacao, args=(excel_path, competencia_str), daemon=True).start()

    def _processar_automacao(self, excel_path, competencia_str):
        # Define o callback para logs de automação na GUI
        def gui_log_callback(mensagem):
            msg_safe = mensagem.replace("'", "\\'").replace('"', '\\"')
            self.window.evaluate_js(f"window.log_automacao('{msg_safe}')")
        self._gui_log_callback = gui_log_callback

        try:
            planilha = Path(excel_path)
            if not planilha.exists():
                self.window.evaluate_js(f"window.log_automacao('Erro: Planilha não encontrada.', true)")
                return

            # Configura diretório de log para a pasta log da planilha
            configurar_pastas_logs([planilha.parent])

            self.window.evaluate_js(f"window.log_automacao('Iniciando automação com a planilha {planilha.name}...')")

            try:
                competencia = interpretar_competencia(competencia_str)
            except Exception as e:
                self.window.evaluate_js(f"window.log_automacao('Erro na competência: {str(e)}', true)")
                return
                
            self.window.evaluate_js(f"window.log_automacao('Competência confirmada: {competencia.mes}/{competencia.ano}')")
            self.window.evaluate_js(f"window.log_automacao('Iniciando navegador Chrome (Selenium). Não feche a janela do robô!')")
            
            # Remove flag de login anterior se existir
            flag_path = Path("brain/funcao2_login_ok.flag")
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

            # Chama a execução principal do robô
            executar_automacao_iss(
                planilha, 
                competencia, 
                depuracao=False, 
                callback_progresso=callback_progresso,
                aguardar_login_por_arquivo=True,
                caminho_confirmacao_login=flag_path,
                executando_em_gui=True
            )
            self.window.evaluate_js("window.update_progresso_automacao(100, 'Escrituração concluída!')")
            self.window.evaluate_js(f"window.log_automacao('Execução do robô concluída.')")

        except Exception as e:
            tb = traceback.format_exc()
            self.window.evaluate_js(f"window.log_automacao('Erro crítico: {str(e)}', true)")
        finally:
            self._gui_log_callback = None
            try:
                self.window.evaluate_js("document.getElementById('card-confirmar-login').classList.add('hidden')")
                self.window.evaluate_js("document.getElementById('btn-pausar-retomar-automacao').classList.add('hidden')")
            except Exception:
                pass

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
        api.iniciar_extracao,
        api.iniciar_extracao_otimizada,
        api.iniciar_automacao,
        api.selecionar_pasta,
        api.selecionar_arquivo,
        api.selecionar_multiplos_xlsx,
        api.fechar_app,
        api.abrir_link,
        api.obter_chaves_salvas,
        api.salvar_chaves,
        api.contar_pdfs_origem,
        api.pausar_extracao,
        api.retomar_extracao,
        api.confirmar_login_feito,
        api.pausar_automacao,
        api.retomar_automacao,
        api.exportar_por_prefeitura,
        api.unir_planilhas_gui,
        api.obter_prefeituras_suportadas,
        api.organizar_pdfs_por_prefeitura_gui,
    )
    
    webview.start(debug=False)
