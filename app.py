import os
import sys
import threading
from pathlib import Path
import traceback

import webview

# Importar lógicas do projeto
from extrair_nf_pdfs import listar_pdfs, extrair_nota_fiscal, gerar_xlsx, registrar_log_incompletos, NotaFiscalExtraida
from gemini_extracao import ErroCotaGemini
from iss_fortaleza_automacao import executar_automacao_iss, solicitar_competencia, interpretar_competencia
from tratamento_erros import registrar_erro, registrar_evento_execucao

class BeAContabAPI:
    def __init__(self, window):
        self.window = window

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

    def abrir_link(self, url):
        import webbrowser
        webbrowser.open(url)
        return True

    def fechar_app(self):
        self.window.destroy()

    def iniciar_extracao(self, origem, destino, api_key):
        # Configurar chave
        os.environ["GEMINI_API_KEY"] = api_key
        
        # Iniciar em thread separada para não travar a UI
        threading.Thread(target=self._processar_extracao, args=(origem, destino, api_key), daemon=True).start()

    def _processar_extracao(self, origem, destino, api_key):
        try:
            pasta = Path(origem)
            if not pasta.exists() or not pasta.is_dir():
                self.window.evaluate_js(f"window.log_extracao('Erro: Pasta de origem inválida.', true)")
                return
            
            caminho_destino = Path(destino)
            if caminho_destino.is_dir():
                caminho_destino = caminho_destino / "nf_compilado.xlsx"
            
            pdfs = list(listar_pdfs(pasta))
            if not pdfs:
                self.window.evaluate_js(f"window.log_extracao('Nenhum PDF encontrado na pasta.', true)")
                return
            
            self.window.evaluate_js(f"window.log_extracao('Encontrados {len(pdfs)} PDFs.')")
            
            registros = []
            notas_com_erro_cota = []
            
            for indice, pdf in enumerate(pdfs, start=1):
                self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Processando PDF: {pdf.name}...')")
                try:
                    registro = extrair_nota_fiscal(pdf, api_key)
                    registros.append(registro)
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] {pdf.name} extraído com sucesso.')")
                except ErroCotaGemini as exc:
                    notas_com_erro_cota.append((pdf.name, str(exc)))
                    registros.append(NotaFiscalExtraida(arquivo_pdf=pdf.name, descricao_servico="NÃO ANALISADO: Limite de cota da IA excedido."))
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Cota excedida em {pdf.name}.', true)")
                except Exception as exc:
                    registros.append(NotaFiscalExtraida(arquivo_pdf=pdf.name, descricao_servico=f"ERRO NA EXTRAÇÃO: {exc}"))
                    self.window.evaluate_js(f"window.log_extracao('[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}.', true)")

            registros_completos = [r for r in registros if not r.campos_vazios()]
            gerar_xlsx(registros_completos, caminho_destino)
            registrar_log_incompletos(registros, caminho_destino)
            
            msg_final = f"Extração concluída. {len(registros_completos)} linhas exportadas em {caminho_destino.name}."
            self.window.evaluate_js(f"window.log_extracao('{msg_final}')")

        except Exception as e:
            tb = traceback.format_exc()
            self.window.evaluate_js(f"window.log_extracao('Erro crítico: {str(e)}', true)")

    def iniciar_automacao(self, excel_path, api_key, competencia_str):
        if api_key:
            os.environ["GEMINI_API_KEY"] = api_key
            
        threading.Thread(target=self._processar_automacao, args=(excel_path, competencia_str), daemon=True).start()

    def _processar_automacao(self, excel_path, competencia_str):
        try:
            planilha = Path(excel_path)
            if not planilha.exists():
                self.window.evaluate_js(f"window.log_automacao('Erro: Planilha não encontrada.', true)")
                return

            self.window.evaluate_js(f"window.log_automacao('Iniciando automação com a planilha {planilha.name}...')")

            try:
                competencia = interpretar_competencia(competencia_str)
            except Exception as e:
                self.window.evaluate_js(f"window.log_automacao('Erro na competência: {str(e)}', true)")
                return
                
            self.window.evaluate_js(f"window.log_automacao('Competência confirmada: {competencia.mes}/{competencia.ano}')")
            self.window.evaluate_js(f"window.log_automacao('Iniciando navegador Playwright. Não feche a janela do robô!')")
            
            # Chama a execução principal do robô
            executar_automacao_iss(planilha, competencia, depuracao=True)
            self.window.evaluate_js(f"window.log_automacao('Execução do robô concluída.')")

        except Exception as e:
            tb = traceback.format_exc()
            self.window.evaluate_js(f"window.log_automacao('Erro crítico: {str(e)}', true)")

if __name__ == '__main__':
    # Cria a janela do pywebview
    # A interface gráfica fica em gui/index.html
    html_path = str((Path(__file__).parent / "gui" / "index.html").resolve())
    
    window = webview.create_window('BeAContab - Automação Fiscal', html_path, width=1280, height=800)
    api = BeAContabAPI(window)
    window.expose(api.iniciar_extracao, api.iniciar_automacao, api.selecionar_pasta, api.selecionar_arquivo, api.fechar_app, api.abrir_link)
    
    webview.start(debug=True)
