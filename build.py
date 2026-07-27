import os
import shutil
import subprocess
import sys
import re

def gerar_chaves_padrao():
    """Lê o arquivo .env e gera um default_keys.py para embutir as chaves de forma segura no executável."""
    env_path = ".env"
    tess_api_key = ""
    tess_agent_id = ""
    
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            conteudo = f.read()
            match_key = re.search(r"^TESS_API_KEY\s*=\s*(.*)$", conteudo, re.MULTILINE)
            match_agent = re.search(r"^TESS_AGENT_ID\s*=\s*(.*)$", conteudo, re.MULTILINE)
            if match_key:
                tess_api_key = match_key.group(1).strip()
            if match_agent:
                tess_agent_id = match_agent.group(1).strip()
                
    with open("default_keys.py", "w", encoding="utf-8") as f:
        f.write("# Este arquivo é gerado temporariamente durante o build pelo script build.py\n")
        f.write("# NÃO ADICIONE SUAS CHAVES AQUI MANUALMENTE E NÃO COMITE ESTE ARQUIVO\n\n")
        f.write(f'TESS_API_KEY = "{tess_api_key}"\n')
        f.write(f'TESS_AGENT_ID = "{tess_agent_id}"\n')
    
    print(f"Gerado default_keys.py com TESS_API_KEY={'***' if tess_api_key else 'vazio'} e TESS_AGENT_ID={tess_agent_id}")

def limpar_chaves_padrao():
    """Remove o default_keys.py após a compilação para não deixar segredos no diretório do projeto."""
    if os.path.exists("default_keys.py"):
        try:
            os.remove("default_keys.py")
            print("Limpou o arquivo temporário default_keys.py com sucesso.")
        except Exception as e:
            print(f"Erro ao deletar default_keys.py: {e}")

def main():
    print("Iniciando processo de compilação do BeAContab...")
    
    # Gera o arquivo com chaves padrão do Tess para embutir no .exe
    gerar_chaves_padrao()
    
    try:
        # 1. Garante que o PyInstaller está instalado
        try:
            import PyInstaller
            print("PyInstaller já está instalado.")
        except ImportError:
            print("PyInstaller não encontrado. Instalando via pip...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            except Exception as e:
                print(f"Erro ao instalar o PyInstaller: {e}")
                sys.exit(1)

        # 2. Limpa diretórios antigos para evitar conflitos de cache
        for folder in ["build", "dist"]:
            if os.path.exists(folder):
                print(f"Limpando pasta antiga: '{folder}'...")
                try:
                    shutil.rmtree(folder)
                except Exception as e:
                    print(f"Aviso ao limpar {folder}: {e}")

        # 3. Executa a compilação do PyInstaller
        # --onedir: Gera uma pasta dist/BeAContab/ contendo o executável e dependências (melhor para Inno Setup)
        # --windowed / --noconsole: Não abre console CMD do Windows ao rodar o executável
        # --add-data: Copia a pasta 'gui' e a planilha 'cnae_oficial.xlsx' para dentro do pacote
        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--clean",
            "--noconfirm",
            "--onedir",
            "--windowed",
            "--name=automacao-iss-fortaleza",
            "--add-data=gui;gui",
            "--add-data=cnae_oficial.xlsx;.",
            "--add-data=EXEMPLOS/planilha exemplo.xlsx;EXEMPLOS",
            "--hidden-import=selenium",
            "--hidden-import=selenium.webdriver",
            "--hidden-import=selenium.webdriver.chrome.webdriver",
            "--hidden-import=selenium.webdriver.chrome.options",
            "--hidden-import=selenium.webdriver.chrome.service",
            "--hidden-import=selenium.webdriver.common.action_chains",
            "--hidden-import=selenium.webdriver.common.by",
            "--hidden-import=selenium.webdriver.common.keys",
            "--hidden-import=selenium.webdriver.remote.webdriver",
            "--hidden-import=selenium.webdriver.remote.webelement",
            "--hidden-import=selenium.webdriver.support.expected_conditions",
            "--hidden-import=selenium.webdriver.support.select",
            "--hidden-import=selenium.webdriver.support.wait",
            "--hidden-import=webdriver_manager",
            "--hidden-import=webdriver_manager.chrome",
            "app.py"
        ]
        
        print(f"Executando comando: {' '.join(cmd)}")
        try:
            subprocess.run(cmd, check=True)
            print("\n" + "="*80)
            print("COMPILAÇÃO DO PYINSTALLER CONCLUÍDA COM SUCESSO!")
            print("="*80)
            print("O aplicativo empacotado está localizado em:")
            print(f" -> {os.path.abspath('dist/automacao-iss-fortaleza')}")
            print("\nPara gerar o instalador final:")
            print(" 1. Abra o Inno Setup.")
            print(" 2. Carregue o arquivo 'setup.iss'.")
            print(" 3. Clique em 'Compile' (ou aperte F9).")
            print(" 4. O instalador gerado ficará em 'installer_output/'.")
            print("="*80)
        except subprocess.CalledProcessError as e:
            print(f"\nErro durante a execução do PyInstaller: {e}")
            sys.exit(1)
            
    finally:
        # Garante que o arquivo com as chaves seja deletado sempre, mesmo se houver erro
        limpar_chaves_padrao()

if __name__ == "__main__":
    main()
