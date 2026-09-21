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
    tess_workspace_id = ""

    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            conteudo = f.read()
            match_key = re.search(r"^TESS_API_KEY\s*=\s*(.*)$", conteudo, re.MULTILINE)
            match_agent = re.search(r"^TESS_AGENT_ID\s*=\s*(.*)$", conteudo, re.MULTILINE)
            match_workspace = re.search(r"^TESS_WORKSPACE_ID\s*=\s*(.*)$", conteudo, re.MULTILINE)
            if match_key:
                tess_api_key = match_key.group(1).strip()
            if match_agent:
                tess_agent_id = match_agent.group(1).strip()
            if match_workspace:
                tess_workspace_id = match_workspace.group(1).strip()

    with open("default_keys.py", "w", encoding="utf-8") as f:
        f.write("# Este arquivo é gerado temporariamente durante o build pelo script build.py\n")
        f.write("# NÃO ADICIONE SUAS CHAVES AQUI MANUALMENTE E NÃO COMITE ESTE ARQUIVO\n\n")
        f.write(f'TESS_API_KEY = "{tess_api_key}"\n')
        f.write(f'TESS_AGENT_ID = "{tess_agent_id}"\n')
        f.write(f'TESS_WORKSPACE_ID = "{tess_workspace_id}"\n')

    print(
        f"Gerado default_keys.py com TESS_API_KEY={'***' if tess_api_key else 'vazio'}, "
        f"TESS_AGENT_ID={tess_agent_id} e TESS_WORKSPACE_ID={'***' if tess_workspace_id else 'VAZIO (a Tess será recusada pela API)'}"
    )

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))


def ler_versao_do_arquivo():
    """Lê a versão do app do arquivo VERSION (única fonte da versão)."""
    with open(os.path.join(PASTA_PROJETO, "VERSION"), "r", encoding="utf-8") as f:
        return f.read().strip()


def verificar_versao_no_changelog(versao):
    """Aborta o build se o topo do CHANGELOG.md não for a mesma versão do arquivo VERSION.

    O CHANGELOG só é lido aqui, no build — ele não vai para a instalação.
    """
    with open(os.path.join(PASTA_PROJETO, "CHANGELOG.md"), "r", encoding="utf-8") as f:
        match = re.search(r"^## \[(\d+\.\d+\.\d+)\]", f.read(), re.MULTILINE)
    versao_changelog = match.group(1) if match else "(nenhuma)"
    if versao_changelog != versao:
        print("=" * 80)
        print(f"ERRO: versão divergente — VERSION diz {versao}, mas o topo do CHANGELOG.md é {versao_changelog}.")
        print("Atualize os dois para a mesma versão antes de gerar o instalador.")
        print("=" * 80)
        sys.exit(1)


def gerar_arquivo_versao(versao):
    """Gera o _versao.py temporário com a versão embutida no executável (mesmo padrão do
    default_keys.py): o app instalado mostra esse número na tela, sem arquivo solto."""
    with open("_versao.py", "w", encoding="utf-8") as f:
        f.write("# Gerado temporariamente pelo build.py a partir do arquivo VERSION — não edite nem comite.\n")
        f.write(f'VERSAO = "{versao}"\n')
    print(f"Gerado _versao.py com VERSAO={versao}")


def limpar_arquivo_versao():
    if os.path.exists("_versao.py"):
        try:
            os.remove("_versao.py")
            print("Limpou o arquivo temporário _versao.py com sucesso.")
        except Exception as e:
            print(f"Erro ao deletar _versao.py: {e}")


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

    # A versão vem do arquivo VERSION e precisa bater com o topo do CHANGELOG.md
    versao = ler_versao_do_arquivo()
    verificar_versao_no_changelog(versao)
    print(f"Versão do app: {versao}")

    # Gera o arquivo com a versão e o de chaves padrão do Tess para embutir no .exe
    gerar_arquivo_versao(versao)
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
            "--icon=design/app_icon.ico",
            "--add-data=gui;gui",
            "--add-data=design/logo.png;design",  # só o logo — o resto de design/ é material de referência do dev, não do app
            "--add-data=design/app_icon.ico;design",
            "--add-data=core;core",
            "--add-data=cnae_oficial.xlsx;.",
            "--add-data=ibge_municipios.json;.",
            "--add-data=EXEMPLOS/planilha exemplo.xlsx;EXEMPLOS",
            "--hidden-import=pydantic",
            "--hidden-import=loguru",
            "--hidden-import=python_spreadsheet_reader",
            "--hidden-import=python_spreadsheet_reader.readers",
            "--hidden-import=python_spreadsheet_reader.readers.xlsx",
            "--hidden-import=python_spreadsheet_reader.styler",
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
            "--hidden-import=cryptography",
            "--hidden-import=cryptography.hazmat.primitives.serialization.pkcs12",
            "--hidden-import=cryptography.hazmat.backends.openssl",
            "--hidden-import=requests",
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
        # Garante que os arquivos temporários (chaves e versão) sejam deletados sempre, mesmo se houver erro
        limpar_chaves_padrao()
        limpar_arquivo_versao()

if __name__ == "__main__":
    main()
