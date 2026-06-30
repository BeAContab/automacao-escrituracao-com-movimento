import os
import shutil
import subprocess
import sys

def main():
    print("Iniciando processo de compilação do BeAContab...")
    
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
        "--name=BeAContab",
        "--add-data=gui;gui",
        "--add-data=cnae_oficial.xlsx;.",
        "app.py"
    ]
    
    print(f"Executando comando: {' '.join(cmd)}")
    try:
        subprocess.run(cmd, check=True)
        print("\n" + "="*80)
        print("COMPILAÇÃO DO PYINSTALLER CONCLUÍDA COM SUCESSO!")
        print("="*80)
        print("O aplicativo empacotado está localizado em:")
        print(f" -> {os.path.abspath('dist/BeAContab')}")
        print("\nPara gerar o instalador final:")
        print(" 1. Abra o Inno Setup.")
        print(" 2. Carregue o arquivo 'setup.iss'.")
        print(" 3. Clique em 'Compile' (ou aperte F9).")
        print(" 4. O instalador gerado ficará em 'installer_output/'.")
        print("="*80)
    except subprocess.CalledProcessError as e:
        print(f"\nErro durante a execução do PyInstaller: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
