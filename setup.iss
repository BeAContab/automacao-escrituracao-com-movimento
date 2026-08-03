; Script de Configuração do Inno Setup para o BeAContab
; Compila o executável gerado pelo PyInstaller e seus recursos em um instalador profissional do Windows.

[Setup]
; AppId único para identificar esta aplicação no Windows (gerado para controle de atualização/desinstalação)
AppId={{54DA395C-DA64-4ABE-938A-2E5061AE3821}
AppName=Automação ISS Fortaleza
AppVersion=2.22.0
AppPublisher=Barreira & Associados
DefaultDirName={autopf}\Automação ISS Fortaleza
DisableProgramGroupPage=yes
; Local e nome do instalador gerado
OutputDir=installer_output
OutputBaseFilename=Setup_Automacao_ISS_Fortaleza_v2.22.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
; Define o idioma padrão do instalador como Português do Brasil
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
; Cria a opção de criar atalho na Área de Trabalho
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copia o executável principal
Source: "dist\automacao-iss-fortaleza\automacao-iss-fortaleza.exe"; DestDir: "{app}"; Flags: ignoreversion
; Copia todos os outros arquivos e subpastas de dependências gerados pelo PyInstaller
Source: "dist\automacao-iss-fortaleza\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Atalho no Menu Iniciar
Name: "{autoprograms}\Automação ISS Fortaleza"; Filename: "{app}\automacao-iss-fortaleza.exe"
; Atalho na Área de Trabalho (se marcado a task desktopicon)
Name: "{autodesktop}\Automação ISS Fortaleza"; Filename: "{app}\automacao-iss-fortaleza.exe"; Tasks: desktopicon

[Run]
; Permite inicializar o programa imediatamente após a conclusão da instalação
Filename: "{app}\automacao-iss-fortaleza.exe"; Description: "{cm:LaunchProgram,Automação ISS Fortaleza}"; Flags: nowait postinstall skipifsilent

