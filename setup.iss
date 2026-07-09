; Script de Configuração do Inno Setup para o BeAContab
; Compila o executável gerado pelo PyInstaller e seus recursos em um instalador profissional do Windows.

#define MyAppName "Automação ISS Fortaleza"
#define MyAppFullName "Automação ISS Fortaleza"
#define MyAppVersion "2.19.1"
#define MyAppPublisher "Barreira & Associados"
#define MyAppExeName "automacao-iss-fortaleza.exe"

[Setup]
; AppId único para identificar esta aplicação no Windows (gerado para controle de atualização/desinstalação)
AppId={{54DA395C-DA64-4ABE-938A-2E5061AE3821}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
; Local e nome do instalador gerado
OutputDir=installer_output
OutputBaseFilename=Setup_Automacao_ISS_Fortaleza_v{#MyAppVersion}
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
Source: "dist\automacao-iss-fortaleza\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Copia todos os outros arquivos e subpastas de dependências gerados pelo PyInstaller
Source: "dist\automacao-iss-fortaleza\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Atalho no Menu Iniciar
Name: "{autoprograms}\{#MyAppFullName}"; Filename: "{app}\{#MyAppExeName}"
; Atalho na Área de Trabalho (se marcado a task desktopicon)
Name: "{autodesktop}\{#MyAppFullName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Permite inicializar o programa imediatamente após a conclusão da instalação
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppFullName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
