[Setup]
AppId={{5C0E4A4A-1234-4F6C-A750-0C2AD1B8B0E8}
AppName=KNOB Audio Mixer
AppVersion=1.0.0
AppPublisher=KNOB
AppPublisherURL=https://example.com
AppSupportURL=https://example.com/support
AppUpdatesURL=https://example.com/downloads
DefaultDirName={autopf}\KNOB\Audio Mixer
DisableDirPage=no
DefaultGroupName=KNOB
AllowNoIcons=no
OutputDir=installer\output
OutputBaseFilename=KNOB_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=..\assets\knob.ico
UninstallDisplayIcon={app}\audio_mixer_app.exe
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el &Escritorio"; GroupDescription: "Opciones adicionales:"; Flags: unchecked
Name: "startmenu"; Description: "Crear acceso directo en el menú Inicio"; GroupDescription: "Opciones adicionales:"; Flags: unchecked

[Files]
Source: "..\dist\audio_mixer_app.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\KNOB"; Filename: "{app}\audio_mixer_app.exe"; Tasks: startmenu
Name: "{commondesktop}\KNOB"; Filename: "{app}\audio_mixer_app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\audio_mixer_app.exe"; Description: "Ejecutar KNOB"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
