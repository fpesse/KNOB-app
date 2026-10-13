; Script de Inno Setup para KNOB App
[Setup]
AppName=KNOB App
AppVersion=4.0.0
DefaultDirName={pf}\KNOB App
DefaultGroupName=KNOB App
OutputBaseFilename=KNOB_App_Installer
Compression=lzma
SolidCompression=yes
SetupIconFile=E:\KNOB\icon.ico
UninstallDisplayIcon={app}\KNOB.exe
CloseApplications=force

[Files]
Source: "E:\KNOB\dist\KNOB\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "E:\KNOB\version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\KNOB App"; Filename: "{app}\KNOB.exe"
Name: "{group}\Desinstalar KNOB App"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\KNOB.exe"; Description: "Iniciar KNOB App"; Flags: nowait postinstall skipifsilent
