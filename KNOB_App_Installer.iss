; Script de Inno Setup para KNOB App
[Setup]
AppName=KNOB App
AppVersion=3.0.0
DefaultDirName={pf}\KNOB App
DefaultGroupName=KNOB App
OutputBaseFilename=KNOB_App_Installer
Compression=lzma
SolidCompression=yes
SetupIconFile=E:\KNOB\icon.ico
UninstallDisplayIcon={app}\KNOB.exe

[Files]
Source: "E:\KNOB\dist\KNOB.exe"; DestDir: "{app}"; Flags: ignoreversion
; Incluye otros archivos necesarios aquí si los hay

[Icons]
Name: "{group}\KNOB App"; Filename: "{app}\KNOB.exe"
Name: "{group}\Desinstalar KNOB App"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\KNOB.exe"; Description: "Iniciar KNOB App"; Flags: nowait postinstall skipifsilent
