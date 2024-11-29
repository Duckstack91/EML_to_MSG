; EML_to_MSG Installer Script
; Dieses Skript erstellt ein Installationspaket für die Anwendung.

[Setup]
AppName=EML_to_MSG
AppVersion=1.3
DefaultDirName={autopf}\EML_to_MSG
DefaultGroupName=EML_to_MSG
UninstallDisplayIcon={app}\EML_to_MSG.exe
Compression=lzma2
SolidCompression=yes
OutputDir={userdocs}\EML_to_MSG_Installer
WizardStyle=modern

[Dirs]
Name: "{app}\src\icons"
Name: "{app}\src\"
Name: "{app}\src\logs"

[Files]
; Hauptanwendungsdatei
Source: "Eml_to_Msg.exe"; DestDir: "{app}"; Flags: ignoreversion
; Icon-Datei (falls benötigt)
Source: "EML_to_MSGIcon.ico"; DestDir: "{app}\icons"; Flags: onlyifdoesntexist
; Konfigurationsdatei
Source "Version.txt"
; README-Datei
Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme

[Icons]
; Startmenü-Verknüpfung
Name: "{group}\EML_to_MSG"; Filename: "{app}\Eml_to_Msg.exe"; IconFilename: "{app}\icons\EML_to_MSGIcon.ico"
; Desktop-Verknüpfung
Name: "{commondesktop}\EML_to_MSG"; Filename: "{app}\Eml_to_Msg.exe"; IconFilename: "{app}\icons\EML_to_MSGIcon.ico"


[Run]
; Startet die Anwendung nach der Installation (optional)
Filename: "{app}\Eml_to_Msg.exe"; Description: "{cm:LaunchProgram,Eml_to_Msg}"; Flags: nowait postinstall skipifsilent
