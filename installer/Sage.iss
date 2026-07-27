; Inno Setup script for the Sage installer
; Requires the app to be published first with:
;   dotnet publish ..\src\Sage\Sage.vbproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

#define MyAppName "Sage"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Yunier"
#define MyAppExeName "Sage.exe"
#define PublishDir "..\src\Sage\bin\Release\net8.0-windows\win-x64\publish"

[Setup]
AppId={{B3B9B8B0-6C1E-4E36-9C2A-2B7B7F0B1A11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=SageSetup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Dirs]
Name: "{app}\Plantillas"; Permissions: users-modify

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Run {#MyAppName}"; Flags: nowait postinstall skipifsilent
