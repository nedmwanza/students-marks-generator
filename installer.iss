
#define MyAppName "Students Marks Generator"
#define MyAppVersion "2.3.0"
#define MyAppPublisher "Sanex Systems, Ned Mwanza Backson Salimu"
#define MyAppExeName "StudentsMarksGenerator.exe"

[Setup]
AppId={{B37F6D25-2B45-4A61-9D2A-87C0F8B4D213}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Students Marks Generator
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=installer_output
OutputBaseFilename=StudentsMarksGeneratorSetup
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\StudentsMarksGenerator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Students Marks Generator"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Students Marks Generator"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Students Marks Generator"; Flags: postinstall nowait skipifsilent
