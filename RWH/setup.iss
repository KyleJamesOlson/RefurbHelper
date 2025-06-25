#define MyAppName "Spectrum E-cycle Refurb Worksheet Helper"
#define MyAppVersion "4.2.0"
#define MyAppPublisher "Spectrum E-cycle" ; Replace with your company or your name
#define MyAppExeName "RefurbHelper.exe"
#define MyAppSourcePath "dist" ; Path to the PyInstaller output folder
#define MyAppIcon "assets\icon.ico" ; Path to the icon file relative to the script
#define ReadmeFile "README.DOC" ; Name of the README file

[Setup]
AppId={{123e4567-e89b-12d3-a456-426614174000}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=Setup_{#MyAppName}_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
SetupIconFile={#MyAppIcon}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "openreadme"; Description: "Open README after installation"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#MyAppSourcePath}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#MyAppSourcePath}\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyAppSourcePath}\Template.docx"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#ReadmeFile}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#ReadmeFile}"; Description: "Open README"; Flags: postinstall skipifsilent shellexec; Tasks: openreadme

[UninstallDelete]
Type: filesandordirs; Name: "{app}\module_cache"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    if DirExists(ExpandConstant('{app}\module_cache')) then
      DelTree(ExpandConstant('{app}\module_cache'), True, True, True);
  end;
end;