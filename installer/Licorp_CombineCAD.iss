#define MyAppName "Licorp CombineCAD"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Licorp"
#define MyAppExeName "Licorp_CombineCAD_Setup_1.0.0.exe"
#ifndef MyArtifactRoot
  #define MyArtifactRoot AddBackslash(SourcePath) + "..\artifacts\release\1.0.0"
#endif
#ifndef MyStagingRoot
  #define MyStagingRoot MyArtifactRoot + "\staging"
#endif

[Setup]
AppId={{F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppPublisher}\{#MyAppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir={#MyArtifactRoot}\installer
OutputBaseFilename=Licorp_CombineCAD_Setup_{#MyAppVersion}
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#MyStagingRoot}\revit\R2020\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2020"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2021\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2021"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2022\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2022"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2023\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2023"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2024\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2024"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2025\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2025"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2026\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2026"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\R2027\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD\R2027"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MyStagingRoot}\revit\Addins\2020\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2020"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2021\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2021"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2022\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2022"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2023\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2023"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2024\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2024"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2025\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2025"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2026\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2026"; Flags: ignoreversion
Source: "{#MyStagingRoot}\revit\Addins\2027\Licorp_CombineCAD.addin"; DestDir: "{commonappdata}\Autodesk\Revit\Addins\2027"; Flags: ignoreversion
Source: "{#MyStagingRoot}\autocad\*"; DestDir: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]

[Run]

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_CombineCAD"
Type: filesandordirs; Name: "{commonappdata}\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2020\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2021\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2022\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2023\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2024\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2025\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2026\Licorp_CombineCAD.addin"
Type: files; Name: "{commonappdata}\Autodesk\Revit\Addins\2027\Licorp_CombineCAD.addin"
