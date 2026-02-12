#define Year "2025"  
#define Month "12"
#define Day "26"
#define Subversion "01"
#define MyAppVersion Year + "." + Month + "." + Day + "." + Subversion 
#define MyAppName "HVACInstaller" + "_" + Year + "-" + Month + "-" + Day + ")"
#define MyAppPublisher "Рудин Андрей"
#define MyAppURL "https://vk.com/scripts_revit"

[Setup]
AppId={{335A6905-FACF-40F8-9363-937062A8D100}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={userappdata}\Autodesk\Revit\Addins\2023
DisableDirPage=yes
;DefaultGroupName={#MyAppName}
AllowNoIcons=yes
PrivilegesRequired=lowest
OutputBaseFilename={#MyAppName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Files]
Source: "ForInstallerAll\23\*"; \
    DestDir: "{userappdata}\Autodesk\Revit\Addins\2023"; \
    Flags: ignoreversion recursesubdirs createallsubdirs
    
Source: "ForInstallerAll\24\*"; \
    DestDir: "{userappdata}\Autodesk\Revit\Addins\2024"; \
    Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
; Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

