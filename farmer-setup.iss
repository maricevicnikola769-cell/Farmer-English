; Installer for Farmer (English) - built by GitHub Actions with Inno Setup
#ifndef AppVersion
  #define AppVersion "1.0"
#endif

[Setup]
AppId={{9D3E7A41-2C6B-4E85-B0F4-7A1C2D9E6B02}
AppName=Farmer English
AppVersion={#AppVersion}
AppPublisher=Nikola Maricevic
DefaultDirName={autopf}\Farmer English
DefaultGroupName=Farmer English
UninstallDisplayIcon={app}\FarmerEnglish.exe
SetupIconFile=farmer.ico
OutputDir=dist
OutputBaseFilename=FarmerEnglish-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DisableProgramGroupPage=yes
CloseApplications=force

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional:"

[Files]
Source: "dist\FarmerEnglish.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "farmer.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Farmer English"; Filename: "{app}\FarmerEnglish.exe"
Name: "{group}\Data (Documents\Farmer English)"; Filename: "{userdocs}\Farmer English"
Name: "{group}\Uninstall Farmer English"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Farmer English"; Filename: "{app}\FarmerEnglish.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\FarmerEnglish.exe"; Description: "Launch Farmer English"; Flags: nowait postinstall skipifsilent

[Code]
// Close Farmer English before installing, so the installer never gets stuck
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Code: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM FarmerEnglish.exe', '', SW_HIDE, ewWaitUntilTerminated, Code);
  Result := '';
end;
