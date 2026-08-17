#define AppName "Consumers Energy to Excel"
#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif
[Setup]
UninstallDisplayIcon={app}\app-icon.ico
SetupIconFile=app-icon.ico
AppId={{945A0AB3-BEF9-4182-B26F-3AB0EAA72B60}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\Consumers Energy to Excel
DefaultGroupName={#AppName}
OutputDir=..\..\artifacts
OutputBaseFilename=ConsumersEnergyToExcel-{#AppVersion}-windows-x64-setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
[Files]
Source: "app-icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\dist\consumers-energy-to-excel.exe"; DestDir: "{app}"; Flags: ignoreversion
[Icons]
Name: "{group}\Consumers Energy to Excel"; Filename: "{app}\consumers-energy-to-excel.exe"; IconFilename: "{app}\app-icon.ico"
