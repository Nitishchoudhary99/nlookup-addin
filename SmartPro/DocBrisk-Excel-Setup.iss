; ==========================================================================
;  DocBrisk for Excel - installer script (Inno Setup 6)
;
;  Builds ONE file:  Output\DocBrisk-Excel-Setup-<version>.exe
;    - Per-user install: no admin rights, no UAC prompt
;    - Installs the 32-bit or the 64-bit .xll to match the user's Excel
;    - Registers it so Excel loads it on every start (Excel 2010 to 365)
;    - Asks the user to close Excel first (Excel rewrites its add-in list on exit)
;    - Clean uninstall from Windows Settings; other add-ins are left untouched
;    - The user's licence key stays on the PC, so a reinstall keeps Pro
;
;  HOW TO BUILD
;    1. In Visual Studio build the project in Release.
;    2. Put this file in the project folder, next to SmartPro.csproj.
;    3. Check the lines marked EDIT below.
;    4. Open this file in Inno Setup and press Ctrl+F9.
; ==========================================================================

#define AppName "DocBrisk for Excel"
; EDIT: raise this number for every new build you publish
#define AppVer "3.2.0"
#define Publisher "Nitish Choudhary (DocBrisk)"
#define Site "https://docbrisk.com/addins"

; EDIT if your build folder is different. Paths start from the folder this file is in.
#define Xll64 "bin\Release\net48\publish\SmartPro-AddIn64-packed.xll"
#define Xll32 "bin\Release\net48\publish\SmartPro-AddIn-packed.xll"

[Setup]
; Do not change AppId: Windows uses it to recognise an upgrade of the same product.
AppId={{F3A6C2D4-7B1E-4E55-9C0A-2D1B6E8F4A71}
AppName={#AppName}
AppVersion={#AppVer}
AppVerName={#AppName} {#AppVer}
AppPublisher={#Publisher}
AppPublisherURL={#Site}
AppSupportURL={#Site}
AppCopyright=Copyright (C) 2026 Nitish Choudhary
VersionInfoVersion={#AppVer}
VersionInfoCompany={#Publisher}
VersionInfoDescription={#AppName} Setup
DefaultDirName={localappdata}\DocBrisk\Excel Add-in
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=DocBrisk-Excel-Setup-{#AppVer}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName}
CloseApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
FinishedLabelNoIcons=[name] is installed.%n%nOpen Excel and look for the DocBrisk tab next to Home.
FinishedLabel=[name] is installed.%n%nOpen Excel and look for the DocBrisk tab next to Home.

[InstallDelete]
; an older install of the other bitness must not stay behind
Type: files; Name: "{app}\DocBrisk-Excel-64.xll"
Type: files; Name: "{app}\DocBrisk-Excel-32.xll"

[Files]
Source: "{#Xll64}"; DestDir: "{app}"; DestName: "DocBrisk-Excel-64.xll"; Flags: ignoreversion; Check: UseExcel64
Source: "{#Xll32}"; DestDir: "{app}"; DestName: "DocBrisk-Excel-32.xll"; Flags: ignoreversion; Check: UseExcel32

[Run]
Filename: "{code:ExcelExe}"; Description: "Open Excel now"; Flags: postinstall nowait skipifsilent; Check: HasExcelExe

[UninstallDelete]
Type: dirifempty; Name: "{app}"

[Code]
const
  XLL64 = 'DocBrisk-Excel-64.xll';
  XLL32 = 'DocBrisk-Excel-32.xll';
  SCS_64BIT_BINARY = 6;

var
  Excel64: Boolean;
  ExcelPath: String;

function GetBinaryType(lpApplicationName: String; var lpBinaryType: DWORD): BOOL;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

function ExcelRunning: Boolean;
begin
  Result := FindWindowByClassName('XLMAIN') <> 0;
end;

{ Full path of EXCEL.EXE, or '' when Windows does not list it }
function FindExcelExe: String;
var
  P: String;
begin
  Result := '';
  P := '';
  if IsWin64 then
    RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', P);
  P := RemoveQuotes(P);
  if (P = '') or (not FileExists(P)) then
  begin
    P := '';
    RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', P);
    P := RemoveQuotes(P);
  end;
  if (P = '') or (not FileExists(P)) then
  begin
    P := '';
    RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', '', P);
    P := RemoveQuotes(P);
  end;
  if (P <> '') and FileExists(P) then
    Result := P;
end;

{ True when the bitness of Excel could be found. A 32-bit Excel cannot load a 64-bit .xll and the other way round. }
function DetectExcel64(var Is64: Boolean): Boolean;
var
  T: DWORD;
  Plat: String;
begin
  Result := False;
  ExcelPath := FindExcelExe;
  if ExcelPath <> '' then
  begin
    if GetBinaryType(ExcelPath, T) then
    begin
      Is64 := (T = SCS_64BIT_BINARY);
      Result := True;
      Exit;
    end;
  end;
  { Click-to-Run Office (2016 and later, Microsoft 365) records its platform here }
  Plat := '';
  if IsWin64 then
    RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Plat);
  if Plat = '' then
    RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', Plat);
  if Plat <> '' then
  begin
    Is64 := (Lowercase(Plat) = 'x64');
    Result := True;
  end;
end;

function InitializeSetup: Boolean;
var
  Rel: Cardinal;
begin
  Result := True;
  if not DetectExcel64(Excel64) then
  begin
    if not IsWin64 then
      Excel64 := False
    else
      Excel64 := MsgBox('Setup could not see which Excel you have.' + #13#10 + #13#10 +
        'In Excel open File > Account > About Excel. The first line ends with 32-bit or 64-bit.' + #13#10 + #13#10 +
        'Is your Excel 64-bit?', mbConfirmation, MB_YESNO) = IDYES;
  end;
  if (not RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Rel)) or (Rel < 528040) then
  begin
    if MsgBox('This add-in needs Microsoft .NET Framework 4.8, which was not found on this PC.' + #13#10 + #13#10 +
      'It is free from Microsoft. Install the .NET Framework 4.8 runtime, then run this setup again.' + #13#10 + #13#10 +
      'Continue anyway?', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

function UseExcel64: Boolean;
begin
  Result := Excel64;
end;

function UseExcel32: Boolean;
begin
  Result := not Excel64;
end;

function HasExcelExe: Boolean;
begin
  Result := ExcelPath <> '';
end;

function ExcelExe(Param: String): String;
begin
  Result := ExcelPath;
end;

{ Excel must be closed: it writes its add-in list when it exits and would undo the registration }
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  while ExcelRunning do
  begin
    if MsgBox('Excel is open.' + #13#10 + #13#10 + 'Save your work, close every Excel window, then click Retry.', mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := 'Setup stopped because Excel is still open. Close Excel and run setup again.';
      Exit;
    end;
  end;
end;

function OpenName(Index: Integer): String;
begin
  if Index = 0 then
    Result := 'OPEN'
  else
    Result := 'OPEN' + IntToStr(Index);
end;

function IsOurs(const Value: String): Boolean;
begin
  Result := (Pos(Lowercase(XLL64), Lowercase(Value)) > 0) or (Pos(Lowercase(XLL32), Lowercase(Value)) > 0);
end;

{ Excel reads the values OPEN, OPEN1, OPEN2 ... and stops at the first gap.
  So the list is rewritten without gaps: every other add-in stays, our old entry goes,
  and AddValue (when not empty) is put at the end. }
procedure RewriteOpenList(const OfficeVer, AddValue: String; Force: Boolean);
var
  Key, V: String;
  Keep: TArrayOfString;
  I, N, Total: Integer;
begin
  if (not Force) and (not RegKeyExists(HKCU, 'Software\Microsoft\Office\' + OfficeVer + '\Excel')) then
    Exit;
  Key := 'Software\Microsoft\Office\' + OfficeVer + '\Excel\Options';
  N := 0;
  Total := 0;
  SetArrayLength(Keep, 0);
  while RegQueryStringValue(HKCU, Key, OpenName(Total), V) do
  begin
    if not IsOurs(V) then
    begin
      SetArrayLength(Keep, N + 1);
      Keep[N] := V;
      N := N + 1;
    end;
    Total := Total + 1;
  end;
  if AddValue <> '' then
  begin
    SetArrayLength(Keep, N + 1);
    Keep[N] := AddValue;
    N := N + 1;
  end;
  for I := 0 to N - 1 do
    RegWriteStringValue(HKCU, Key, OpenName(I), Keep[I]);
  for I := N to Total do
    RegDeleteValue(HKCU, Key, OpenName(I));
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Value: String;
begin
  if CurStep = ssPostInstall then
  begin
    if Excel64 then
      Value := '/R "' + ExpandConstant('{app}') + '\' + XLL64 + '"'
    else
      Value := '/R "' + ExpandConstant('{app}') + '\' + XLL32 + '"';
    { 16.0 = Excel 2016, 2019, 2021 and Microsoft 365; 15.0 = Excel 2013; 14.0 = Excel 2010 }
    RewriteOpenList('16.0', Value, True);
    RewriteOpenList('15.0', Value, False);
    RewriteOpenList('14.0', Value, False);
  end;
end;

function InitializeUninstall: Boolean;
begin
  Result := True;
  while ExcelRunning do
  begin
    if MsgBox('Excel is open.' + #13#10 + #13#10 + 'Close every Excel window, then click Retry.', mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RewriteOpenList('16.0', '', False);
    RewriteOpenList('15.0', '', False);
    RewriteOpenList('14.0', '', False);
  end;
end;
