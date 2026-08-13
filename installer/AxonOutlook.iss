; Axon Outlook add-in — standalone installer (per-user, no admin).
; Installs ONLY the Outlook add-in (no floating dot). Registers the managed COM add-in and writes its
; config (Mistral primary, OpenAI backup). For colleagues who just want the email features.
;
; Build:  ISCC.exe /DMistralKey=... /DOpenAIKey=... installer\AxonOutlook.iss
; Output: installer\Output\Axon-Outlook-Setup.exe

#define AppName     "Axon Outlook add-in"
#define AppVersion  "1.0.0.5"
#define AppPublisher "Axon Group"
#define AddinProgId "Axon.OutlookAddin"
; Keys for the add-in config.json, passed at build time (empty by default so the committed script
; holds no secret). build.ps1 reads them from assistant\.env. Standalone add-in has no dot .env to
; read from, so BOTH keys are baked directly into config.json.
#ifndef MistralKey
  #define MistralKey ""
#endif
#ifndef OpenAIKey
  #define OpenAIKey ""
#endif
#define AddinDir  "..\outlook_addin"

[Setup]
AppId={{B2D8F1A3-4E77-4C9B-A1E6-8D3F2B9C7E52}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Machine-wide: C:\Program Files\Axon\Disassist, registered under HKLM so it loads for EVERY user.
DefaultDirName={commonpf}\Axon\Disassist
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=admin
OutputDir=Output
OutputBaseFilename=Axon-Outlook-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=axon.ico

[Files]
; The add-in DLL + its ribbon icons must sit in the same folder (the add-in loads icons from its
; own DLL directory).
Source: "{#AddinDir}\AxonAddin.dll";      DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-move.png";      DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-download.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-summarize.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-reply.png";     DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-schedule.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-followup.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-sendlater.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-write.png";     DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-attach.png";    DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-settings.png";  DestDir: "{app}"; Flags: ignoreversion

[UninstallDelete]
; WriteAddinConfig creates config.json at run time, so Setup doesn't track it and uninstall would
; otherwise leave the API keys sitting in Program Files. Remove it, then the now-empty folders.
Type: files;     Name: "{app}\config.json"
Type: dirifempty; Name: "{app}"
Type: dirifempty; Name: "{commonpf}\Axon"

[Code]
const
  CLSID  = '{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}';
  PROGID = 'Axon.OutlookAddin';
  ASM    = 'AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';

{ Register the managed COM Outlook add-in MACHINE-WIDE (HKLM) in ONE registry view. The AnyCPU DLL
  runs in either bitness; only the view and the CLR loader differ. Called twice: the 64-bit view with
  the 64-bit mscoree (for 64-bit Outlook), and the 32-bit view with SysWOW64\mscoree (for 32-bit
  Outlook). Plain strings avoid the [Registry] GUID-brace parsing problem. }
procedure RegisterAddinView(root: Integer; mscoree: String);
var
  code, clsKey, inproc, addins: String;
begin
  code := ExpandConstant('{app}\AxonAddin.dll');
  StringChangeEx(code, '\', '/', True);
  code := 'file:///' + code;

  RegWriteStringValue(root, 'Software\Classes\' + PROGID + '\CLSID', '', CLSID);

  clsKey := 'Software\Classes\CLSID\' + CLSID;
  inproc := clsKey + '\InprocServer32';
  RegWriteStringValue(root, inproc, '', mscoree);
  RegWriteStringValue(root, inproc, 'ThreadingModel', 'Both');
  RegWriteStringValue(root, inproc, 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(root, inproc, 'Assembly', ASM);
  RegWriteStringValue(root, inproc, 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(root, inproc, 'CodeBase', code);
  RegWriteStringValue(root, inproc + '\0.0.0.0', 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(root, inproc + '\0.0.0.0', 'Assembly', ASM);
  RegWriteStringValue(root, inproc + '\0.0.0.0', 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(root, inproc + '\0.0.0.0', 'CodeBase', code);
  RegWriteStringValue(root, clsKey + '\ProgId', '', PROGID);

  addins := 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID;
  RegWriteStringValue(root, addins, 'FriendlyName', 'Axon intelligence');
  RegWriteStringValue(root, addins, 'Description', 'Summarize, reply, schedule, file and download emails with Axon');
  RegWriteDWordValue(root, addins, 'LoadBehavior', 3);
end;

{ Register in BOTH views so the same installer works on 32-bit and 64-bit Office. }
procedure RegisterAddin;
begin
  RegisterAddinView(HKLM64, ExpandConstant('{win}\System32\mscoree.dll'));
  RegisterAddinView(HKLM32, ExpandConstant('{win}\SysWOW64\mscoree.dll'));
end;

{ Write the add-in config with both keys baked, NEXT TO THE DLL in the install folder. A machine-wide
  install cannot write into every user's %APPDATA%, so the keys live beside the assembly and serve all
  users. A user who wants their own model server can still place a config.json in %APPDATA%\AxonOutlook,
  which wins. }
procedure WriteAddinConfig;
var
  path, json, key, okey: String;
begin
  key := '{#MistralKey}';
  okey := '{#OpenAIKey}';
  path := ExpandConstant('{app}\config.json');
  if key = '' then
    json := '{"api_base": "https://api.openai.com/v1", "api_key": "' + okey + '", "model": "gpt-4o"}'
  else
    json := '{"api_base": "https://api.mistral.ai/v1", "api_key": "' + key + '", "model": "mistral-medium-latest",' +
            ' "backup_api_base": "https://api.openai.com/v1", "backup_api_key": "' + okey + '", "backup_model": "gpt-4o"}';
  SaveStringToFile(path, json, False);
end;

{ Earlier builds installed PER-USER (HKCU + %LOCALAPPDATA%\AxonOutlook). HKCU COM registration takes
  precedence over HKLM, so a leftover per-user copy would shadow this machine-wide one and the user
  would keep running the old add-in. Scrub it. }
procedure RemoveOldPerUserInstall;
var
  old: String;
begin
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\' + CLSID);
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\' + PROGID);
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  old := ExpandConstant('{localappdata}\AxonOutlook');
  if DirExists(old) then DelTree(old, True, True, True);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RemoveOldPerUserInstall;
    RegisterAddin;
    WriteAddinConfig;
  end;
end;

{ Close Outlook before we remove files. It holds AxonAddin.dll and keeps handles in the per-user data
  folder, so without this the DLL and the residue get left behind. Ask nicely, give it a moment to
  save, then force. }
procedure CloseOutlook;
var rc: Integer;
begin
  Exec('taskkill.exe', '/im OUTLOOK.EXE', '', SW_HIDE, ewWaitUntilTerminated, rc);
  Sleep(3000);
  Exec('taskkill.exe', '/f /im OUTLOOK.EXE', '', SW_HIDE, ewWaitUntilTerminated, rc);
end;

{ Remove the per-user runtime data (%APPDATA%\AxonOutlook: learned tone, archive memory, reminders)
  from EVERY user profile — the machine install never tracked it, so it is the residue left behind. }
procedure CleanResidueAllUsers;
var
  rec: TFindRec;
  users, folder: String;
begin
  users := ExpandConstant('{sd}\Users');
  if FindFirst(users + '\*', rec) then
  try
    repeat
      if (rec.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0) and (rec.Name <> '.') and (rec.Name <> '..') then
      begin
        folder := users + '\' + rec.Name + '\AppData\Roaming\AxonOutlook';
        if DirExists(folder) then DelTree(folder, True, True, True);
      end;
    until not FindNext(rec);
  finally
    FindClose(rec);
  end;
end;

{ Uninstall runs elevated (admin), so it can reach HKLM and every user's profile. }
function InitializeUninstall(): Boolean;
begin
  CloseOutlook;
  Result := True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    { Remove BOTH registry views (32-bit and 64-bit), since we register in both. }
    RegDeleteKeyIncludingSubkeys(HKLM64, 'Software\Classes\CLSID\' + CLSID);
    RegDeleteKeyIncludingSubkeys(HKLM64, 'Software\Classes\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKLM64, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKLM32, 'Software\Classes\CLSID\' + CLSID);
    RegDeleteKeyIncludingSubkeys(HKLM32, 'Software\Classes\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKLM32, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  end;
  if CurUninstallStep = usPostUninstall then
    CleanResidueAllUsers;
end;
