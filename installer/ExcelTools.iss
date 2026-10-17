; ============================================================================
;  ExcelTools 安装包 —— Inno Setup 6 脚本（.NET 10 版）
;
;  两种编法：
;    Lite（小，3 MB 级）： ISCC.exe ExcelTools.iss
;    Full（大，60 MB 级）： ISCC.exe /DBundleRuntime ExcelTools.iss
;    Lite：假设客户机已装 .NET 10 桌面运行时，缺了提示下载地址
;    Full：内置 windowsdesktop-runtime-10-x64.exe，缺了自动装上（弹一次 UAC）
;
;  产物：installer\output\ExcelTools-Setup-<版本>-Lite.exe / -Full.exe
;
;  安装包做的事（客户双击一次，之后永久可用）：
;    1. 拷文件到 %LOCALAPPDATA%\Programs\ExcelTools —— 免管理员、免 UAC
;    2. 自动判断客户机 Office 是 32 位还是 64 位，注册对应的那个 xll
;    3. 写「Excel 信任位置」—— 消掉"没有可用的数字签名"那个弹窗
;    4. 写「Excel 加载项 OPEN 键」—— 以后打开 Excel 自动加载，不用双击 xll
;    5. 清掉 Excel「禁用项目」里本插件的残留（曾经加载失败过就会卡在这）
;    6. 卸载时把上面写的注册表项一并删干净
;    （Full 包另有第 0 步：缺 .NET 运行时先装运行时）
;
;  关于「自包含 self-contained」：Excel-DNA 官方明确不支持
;  （github.com/excel-dna/exceldna/issues/811，作者原话），所以走的是
;  「安装包内嵌运行时安装器」这条路，客户同样不用自己去找 .NET。
;
;  注意：必须在 Excel 全部关闭的情况下安装。Excel 退出时会回写自己的 Options
;        键，可能把刚写好的 OPEN 覆盖掉 —— 脚本里有检测，检测不到就算了，
;        但别赌。
; ============================================================================

#define AppName        "ExcelTools"
#define AppVersion     "2.2.0"
#define AppPublisher   "HJH Rubber Duck"
#define AppCopyright   "© 2026 HJH Rubber Duck"
; 目标框架主版本（决定检测哪一个桌面运行时）
#define RuntimeMajor   "10"
; 相对本 .iss 所在目录的构建产物位置
#define SrcDir         "..\bin\Release\net10.0-windows"
; 内置的 .NET 桌面运行时（仅 Full 包用，放在 runtime\ 目录下）
#define RuntimeInstaller "windowsdesktop-runtime-10-x64.exe"

; 编译时加 -DBundleRuntime 出 Full 包（内置运行时），不加出 Lite 包
#ifdef BundleRuntime
  #define SetupSuffix "Full"
#else
  #define SetupSuffix "Lite"
#endif

[Setup]
AppId={{9C4E2A17-6F3B-4D58-9A21-7E5B8C3D4F60}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppCopyright={#AppCopyright}
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
AllowNoIcons=yes
; 纯每用户安装：不弹 UAC，和 HKCU 的加载项注册天然同一层
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename={#AppName}-Setup-{#AppVersion}-{#SetupSuffix}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; 只支持 64 位 Windows（32 位 Office 在 64 位 Windows 上照样能装）
ArchitecturesAllowed=x64compatible
; 升级时若 Excel 占着文件，交给 Restart Manager 处理
CloseApplications=yes
RestartApplications=no
UninstallDisplayName={#AppName}
; 安装包自身在资源管理器/任务栏里显示的图标
SetupIconFile=ExcelTools.ico
; 「应用和功能」列表里那一格图标，读的就是这个（指向装进 {app} 的 ico）
UninstallDisplayIcon={app}\ExcelTools.ico

[Languages]
Name: "cn"; MessagesFile: "lang\ChineseSimplified.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[Files]
; 64 位 Office 用这一套
Source: "{#SrcDir}\ExcelTools64.xll";              DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\ExcelTools64.dna";              DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\ExcelTools64.deps.json";        DestDir: "{app}"; Flags: ignoreversion
; 32 位 Office 用这一套（两套都拷进来，运行时按位数注册其中一套）
Source: "{#SrcDir}\ExcelTools.xll";                DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\ExcelTools.dna";                DestDir: "{app}"; Flags: ignoreversion
; 两套共用
Source: "{#SrcDir}\ExcelTools.dll";                DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\ExcelTools.deps.json";          DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\ExcelTools.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SrcDir}\runtimes\*";                    DestDir: "{app}\runtimes"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "README_客户版.txt";                        DestDir: "{app}"; DestName: "使用说明.txt"; Flags: ignoreversion isreadme
; 图标：装进安装目录，供「应用和功能」的 UninstallDisplayIcon 引用
Source: "ExcelTools.ico";                           DestDir: "{app}"; Flags: ignoreversion
; WPS 支持：预置托管宿主，避免 Excel-DNA 首次加载时在 WPS 沙箱里写文件被拒（0x80070005）
Source: "ExcelDna.ManagedHost.dll";                 DestDir: "{app}"; Flags: ignoreversion
#ifdef BundleRuntime
; 内置运行时：不复制到安装目录，只在检测到缺失时解压到临时目录运行
Source: "runtime\{#RuntimeInstaller}";              DestDir: "{tmp}"; Flags: dontcopy nocompression
#endif

; 说明文件的"查阅/打开"入口只保留 [Files] 里那条 isreadme（Inno 内置、自带中文标签）。
; 早先这里还有一条 [Run] postinstall 的"打开使用说明"，与 isreadme 在完成页重复，已删除。

[UninstallDelete]
Type: filesandordirs; Name: "{app}\runtimes"
Type: dirifempty;     Name: "{app}"

[Code]
var
  OfficeVer: String;      { 16.0 / 15.0 / 14.0 ... }
  OfficeIs64: Boolean;

{ ---------------------------------------------------------------- 基础工具 }

function BuildUtf16(const S: String): String;
var
  I: Integer;
begin
  Result := '';
  for I := 1 to Length(S) do
    Result := Result + S[I] + #0;
end;

function OfficeReg(const Sub: String): String;
begin
  Result := 'Software\Microsoft\Office\' + OfficeVer + '\Excel\' + Sub;
end;

function AddInXllPath(): String;
begin
  if OfficeIs64 then
    Result := ExpandConstant('{app}\ExcelTools64.xll')
  else
    Result := ExpandConstant('{app}\ExcelTools.xll');
end;

function OpenValueName(const I: Integer): String;
begin
  if I < 0 then
    Result := 'OPEN'
  else
    Result := 'OPEN' + IntToStr(I);
end;

{ ------------------------------------------------------------ 环境探测 }

procedure DetectOffice();
var
  Vers: array[0..4] of String;
  S: String;
  I: Integer;
begin
  Vers[0] := '16.0'; Vers[1] := '15.0'; Vers[2] := '14.0';
  Vers[3] := '12.0'; Vers[4] := '11.0';
  OfficeVer := '';
  for I := 0 to 4 do
    if RegKeyExists(HKCU, 'Software\Microsoft\Office\' + Vers[I] + '\Excel') then
    begin
      OfficeVer := Vers[I];
      Break;
    end;
  if OfficeVer = '' then
    OfficeVer := '16.0';

  { 1) 即点即用（Microsoft 365 / 2016+ 绝大多数） }
  OfficeIs64 := True;
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', S)
  or RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'Platform', S) then
  begin
    OfficeIs64 := CompareText(Trim(S), 'x64') = 0;
    Exit;
  end;

  { 2) MSI 安装：安装路径带 (x86) 就是 32 位 }
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\' + OfficeVer + '\Excel\InstallRoot', 'Path', S)
  or RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Office\' + OfficeVer + '\Excel\InstallRoot', 'Path', S) then
  begin
    OfficeIs64 := Pos('(x86)', S) = 0;
    Exit;
  end;

  { 3) 兜底：看 32 位 Program Files 下有没有 Excel }
  OfficeIs64 := not FileExists(ExpandConstant('{commonpf32}\Microsoft Office\root\Office16\EXCEL.EXE'));
end;

function OfficeRunning(): Boolean;
var
  TmpFile, Tmp: String;
  S: AnsiString;
  Code: Integer;
begin
  Result := False;
  TmpFile := ExpandConstant('{tmp}\__xltasklist.txt');
  if not Exec(ExpandConstant('{cmd}'),
              '/c tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH > "' + TmpFile + '"',
              '', SW_HIDE, ewWaitUntilTerminated, Code) then
    Exit;
  if not FileExists(TmpFile) then
    Exit;
  if LoadStringFromFile(TmpFile, S) then
  begin
    Tmp := S;
    Result := Pos('EXCEL.EXE', Uppercase(Tmp)) > 0;
  end;
  DeleteFile(TmpFile);
end;

{ 检查某个注册表视图下是否有 {#RuntimeMajor}.x 的桌面运行时。
  注意：.NET 安装器把版本号写成「值名」（如 10.0.13 = 1），不是子键；
  而且这个键固定在 WOW6432Node（32 位视图）里，跟运行时本身的位数无关。 }
function RuntimeKeyHasTarget(const Root: Integer; const Key: String): Boolean;
var
  Names: TArrayOfString;
  I: Integer;
begin
  Result := False;
  if RegGetValueNames(Root, Key, Names) then
    for I := 0 to GetArrayLength(Names) - 1 do
      if Pos('{#RuntimeMajor}.', Names[I]) = 1 then
      begin
        Result := True;
        Exit;
      end;
end;

{ .NET {#RuntimeMajor} 桌面运行时：32 位 Office 需要 x86 版，64 位需要 x64 版 }
function HasDesktopRuntime(): Boolean;
var
  Key, Arch: String;
begin
  if OfficeIs64 then
    Arch := 'x64'
  else
    Arch := 'x86';
  Key := 'SOFTWARE\dotnet\Setup\InstalledVersions\' + Arch +
         '\sharedfx\Microsoft.WindowsDesktop.App';
  Result := RuntimeKeyHasTarget(HKLM64, Key) or RuntimeKeyHasTarget(HKLM32, Key);
end;

{ ===== Full 包专用：缺运行时就把内置的装上 ===== }
#ifdef BundleRuntime
procedure EnsureDesktopRuntime();
var
  Exe: String;
  Code, I: Integer;
begin
  if HasDesktopRuntime() then
  begin
    Log('.NET {#RuntimeMajor} 桌面运行时已存在，跳过内置安装');
    Exit;
  end;

  if not OfficeIs64 then
  begin
    Log('注意：本包内置的是 x64 运行时，而客户机是 32 位 Office，需要 x86 版');
    if not WizardSilent then
      MsgBox('检测到 32 位 Office，但本安装包内置的是 x64 版 .NET 运行时，装不上。' + #13#10 + #13#10 +
             '请先在客户端安装 x86 版后重试：' + #13#10 +
             'https://dotnet.microsoft.com/download/dotnet/{#RuntimeMajor}.0', mbInformation, MB_OK);
    Exit;
  end;

  Log('未检测到 .NET {#RuntimeMajor} 桌面运行时，开始安装内置运行时（会请求管理员权限）');
  ExtractTemporaryFile('{#RuntimeInstaller}');
  Exe := ExpandConstant('{tmp}\{#RuntimeInstaller}');
  if not ShellExec('runas', Exe, '/passive /norestart', '', SW_SHOW, ewNoWait, Code) then
  begin
    Log('启动运行时安装器失败，错误码 ' + IntToStr(Code) + '（1223 = 用户在 UAC 点了取消）');
    Exit;
  end;

  { 最多等 5 分钟（.NET 安装器自带进度窗口，/passive 无需交互） }
  I := 0;
  while (I < 60) and (not HasDesktopRuntime()) do
  begin
    Sleep(5000);
    I := I + 1;
  end;

  if HasDesktopRuntime() then
    Log('运行时安装完成')
  else
    Log('运行时安装未确认完成（被取消或仍在进行）；加载项要等运行时装好才能加载');
end;
#else
procedure EnsureDesktopRuntime();
begin
  { Lite 包不内置运行时；缺失时已在 InitializeSetup 里提示下载地址 }
end;
#endif

{ ------------------------------------------------------------ 注册动作 }

{ 信任位置：消掉"没有可用的数字签名 / 仅为本会话启用"弹窗 }
procedure AddTrustedLocation();
var
  Base, P, V: String;
  I: Integer;
begin
  Base := OfficeReg('Security\Trusted Locations');
  P := AddBackslash(ExpandConstant('{app}'));

  for I := 0 to 99 do
    if RegQueryStringValue(HKCU, Base + '\Location' + IntToStr(I), 'Path', V) then
      if CompareText(AddBackslash(V), P) = 0 then
      begin
        Log('信任位置已存在：Location' + IntToStr(I));
        Exit;
      end;

  for I := 0 to 99 do
    if not RegKeyExists(HKCU, Base + '\Location' + IntToStr(I)) then
    begin
      RegWriteStringValue(HKCU, Base + '\Location' + IntToStr(I), 'Path', P);
      RegWriteDWordValue (HKCU, Base + '\Location' + IntToStr(I), 'AllowSubfolders', 1);
      RegWriteStringValue(HKCU, Base + '\Location' + IntToStr(I), 'Description', 'ExcelTools');
      Log('已添加信任位置：Location' + IntToStr(I) + ' -> ' + P);
      Exit;
    end;

  Log('警告：信任位置 0..99 已满，未能添加');
end;

{ 加载项注册：写 Options\OPEN 键，Excel 启动时自动加载 }
procedure RegisterAddIn();
var
  Base, Xll, V, Name: String;
  I: Integer;
begin
  Base := OfficeReg('Options');
  Xll  := AddInXllPath();

  for I := -1 to 99 do
  begin
    Name := OpenValueName(I);
    if RegQueryStringValue(HKCU, Base, Name, V) then
      if Pos(Lowercase(Xll), Lowercase(V)) > 0 then
      begin
        Log('加载项已注册：' + Name);
        Exit;
      end;
  end;

  for I := -1 to 99 do
  begin
    Name := OpenValueName(I);
    if not RegValueExists(HKCU, Base, Name) then
    begin
      RegWriteStringValue(HKCU, Base, Name, '"' + Xll + '"');
      Log('已注册加载项：' + Name + ' = "' + Xll + '"');
      Exit;
    end;
  end;

  Log('警告：OPEN 键 0..99 已满，未能注册加载项');
end;

{ 清掉 Excel「禁用项目」里本插件的残留（加载失败过一次就会卡在这里） }
procedure ClearDisabledItems();
var
  Base, Pat, Pat2, Tmp: String;
  Names: TArrayOfString;
  Data: AnsiString;
  I: Integer;
begin
  Base := OfficeReg('Resiliency\DisabledItems');
  if not RegGetValueNames(HKCU, Base, Names) then
    Exit;
  Pat  := BuildUtf16('ExcelTools');
  Pat2 := BuildUtf16('exceltools');
  for I := 0 to GetArrayLength(Names) - 1 do
    if RegQueryBinaryValue(HKCU, Base, Names[I], Data) then
    begin
      Tmp := Data;
      if (Pos(Pat, Tmp) > 0) or (Pos(Pat2, Tmp) > 0) then
      begin
        RegDeleteValue(HKCU, Base, Names[I]);
        Log('已清除禁用项目：' + Names[I]);
      end;
    end;
end;

{ ------------------------------------------------------------ 生命周期 }

function InitializeSetup(): Boolean;
var
  S, Bit: String;
begin
  Result := True;
  DetectOffice();
  if OfficeIs64 then
    Bit := '64 位'
  else
    Bit := '32 位';
  Log('检测到 Office 版本：' + OfficeVer + '，位数：' + Bit);

  if not HasDesktopRuntime() then
  begin
    Log('警告：未检测到 .NET {#RuntimeMajor} 桌面运行时（' + Bit + '）');
#ifdef BundleRuntime
    Log('本包为 Full 版，内置运行时安装器，将在文件复制完成后自动安装');
#else
    if not WizardSilent then
    begin
      S := '这台电脑没有检测到 .NET {#RuntimeMajor} 桌面运行时（.NET {#RuntimeMajor} Desktop Runtime）。' + #13#10 + #13#10;
      S := S + 'ExcelTools 依赖它，缺了会加载失败。' + #13#10;
      S := S + '请先安装（32 位 Office 装 x86 版，64 位 Office 装 x64 版）：' + #13#10;
      S := S + 'https://dotnet.microsoft.com/download/dotnet/{#RuntimeMajor}.0' + #13#10 + #13#10;
      S := S + '也可以改用 Full 版安装包，它会自动装好运行时。' + #13#10 + #13#10;
      S := S + '仍要继续安装吗？';
      if MsgBox(S, mbConfirmation, MB_YESNO) = IDNO then
      begin
        Result := False;
        Exit;
      end;
    end;
#endif
  end;

  if OfficeRunning() then
  begin
    Log('警告：检测到 Excel 正在运行');
    if not WizardSilent then
    begin
      S := '检测到 Excel 正在运行。' + #13#10 + #13#10;
      S := S + 'Excel 退出时会回写自己的设置，可能把刚注册的加载项覆盖掉。' + #13#10;
      S := S + '请关闭所有 Excel 窗口后再继续。' + #13#10 + #13#10;
      S := S + '已经全部关闭了就点「是」继续，否则点「否」退出安装。';
      if MsgBox(S, mbConfirmation, MB_YESNO) = IDNO then
      begin
        Result := False;
        Exit;
      end;
    end;
  end;
end;

{ ------------------------------------------------------------ WPS 支持 }

{ WPS 把 xll 加载项记在 LoadMacros：值名 = xll 完整路径，数据 = "0"（String）。
  实测自 12.1.0.29158 商业版（2026-10-17 注册表 diff 得出）。 }
procedure RegisterWPS();
var
  AppDir, Base, Sub, Path: String;
  I: Integer;
begin
  AppDir := AddBackslash(ExpandConstant('{app}'));
  Base := 'Software\kingsoft\Office\6.0\et\LoadMacros';
  { 32/64 两套 xll 都登记，WPS 32 位与 64 位都能认 }
  RegWriteStringValue(HKCU, Base, AppDir + 'ExcelTools64.xll', '0');
  RegWriteStringValue(HKCU, Base, AppDir + 'ExcelTools.xll', '0');
  { WPS 自己的信任位置（豁免沙箱对插件写文件的限制） }
  Sub := 'Software\kingsoft\Office\6.0\et\Security\Trusted Locations';
  for I := 0 to 99 do
  begin
    Path := Sub + '\Location' + IntToStr(I);
    if not RegKeyExists(HKCU, Path) then
    begin
      RegWriteStringValue(HKCU, Path, 'Path', AppDir);
      RegWriteDWordValue(HKCU, Path, 'AllowSubFolders', 1);
      break;
    end;
  end;
end;

procedure RemoveWpsEntries();
var
  Base, AppDir, P, X64, X32: String;
  I: Integer;
begin
  AppDir := AddBackslash(ExpandConstant('{app}'));
  X64 := AppDir + 'ExcelTools64.xll';
  X32 := AppDir + 'ExcelTools.xll';
  Base := 'Software\kingsoft\Office\6.0\et\LoadMacros';
  if RegKeyExists(HKCU, Base) then
  begin
    if RegValueExists(HKCU, Base, X64) then
    begin
      RegDeleteValue(HKCU, Base, X64);
      Log('已移除 WPS 加载项：ExcelTools64.xll');
    end;
    if RegValueExists(HKCU, Base, X32) then
    begin
      RegDeleteValue(HKCU, Base, X32);
      Log('已移除 WPS 加载项：ExcelTools.xll');
    end;
  end;
  Base := 'Software\kingsoft\Office\6.0\et\Security\Trusted Locations';
  for I := 0 to 99 do
    if RegQueryStringValue(HKCU, Base + '\Location' + IntToStr(I), 'Path', P) then
      if CompareText(AddBackslash(P), AppDir) = 0 then
      begin
        RegDeleteKeyIncludingSubkeys(HKCU, Base + '\Location' + IntToStr(I));
        Log('已移除 WPS 信任位置：Location' + IntToStr(I));
      end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  S: String;
begin
  if CurStep = ssPostInstall then
  begin
    EnsureDesktopRuntime();
    ClearDisabledItems();
    AddTrustedLocation();
    RegisterAddIn();
    RegisterWPS();
    if not WizardSilent then
    begin
      S := '安装完成。' + #13#10 + #13#10;
      S := S + '接下来：打开 Excel 或 WPS 表格，功能区会出现「Excel 工具」选项卡。' + #13#10;
      S := S + 'WPS 首次加载如弹出 ExcelDna.Host 提示，点确定后重启 WPS 即可。' + #13#10 + #13#10;
      S := S + '如果功能区没出现，请完全退出程序（任务管理器里确认没有 EXCEL.EXE / wps.exe）再打开。';
      MsgBox(S, mbInformation, MB_OK);
    end;
  end;
end;

{ ------------------------------------------------------------ 卸载清理 }

procedure RemoveOwnEntries();
var
  Base, P, Xll, V, Name: String;
  I: Integer;
begin
  P := AddBackslash(ExpandConstant('{app}'));

  Base := OfficeReg('Security\Trusted Locations');
  for I := 0 to 99 do
    if RegQueryStringValue(HKCU, Base + '\Location' + IntToStr(I), 'Path', V) then
      if CompareText(AddBackslash(V), P) = 0 then
      begin
        RegDeleteKeyIncludingSubkeys(HKCU, Base + '\Location' + IntToStr(I));
        Log('已移除信任位置：Location' + IntToStr(I));
      end;

  Base := OfficeReg('Options');
  Xll  := AddInXllPath();
  for I := -1 to 99 do
  begin
    Name := OpenValueName(I);
    if RegQueryStringValue(HKCU, Base, Name, V) then
      if Pos(Lowercase(Xll), Lowercase(V)) > 0 then
      begin
        RegDeleteValue(HKCU, Base, Name);
        Log('已移除加载项：' + Name);
      end;
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  DetectOffice();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RemoveOwnEntries();
    RemoveWpsEntries();
  end;
end;
