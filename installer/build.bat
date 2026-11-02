@echo off
rem ==========================================================================
rem  ExcelTools : one-click build - produces BOTH Lite and Full packages
rem
rem  RULE (2026-10-27): every installer must ship both variants.
rem    Lite : ~3 MB,  assumes .NET 10 desktop runtime already installed
rem    Full : ~63 MB, bundles windowsdesktop-runtime-10-x64.exe
rem  One .bat builds both - do not split it back into build_lite/build_full.
rem
rem  All packaging files live in <repo>\installer\  (project rule, 2026-10-27)
rem  Source of truth: <repo>\bin\Release\net10.0-windows  (see SrcDir)
rem  Output: installer\output\ExcelTools-Setup-<ver>-Lite.exe / -Full.exe
rem
rem  This is an Excel-DNA project: build it in Visual Studio (Release) first.
rem  There is no publish step here.
rem ==========================================================================
setlocal
cd /d "%~dp0"

if not exist "..\bin\Release\net10.0-windows\ExcelTools64.xll" (
  echo BUILD FAILED: ..\bin\Release\net10.0-windows\ExcelTools64.xll not found.
  echo Open the solution in Visual Studio, switch to Release, and build first.
  goto fail
)

rem Inno Setup 7 is required: compiler:ChineseSimplified.isl only exists in 7.0.2+
rem Inno Setup 6 has been uninstalled (2026-10-27) - do not add its path back.
set "ISCC=C:\Program Files\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=C:\Program Files (x86)\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" (
  echo ISCC.exe not found. Install Inno Setup 7 first.
  goto fail
)

echo [1/2] ISCC -^> Lite
"%ISCC%" "%~dp0ExcelTools.iss"
if errorlevel 1 goto fail

echo [2/2] ISCC -^> Full -DBundleRuntime
"%ISCC%" /DBundleRuntime "%~dp0ExcelTools.iss"
if errorlevel 1 goto fail

echo.
echo BUILD OK : installer\output
dir /b "%~dp0output\*.exe"
pause
exit /b 0

:fail
echo.
echo BUILD FAILED
pause
exit /b 1
