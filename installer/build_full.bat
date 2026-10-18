@echo off
rem Build FULL package (bundles .NET 10 desktop runtime, ~63MB)
rem In this cmd context the switch must be single-slash /DBundleRuntime
rem (double-slash //D is only needed in Git Bash because of MSYS path conversion)
"C:\Users\huangjianhua\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DBundleRuntime "%~dp0ExcelTools.iss"
pause
