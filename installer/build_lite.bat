@echo off
rem Build LITE package (~3MB, assumes .NET 10 desktop runtime already installed)
"C:\Users\huangjianhua\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "%~dp0ExcelTools.iss"
pause
