@echo off
setlocal

set "CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
set "FW=C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
set "WINMD=C:\Windows\System32\WinMetadata"
set "SCRIPT_DIR=%~dp0"
set "OUT_EXE=%SCRIPT_DIR%..\Print Selector.exe"

echo Compiling Print Selector.exe...
"%CSC%" /nologo /target:winexe /optimize+ ^
    /win32icon:"%SCRIPT_DIR%app_icon.ico" ^
    /r:"%FW%\System.dll" ^
    /r:"%FW%\System.Core.dll" ^
    /r:"%FW%\System.Drawing.dll" ^
    /r:"%FW%\System.Windows.Forms.dll" ^
    /r:"%FW%\System.Web.Extensions.dll" ^
    /r:"%FW%\Microsoft.CSharp.dll" ^
    /r:"%FW%\System.Runtime.dll" ^
    /r:"%FW%\System.Runtime.WindowsRuntime.dll" ^
    /r:"%FW%\System.Runtime.InteropServices.WindowsRuntime.dll" ^
    /r:"%WINMD%\Windows.Data.winmd" ^
    /r:"%WINMD%\Windows.Storage.winmd" ^
    /r:"%WINMD%\Windows.Foundation.winmd" ^
    /out:"%OUT_EXE%" ^
    "%SCRIPT_DIR%Program.cs" ^
    "%SCRIPT_DIR%MainForm.cs" ^
    "%SCRIPT_DIR%PrinterEngine.cs" ^
    "%SCRIPT_DIR%Converter.cs" ^
    "%SCRIPT_DIR%Translations.cs" ^
    "%SCRIPT_DIR%CrashLogger.cs"

if %ERRORLEVEL% EQU 0 (
    echo Build succeeded: "%OUT_EXE%"
) else (
    echo Build failed with exit code %ERRORLEVEL%.
    exit /b %ERRORLEVEL%
)
