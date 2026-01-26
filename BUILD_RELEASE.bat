@echo off
chcp 65001 >nul
echo Building dwg2rvt Release version...
cd /d "%~dp0"
dotnet build -c Release
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo Running PostBuild script...
    powershell -ExecutionPolicy Bypass -File "%~dp0PostBuild.ps1" -ProjectDir "%~dp0" -TargetDir "%~dp0bin\Release\" -TargetFileName "dwg2rvt.dll"
    echo.
    echo Done! Check the output folder.
    pause
) else (
    echo.
    echo Build failed!
    pause
    exit /b 1
)
