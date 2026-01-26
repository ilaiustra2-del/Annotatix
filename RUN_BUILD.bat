@echo off
chcp 65001 >nul
echo ========================================
echo Building dwg2rvt v2.120 Release
echo ========================================
cd /d "%~dp0"
echo Current directory: %CD%
echo.

echo Cleaning previous build...
dotnet clean -c Release --nologo
echo.

echo Building Release configuration...
dotnet build -c Release --nologo
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)
echo.
echo [SUCCESS] Build completed!
echo.

echo Running PostBuild script...
powershell -ExecutionPolicy Bypass -File "%~dp0PostBuild.ps1" -ProjectDir "%~dp0" -TargetDir "%~dp0bin\Release\" -TargetFileName "dwg2rvt.dll"

echo.
echo ========================================
echo DONE! Check output folder:
echo %~dp0..\dwg2rvt_ver2.120
echo ========================================
pause
