@echo off
chcp 65001 >nul
echo ========================================
echo Building dwg2rvt v2.120 Release
echo ========================================
cd /d "%~dp0"
echo Current directory: %CD%
echo.

echo Cleaning previous build...
dotnet clean dwg2rvt.sln -c Release --nologo
echo.

echo Building Release configuration...
dotnet build dwg2rvt.sln -c Release --nologo
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)
echo.
echo [SUCCESS] Build completed!
echo PostBuild deployment executed automatically via PostBuildTrigger.csproj
echo.

echo.
echo ========================================
echo DONE! Check output folder in annotatix_ver3.XXX
echo ========================================
pause
