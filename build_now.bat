@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Current directory: %CD%
echo.
echo Building solution...
dotnet build dwg2rvt.sln -c Release
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%
)
