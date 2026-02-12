@echo off
echo ========================================
echo Annotatix - Add Windows Defender Exception
echo ========================================
echo.
echo This script adds an exclusion to Windows Defender to prevent
echo PluginsManager.dll from being blocked or deleted.
echo.
echo You MUST run this script as Administrator!
echo.
pause

set "REVIT_ADDINS=%APPDATA%\Autodesk\Revit\Addins\2024\annotatix_dependencies"

echo Adding Windows Defender exclusion for:
echo %REVIT_ADDINS%
echo.

powershell -Command "Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command \"Add-MpPreference -ExclusionPath ''%REVIT_ADDINS%''\"' -Verb RunAs"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Success! Folder added to exclusions.
    echo ========================================
    echo.
    echo Now you can install Annotatix without files being blocked.
) else (
    echo.
    echo [ERROR] Failed to add exclusion
    echo Please run this script as Administrator
)

echo.
pause
