@echo off
setlocal enabledelayedexpansion

REM Build number management script for dwg2rvt plugin
echo ========================================
echo DWG2RVT Build Number Manager
echo ========================================

REM Path to build number file
set BUILD_NUMBER_FILE=%~dp0BuildNumber.txt

REM Read current build number
if exist "%BUILD_NUMBER_FILE%" (
    set /p BUILD_NUMBER=<"%BUILD_NUMBER_FILE%"
) else (
    set BUILD_NUMBER=1
)

echo Current build number: %BUILD_NUMBER%

REM Increment build number
set /a NEW_BUILD_NUMBER=%BUILD_NUMBER%+1
echo !NEW_BUILD_NUMBER!> "%BUILD_NUMBER_FILE%"

echo New build number: !NEW_BUILD_NUMBER!
echo Build number updated successfully.
echo ========================================

endlocal
exit /b 0
