@echo off
chcp 65001 >nul
echo ========================================
echo Annotatix Quick Build
echo ========================================
cd /d "%~dp0"

REM Increment build number
echo Incrementing version...
for /f "usebackq delims=" %%a in ("BuildNumber.txt") do set BUILD_NUM=%%a
set /a NEW_BUILD=%BUILD_NUM%+1
echo %NEW_BUILD%> BuildNumber.txt
echo Version: 3.%NEW_BUILD%
echo.

REM Build solution
echo Building solution...
dotnet msbuild dwg2rvt.sln /p:Configuration=Release /t:Rebuild /v:minimal

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Build SUCCESSFUL! Version 3.%NEW_BUILD%
    echo ========================================
    echo.
    echo Check output folder for annotatix_ver3.%NEW_BUILD%
) else (
    echo.
    echo ========================================
    echo Build FAILED!
    echo ========================================
)

pause
