@echo off
echo ========================================
echo DWG2RVT Build Script
echo ========================================
echo.

REM Increment build number and update AssemblyInfo
echo Incrementing build number...
powershell -ExecutionPolicy Bypass -Command "& '%~dp0IncrementBuildNumber.ps1'"
echo.

REM Clean previous build
echo Cleaning previous build...
dotnet clean
echo.

REM Build the project
echo Building project...
dotnet build --configuration Debug
echo.

REM Check if build succeeded
if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Build completed successfully!
echo.

REM Run post-build deployment
echo Running post-build deployment...
set PROJECT_DIR=%~dp0
set TARGET_DIR=%~dp0bin\Debug\
set TARGET_FILE=dwg2rvt.dll
powershell -ExecutionPolicy Bypass -Command "& '%~dp0PostBuild.ps1' -ProjectDir '%PROJECT_DIR%' -TargetDir '%TARGET_DIR%' -TargetFileName '%TARGET_FILE%'"

echo.
echo ========================================
echo Build and Deployment Complete!
echo ========================================
pause
