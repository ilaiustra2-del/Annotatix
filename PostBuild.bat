@echo off
setlocal enabledelayedexpansion

REM Post-build deployment script for dwg2rvt plugin
echo ========================================
echo DWG2RVT Post-Build Deployment
echo ========================================

REM Parameters from Visual Studio/MSBuild
set PROJECT_DIR=%~1
set TARGET_DIR=%~2
set TARGET_FILE=%~3

echo Project Directory: %PROJECT_DIR%
echo Target Directory: %TARGET_DIR%
echo Target File: %TARGET_FILE%

REM Read build number
set BUILD_NUMBER_FILE=%PROJECT_DIR%BuildNumber.txt
if exist "%BUILD_NUMBER_FILE%" (
    set /p BUILD_NUMBER=<"%BUILD_NUMBER_FILE%"
) else (
    set BUILD_NUMBER=1
)

echo Build Number: %BUILD_NUMBER%

REM Format build number with leading zero (01, 02, etc.)
set FORMATTED_BUILD=%BUILD_NUMBER%
if %BUILD_NUMBER% LSS 10 set FORMATTED_BUILD=0%BUILD_NUMBER%

REM Define base output directory
set BASE_OUTPUT=C:\Users\Свеж как огурец\Desktop\Эксперимент Annotatix\dwg2rvt
set VERSION_FOLDER=dwg2rvt_ver1.%FORMATTED_BUILD%
set OUTPUT_DIR=%BASE_OUTPUT%\%VERSION_FOLDER%

echo Output Directory: %OUTPUT_DIR%

REM Create version folder
if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
    echo Created directory: %OUTPUT_DIR%
)

REM Create dwg2rvt-dll folder
set DLL_FOLDER=%OUTPUT_DIR%\dwg2rvt-dll
if not exist "%DLL_FOLDER%" (
    mkdir "%DLL_FOLDER%"
    echo Created directory: %DLL_FOLDER%
)

REM Create dwg2rvtdependencies folder
set DEPENDENCIES_FOLDER=%OUTPUT_DIR%\dwg2rvtdependencies
if not exist "%DEPENDENCIES_FOLDER%" (
    mkdir "%DEPENDENCIES_FOLDER%"
    echo Created directory: %DEPENDENCIES_FOLDER%
)

echo.
echo Copying files...
echo ----------------

REM Copy DLL and dependencies to dwg2rvt-dll folder
echo Copying to dwg2rvt-dll...
xcopy /Y /I "%TARGET_DIR%*.*" "%DLL_FOLDER%\" >nul
echo - Copied all build output to dwg2rvt-dll

REM Copy main DLL and dependencies to dwg2rvtdependencies folder
echo Copying to dwg2rvtdependencies...
xcopy /Y /I "%TARGET_DIR%dwg2rvt.dll" "%DEPENDENCIES_FOLDER%\" >nul
xcopy /Y /I "%TARGET_DIR%dwg2rvt.pdb" "%DEPENDENCIES_FOLDER%\" >nul
echo - Copied dwg2rvt.dll and dwg2rvt.pdb to dwg2rvtdependencies

REM Create .addin file
set ADDIN_FILE=%OUTPUT_DIR%\dwg2rvt.addin
echo Creating .addin file...
echo ^<?xml version="1.0" encoding="utf-8"?^> > "%ADDIN_FILE%"
echo ^<RevitAddIns^> >> "%ADDIN_FILE%"
echo   ^<AddIn Type="Application"^> >> "%ADDIN_FILE%"
echo     ^<Name^>dwg2rvt^</Name^> >> "%ADDIN_FILE%"
echo     ^<Assembly^>C:\Users\Свеж как огурец\AppData\Roaming\Autodesk\Revit\Addins\2024\dwg2rvtdependencies\dwg2rvt.dll^</Assembly^> >> "%ADDIN_FILE%"
echo     ^<ClientId^>a1b2c3d4-e5f6-7890-abcd-ef1234567890^</ClientId^> >> "%ADDIN_FILE%"
echo     ^<FullClassName^>dwg2rvt.App^</FullClassName^> >> "%ADDIN_FILE%"
echo     ^<VendorId^>DWG2RVT^</VendorId^> >> "%ADDIN_FILE%"
echo     ^<VendorDescription^>DWG to Revit Analysis Tool^</VendorDescription^> >> "%ADDIN_FILE%"
echo   ^</AddIn^> >> "%ADDIN_FILE%"
echo ^</RevitAddIns^> >> "%ADDIN_FILE%"
echo - Created dwg2rvt.addin

echo.
echo ========================================
echo Build Deployment Complete!
echo ========================================
echo Version: 1.%FORMATTED_BUILD%
echo Output: %OUTPUT_DIR%
echo.
echo Contents:
echo - dwg2rvt.addin (manifest file)
echo - dwg2rvtdependencies\ (runtime files)
echo - dwg2rvt-dll\ (complete build output)
echo ========================================

endlocal
exit /b 0
