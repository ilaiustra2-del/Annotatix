@echo off
chcp 65001 >nul
echo ========================================
echo Annotatix - Создание бекапа
echo ========================================
echo.

set "REVIT_ADDINS=%APPDATA%\Autodesk\Revit\Addins\2024"
set "SOURCE_DIR=%REVIT_ADDINS%\annotatix_dependencies"

REM Check if installed
if not exist "%SOURCE_DIR%\main\PluginsManager.dll" (
    echo [ERROR] Annotatix не установлен в Revit Addins
    pause
    exit /b 1
)

REM Get current version
if exist "%SOURCE_DIR%\main\BuildNumber.txt" (
    set /p BUILD_NUMBER=<"%SOURCE_DIR%\main\BuildNumber.txt"
) else (
    set BUILD_NUMBER=unknown
)

echo Текущая установленная версия: 3.%BUILD_NUMBER%
echo.

REM Create backup folder name with timestamp
set TIMESTAMP=%DATE:~-4%%DATE:~3,2%%DATE:~0,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%
set TIMESTAMP=%TIMESTAMP: =0%
set "BACKUP_DIR=%~dp0..\backup_annotatix_3.%BUILD_NUMBER%_%TIMESTAMP%"

echo Создание бекапа в:
echo %BACKUP_DIR%
echo.

choice /C YN /M "Продолжить"
if errorlevel 2 (
    echo Создание бекапа отменено.
    pause
    exit /b 0
)

echo.
echo Копирование файлов...

REM Create backup directory
mkdir "%BACKUP_DIR%" 2>nul

REM Copy annotatix_dependencies
xcopy /E /I /Y "%SOURCE_DIR%\*" "%BACKUP_DIR%\annotatix_dependencies\" >nul
if errorlevel 1 (
    echo [ERROR] Ошибка при копировании
    pause
    exit /b 1
)

REM Copy .addin file
copy /Y "%REVIT_ADDINS%\Annotatix.addin" "%BACKUP_DIR%\Annotatix.addin" >nul

echo.
echo ========================================
echo Бекап создан успешно!
echo ========================================
echo Папка: %BACKUP_DIR%
echo Версия: 3.%BUILD_NUMBER%
echo.
pause
