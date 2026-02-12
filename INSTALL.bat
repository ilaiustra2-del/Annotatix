@echo off
chcp 65001 >nul
echo ========================================
echo Annotatix - Установка/Обновление
echo ========================================
echo.

REM Get the latest built version folder
for /f "delims=" %%i in ('dir /b /ad /o-n "..\annotatix_ver3.*" 2^>nul ^| findstr /r "annotatix_ver3\.[0-9][0-9][0-9]$" ^| sort /r') do (
    set "LATEST_VERSION=%%i"
    goto :found
)

:found
if "%LATEST_VERSION%"=="" (
    echo [ERROR] Не найдена собранная версия annotatix_ver3.XXX
    echo.
    echo Сначала выполните сборку проекта через RUN_BUILD.bat
    pause
    exit /b 1
)

set "SOURCE_DIR=%~dp0..\%LATEST_VERSION%"
set "REVIT_ADDINS=%APPDATA%\Autodesk\Revit\Addins\2024"
set "TARGET_DIR=%REVIT_ADDINS%\annotatix_dependencies"
set "ADDIN_FILE=%REVIT_ADDINS%\Annotatix.addin"

echo Найдена версия для установки: %LATEST_VERSION%
echo.
echo Источник: %SOURCE_DIR%
echo Целевая папка: %TARGET_DIR%
echo.

REM Check if already installed
if exist "%TARGET_DIR%\main\BuildNumber.txt" (
    set /p CURRENT_BUILD=<"%TARGET_DIR%\main\BuildNumber.txt"
    echo Текущая установленная версия: 3.!CURRENT_BUILD!
    echo Новая версия: %LATEST_VERSION%
    echo.
    echo ВНИМАНИЕ: Установка перезапишет текущую версию!
    echo.
    choice /C YN /M "Продолжить установку"
    if errorlevel 2 (
        echo Установка отменена.
        pause
        exit /b 0
    )
) else (
    echo Annotatix ещё не установлен.
    echo.
    choice /C YN /M "Установить %LATEST_VERSION%"
    if errorlevel 2 (
        echo Установка отменена.
        pause
        exit /b 0
    )
)

echo.
echo ========================================
echo Установка...
echo ========================================

REM Close Revit reminder
echo.
echo ВАЖНО: Убедитесь, что Revit ПОЛНОСТЬЮ закрыт!
pause

REM Copy files
echo.
echo Копирование файлов annotatix_dependencies...
xcopy /E /I /Y "%SOURCE_DIR%\annotatix_dependencies\*" "%TARGET_DIR%\" >nul
if errorlevel 1 (
    echo [ERROR] Ошибка при копировании annotatix_dependencies
    pause
    exit /b 1
)
echo ✓ annotatix_dependencies скопированы

echo.
echo Копирование файла Annotatix.addin...
copy /Y "%SOURCE_DIR%\Annotatix.addin" "%ADDIN_FILE%" >nul
if errorlevel 1 (
    echo [ERROR] Ошибка при копировании Annotatix.addin
    pause
    exit /b 1
)
echo ✓ Annotatix.addin скопирован

echo.
echo Разблокировка файлов (снятие защиты Windows)...
powershell -Command "Get-ChildItem -Path '%TARGET_DIR%' -Recurse -File | Unblock-File" >nul 2>&1
powershell -Command "Unblock-File '%ADDIN_FILE%'" >nul 2>&1
echo ✓ Файлы разблокированы

echo.
echo ========================================
echo Установка завершена!
echo ========================================
echo Версия: %LATEST_VERSION%
echo.
echo Теперь можете запустить Revit.
echo.
pause
