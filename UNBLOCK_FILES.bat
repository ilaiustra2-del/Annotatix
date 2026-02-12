@echo off
chcp 65001 >nul
echo ========================================
echo Annotatix - Разблокировка файлов
echo ========================================
echo.
echo Этот скрипт снимает блокировку Windows с DLL файлов
echo чтобы Revit мог их загрузить.
echo.

set "REVIT_ADDINS=%APPDATA%\Autodesk\Revit\Addins\2024\annotatix_dependencies"

if not exist "%REVIT_ADDINS%\main\PluginsManager.dll" (
    echo [ERROR] Файлы плагина не найдены в Revit Addins
    echo Сначала установите плагин.
    pause
    exit /b 1
)

echo Разблокировка файлов в: %REVIT_ADDINS%
echo.

powershell -Command "Get-ChildItem -Path '%REVIT_ADDINS%' -Recurse -File | Unblock-File"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Готово! Все файлы разблокированы.
    echo ========================================
    echo.
    echo Теперь можете запустить Revit.
) else (
    echo.
    echo [ERROR] Ошибка при разблокировке файлов
    echo Попробуйте запустить скрипт от имени администратора.
)

echo.
pause
