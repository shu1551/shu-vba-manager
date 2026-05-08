@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
py "project\python_scripts\menu_launcher.py"
if %ERRORLEVEL% neq 0 (
    echo.
    echo エラーが発生しました。
    pause
)
endlocal
