@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ============================================
echo   PIF 難民特賣會 - 一鍵同步 Excel 到網站
echo ============================================
echo.

REM 嘗試找 Python
where python >nul 2>nul
if %errorlevel% equ 0 (
    python sync_from_excel.py
    goto :end
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    py sync_from_excel.py
    goto :end
)

echo.
echo X 找不到 Python，請先安裝：
echo   https://www.python.org/downloads/
echo   安裝時請勾選「Add Python to PATH」
echo.
pause
exit /b 1

:end
