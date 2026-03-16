@echo off
chcp 65001 >nul
echo ============================================================
echo   Update Dashboard
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python not found! Please install Python and add to PATH.
        echo         https://www.python.org/downloads/
        pause
        exit /b 1
    )
    py "%~dp0build_dashboard.py"
) else (
    python "%~dp0build_dashboard.py"
)

echo.
pause
