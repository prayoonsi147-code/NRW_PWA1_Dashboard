@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ============================================================
echo   Update All Dashboards
echo ============================================================
echo.

cd /d "%~dp0"

:: Detect Python command
set PYCMD=
where python >nul 2>nul
if not errorlevel 1 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        set PYCMD=py
    ) else (
        echo [ERROR] Python not found! Please install Python and add to PATH.
        echo         https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

set /a COUNT=0
set /a TOTAL=0
for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.py" (
        set /a TOTAL+=1
    )
)

for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.py" (
        set /a COUNT+=1
        echo [!COUNT!/%TOTAL%] Updating %%D ...
        cd "%%D"
        %PYCMD% build_dashboard.py
        cd ..
        if errorlevel 1 (
            echo   [ERROR] %%D update failed!
        ) else (
            echo   [OK] %%D updated.
        )
        echo.
    )
)

echo ============================================================
echo   All %TOTAL% dashboards updated!
echo ============================================================
pause
