@echo off
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

echo [1/2] Updating Dashboard Water Loss ...
for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.py" (
        if /i not "%%D"=="Dashboard_PR" (
            echo   Found: %%D
            cd "%%D"
            %PYCMD% build_dashboard.py
            cd ..
            if errorlevel 1 (
                echo   [ERROR] %%D update failed!
            ) else (
                echo   [OK] %%D updated.
            )
        )
    )
)
echo.

echo [2/2] Updating Dashboard PR ...
cd Dashboard_PR
%PYCMD% build_dashboard.py
cd ..
if errorlevel 1 (
    echo   [ERROR] Dashboard PR update failed!
) else (
    echo   [OK] Dashboard PR updated.
)
echo.

echo ============================================================
echo   Done!
echo ============================================================
pause
