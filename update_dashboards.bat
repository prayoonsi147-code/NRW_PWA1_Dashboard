@echo off
chcp 874 >nul
echo ============================================================
echo   Update All Dashboards
echo ============================================================
echo.

cd /d "%~dp0"

echo [1/2] Updating Dashboard Water Loss ...
for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.py" (
        if /i not "%%D"=="Dashboard_PR" (
            echo   Found: %%D
            cd "%%D"
            py build_dashboard.py
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
py build_dashboard.py
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
