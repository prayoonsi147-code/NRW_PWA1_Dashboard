@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ============================================================
echo   Update All Dashboards (PHP)
echo ============================================================
echo.

cd /d "%~dp0"

:: Detect PHP command (XAMPP)
set PHPCMD=
if exist "C:\xampp\php\php.exe" (
    set PHPCMD=C:\xampp\php\php.exe
) else (
    where php >nul 2>nul
    if not errorlevel 1 (
        set PHPCMD=php
    ) else (
        echo [ERROR] PHP not found!
        echo         ติดตั้ง XAMPP แล้วลองอีกครั้ง: https://www.apachefriends.org/
        pause
        exit /b 1
    )
)

echo   ใช้ PHP: %PHPCMD%
echo.

set /a COUNT=0
set /a TOTAL=0
for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.php" (
        set /a TOTAL+=1
    )
)

for /d %%D in (Dashboard_*) do (
    if exist "%%D\build_dashboard.php" (
        set /a COUNT+=1
        echo [!COUNT!/%TOTAL%] Updating %%D ...
        pushd "%%D"
        %PHPCMD% build_dashboard.php
        set PHPERR=!errorlevel!
        popd
        if !PHPERR! neq 0 (
            echo   [ERROR] %%D update failed! ^(exit code !PHPERR!^)
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
