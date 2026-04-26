@echo off
chcp 65001 >nul
title Install PhpSpreadsheet 1.28 for PHP 7.4
cd /d "%~dp0"

echo.
echo ============================================================
echo  Install PhpSpreadsheet 1.28 (PHP 7.4 compatible)
echo ============================================================
echo.

REM --- Find PHP ---
set "PHP_EXE=C:\xampp\php\php.exe"
if not exist "%PHP_EXE%" (
    where php >nul 2>nul
    if errorlevel 1 (
        echo [X] PHP not found. Install XAMPP first.
        pause
        exit /b 1
    )
    set "PHP_EXE=php"
)

echo [1/7] PHP version:
"%PHP_EXE%" -v
if errorlevel 1 (
    echo [X] PHP failed to run.
    pause
    exit /b 1
)
echo.

REM --- Check ALL required extensions at once ---
"%PHP_EXE%" check_extensions.php
if errorlevel 1 (
    pause
    exit /b 1
)
echo.

REM --- Backup ---
echo [2/7] Backup composer.json + composer.lock + vendor
set "BACKUP_DIR=backup_php74_install"
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
if exist composer.json copy /Y composer.json "%BACKUP_DIR%\composer.json.before" >nul
if exist composer.lock copy /Y composer.lock "%BACKUP_DIR%\composer.lock.before" >nul
if exist vendor (
    if not exist "%BACKUP_DIR%\vendor.before" (
        echo      backup vendor (~10 sec)
        xcopy /E /I /Q /Y vendor "%BACKUP_DIR%\vendor.before" >nul
    )
)
echo [OK] backup done.
echo.

REM --- Untrack vendor from git (best-effort, no error if not a git repo) ---
echo [3/7] Untrack vendor from git (best-effort)
git rm --cached -r vendor >nul 2>nul
echo [OK]
echo.

REM --- Remove old vendor + composer.lock ---
echo [4/7] Remove old vendor + composer.lock
if exist composer.lock del /F /Q composer.lock
if exist vendor rmdir /S /Q vendor
echo [OK]
echo.

REM --- composer install ---
echo [5/7] Run: php composer.phar install
"%PHP_EXE%" composer.phar install --no-interaction --prefer-dist
if errorlevel 1 (
    echo.
    echo [X] composer install failed. Run rollback_php74.bat to undo.
    pause
    exit /b 1
)
echo.

REM --- Verify version ---
echo [6/7] Verify installed version
"%PHP_EXE%" verify_install.php
if errorlevel 1 (
    echo [!] verify_install.php reported issues, but install proceeded.
)
echo.

REM --- Smoke test ---
echo [7/7] Smoke test: test_php74.php
if exist test_php74.php (
    "%PHP_EXE%" test_php74.php
    if errorlevel 1 (
        echo.
        echo [!] Smoke test reported failures. Read messages above.
        echo     If unacceptable, run rollback_php74.bat to undo.
        pause
        exit /b 1
    )
) else (
    echo [!] test_php74.php not found, skipped.
)
echo.

echo ============================================================
echo  DONE. Next steps:
echo   1. Open localhost and test all 4 dashboards
echo   2. If OK, run push_to_github.bat
echo   3. If broken, run rollback_php74.bat
echo ============================================================
echo.
pause
