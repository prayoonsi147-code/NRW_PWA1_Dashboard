@echo off
chcp 65001 >nul
title Rollback PhpSpreadsheet
cd /d "%~dp0"

echo.
echo ============================================================
echo  Rollback: restore PhpSpreadsheet to previous version
echo ============================================================
echo.

set "BACKUP_DIR=backup_php74_install"
if not exist "%BACKUP_DIR%" (
    echo [X] %BACKUP_DIR% not found. Nothing to rollback.
    pause
    exit /b 1
)

echo Will restore from:
dir /B "%BACKUP_DIR%"
echo.

set /p CONFIRM=Type YES to confirm rollback:
if /I not "%CONFIRM%"=="YES" (
    echo Cancelled.
    pause
    exit /b 0
)
echo.

echo [1/4] Restore composer.json
if exist "%BACKUP_DIR%\composer.json.before" (
    copy /Y "%BACKUP_DIR%\composer.json.before" composer.json >nul
    echo [OK]
) else (
    echo [!] backup not found, skipped.
)

echo [2/4] Restore composer.lock
if exist "%BACKUP_DIR%\composer.lock.before" (
    copy /Y "%BACKUP_DIR%\composer.lock.before" composer.lock >nul
    echo [OK]
) else (
    if exist composer.lock del /F /Q composer.lock
    echo [!] backup not found, deleted current composer.lock instead.
)

echo [3/4] Remove current vendor
if exist vendor rmdir /S /Q vendor
echo [OK]

echo [4/4] Restore vendor from backup
if exist "%BACKUP_DIR%\vendor.before" (
    echo      copy vendor back (~10 sec)
    xcopy /E /I /Q /Y "%BACKUP_DIR%\vendor.before" vendor >nul
    echo [OK]
) else (
    echo [!] backup vendor not found.
    echo     You need to run: php composer.phar install
)

echo.
echo ============================================================
echo  Rollback done. System restored to pre-install state.
echo ============================================================
echo.
pause
