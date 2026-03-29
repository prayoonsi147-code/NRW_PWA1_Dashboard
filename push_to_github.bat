@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ============================================
echo   Push to GitHub: NRW_PWA1_Dashboard
echo ============================================
echo.

cd /d "%~dp0"

REM --- Step 0: Detect PHP (XAMPP) ---
set PHPCMD=
if exist "C:\xampp\php\php.exe" (
    set PHPCMD=C:\xampp\php\php.exe
) else (
    where php >nul 2>nul
    if not errorlevel 1 (
        set PHPCMD=php
    )
)

REM --- Step 1: Build all dashboards (update static data) ---
echo [1/8] Building dashboards (embed latest data)...
set BUILD_OK=1
if defined PHPCMD (
    for /d %%D in (Dashboard_*) do (
        if exist "%%D\build_dashboard.php" (
            echo   Building %%D...
            pushd "%%D"
            %PHPCMD% build_dashboard.php
            set PHPERR=!errorlevel!
            popd
            if !PHPERR! neq 0 (
                echo   [WARNING] %%D build failed ^(exit code !PHPERR!^) - continuing anyway
                set BUILD_OK=0
            ) else (
                echo   [OK] %%D
            )
        )
    )
) else (
    echo   [WARNING] PHP not found - skipping build step
    echo   Install XAMPP then try again: https://www.apachefriends.org/
    echo   Data in index.html may be outdated!
    set BUILD_OK=0
)
echo.

REM --- Step 2: Initialize git if not already ---
if not exist ".git" (
    echo [2/8] Initializing git repository...
    git init
    git branch -M main
    git remote add origin https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
    echo.
) else (
    echo [2/8] Git already initialized.
    echo.
)

REM --- Step 3: Set git identity ---
echo [3/8] Setting git identity...
git config user.email "prayoonsi147@gmail.com"
git config user.name "prayoonsi147-code"
echo.

REM --- Step 4: Pull latest from remote ---
echo [4/8] Pulling latest changes...
git pull origin main --allow-unrelated-histories 2>nul
echo.

REM --- Step 4.5: Squash unpushed commits to remove large files from history ---
echo [4.5/8] Cleaning unpushed history (prevent large file errors)...
git rev-parse origin/main >nul 2>nul
if not errorlevel 1 (
    echo   Soft reset to origin/main...
    git reset --soft origin/main
    echo   Done - will re-commit all changes as one clean commit.
)
echo.

REM --- Step 5: Stage all files (respects .gitignore) ---
echo [5/8] Staging all files...
git add -A
echo   Checking for oversized tracked files...
git ls-files --cached "*.sqlite" >nul 2>nul
git rm --cached "*.sqlite" >nul 2>nul
git rm --cached "*.cache.json" >nul 2>nul
echo.

REM --- Step 6: Show status ---
echo [6/8] Files to be committed:
git status --short
echo.

REM --- Step 7: Commit ---
echo [7/8] Creating commit...
git commit -m "Update dashboard"
if errorlevel 1 (
    echo.
    echo No changes to commit or commit failed.
    pause
    exit /b 1
)
echo.

REM --- Step 8: Push ---
echo [8/8] Pushing to GitHub...
git push -u origin main
echo.

echo ============================================
if "!BUILD_OK!"=="1" (
    echo   Done! All dashboards built + pushed.
) else (
    echo   Done! Pushed, but some builds had warnings.
)
echo   Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
