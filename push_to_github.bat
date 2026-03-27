@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ============================================
echo   Push to GitHub: NRW_PWA1_Dashboard
echo ============================================
echo.

cd /d "%~dp0"

REM --- Step 0: Detect Python ---
set PYCMD=
where python >nul 2>nul
if not errorlevel 1 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        set PYCMD=py
    )
)

REM --- Step 1: Build all dashboards (update static data) ---
echo [1/8] Building dashboards (embed latest data)...
set BUILD_OK=1
if defined PYCMD (
    for /d %%D in (Dashboard_*) do (
        if exist "%%D\build_dashboard.py" (
            echo   Building %%D...
            pushd "%%D"
            %PYCMD% build_dashboard.py
            set PYERR=!errorlevel!
            popd
            if !PYERR! neq 0 (
                echo   [WARNING] %%D build failed ^(exit code !PYERR!^) - continuing anyway
                set BUILD_OK=0
            ) else (
                echo   [OK] %%D
            )
        )
    )
) else (
    echo   [WARNING] Python not found - skipping build step
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

REM --- Step 5: Stage all files ---
echo [5/8] Staging all files...
git add -A
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
if !BUILD_OK!==1 (
    echo   Done! All dashboards built + pushed.
) else (
    echo   Done! Pushed, but some builds had warnings.
)
echo   Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
