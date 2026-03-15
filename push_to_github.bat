@echo off
chcp 65001 >nul
echo ============================================
echo   Push to GitHub: NRW_PWA1_Dashboard
echo ============================================
echo.

REM --- Step 1: Initialize git if not already ---
if not exist ".git" (
    echo [1/6] Initializing git repository...
    git init
    git branch -M main
    git remote add origin https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
    echo.
) else (
    echo [1/6] Git already initialized.
    echo.
)

REM --- Step 2: Set git identity (AFTER init so .git exists) ---
echo [2/6] Setting git identity...
git config user.email "prayoonsi147@gmail.com"
git config user.name "prayoonsi147-code"
echo.

REM --- Step 3: Stage all files ---
echo [3/6] Staging all files...
git add -A
echo.

REM --- Step 4: Show status ---
echo [4/6] Files to be committed:
git status --short
echo.

REM --- Step 5: Commit ---
echo [5/6] Creating commit...
git commit -m "Update dashboards: rename folder to Dashboard_Leak, fix charts, update references"
if errorlevel 1 (
    echo.
    echo *** Commit failed! Check errors above ***
    pause
    exit /b 1
)
echo.

REM --- Step 6: Push ---
echo [6/6] Pushing to GitHub...
git push -u origin main
if errorlevel 1 (
    echo.
    echo *** Push failed! Trying force push... ***
    git push -u origin main --force
)
echo.

echo ============================================
echo   Done! Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
