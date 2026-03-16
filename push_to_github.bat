@echo off
chcp 65001 >nul
echo ============================================
echo   Push to GitHub: NRW_PWA1_Dashboard
echo ============================================
echo.

REM --- Step 1: Initialize git if not already ---
if not exist ".git" (
    echo [1/7] Initializing git repository...
    git init
    git branch -M main
    git remote add origin https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
    echo.
) else (
    echo [1/7] Git already initialized.
    echo.
)

REM --- Step 2: Set git identity ---
echo [2/7] Setting git identity...
git config user.email "prayoonsi147@gmail.com"
git config user.name "prayoonsi147-code"
echo.

REM --- Step 3: Pull latest from remote ---
echo [3/7] Pulling latest changes...
git pull origin main --allow-unrelated-histories 2>nul
echo.

REM --- Step 4: Stage all files ---
echo [4/7] Staging all files...
git add -A
echo.

REM --- Step 5: Show status ---
echo [5/7] Files to be committed:
git status --short
echo.

REM --- Step 6: Commit ---
echo [6/7] Creating commit...
git commit -m "Update dashboard"
if errorlevel 1 (
    echo.
    echo No changes to commit or commit failed.
    pause
    exit /b 1
)
echo.

REM --- Step 7: Push ---
echo [7/7] Pushing to GitHub...
git push -u origin main
echo.

echo ============================================
echo   Done! Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
