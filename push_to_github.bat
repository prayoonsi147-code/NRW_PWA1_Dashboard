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

REM ============================================================
REM   Step 1: CHECKPOINT — backup index.html ทุก Dashboard
REM ============================================================
echo [1/9] Creating checkpoint (backup index.html)...
set CHECKPOINT_OK=1
set DASHBOARDS=Dashboard_PR Dashboard_Leak Dashboard_GIS Dashboard_Meter
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html" (
        copy /Y "%%D\index.html" "%%D\index.html.checkpoint" >nul 2>nul
        if errorlevel 1 (
            echo   [ERROR] Failed to backup %%D\index.html
            set CHECKPOINT_OK=0
        ) else (
            echo   [OK] %%D\index.html.checkpoint created
        )
    ) else (
        echo   [SKIP] %%D\index.html not found
    )
)
if "!CHECKPOINT_OK!"=="0" (
    echo.
    echo   [ABORT] Checkpoint failed - cannot continue safely.
    pause
    exit /b 1
)
echo.

REM ============================================================
REM   Step 2: BUILD — ฝังข้อมูลล่าสุดลง index.html ทุก Dashboard
REM ============================================================
echo [2/9] Building dashboards (embed latest data into HTML)...
set BUILD_OK=1
if defined PHPCMD (
    for %%D in (%DASHBOARDS%) do (
        if exist "%%D\build_dashboard.php" (
            echo   Building %%D...
            pushd "%%D"
            %PHPCMD% build_dashboard.php
            set PHPERR=!errorlevel!
            popd
            if !PHPERR! neq 0 (
                echo   [WARNING] %%D build failed ^(exit code !PHPERR!^)
                set BUILD_OK=0
            ) else (
                echo   [OK] %%D built successfully
            )
        ) else (
            echo   [SKIP] %%D has no build_dashboard.php
        )
    )
) else (
    echo   [WARNING] PHP not found - skipping build step
    echo   Install XAMPP: https://www.apachefriends.org/
    set BUILD_OK=0
)
echo.

REM ============================================================
REM   Step 3: VALIDATE — ตรวจว่า index.html ทุกตัวยังเป็น HTML ปกติ
REM ============================================================
echo [3/9] Validating built HTML files...
set VALID_OK=1
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html" (
        REM Check 1: file must contain DOCTYPE
        findstr /i "DOCTYPE" "%%D\index.html" >nul 2>nul
        if errorlevel 1 (
            echo   [FAIL] %%D\index.html missing DOCTYPE - restoring checkpoint
            copy /Y "%%D\index.html.checkpoint" "%%D\index.html" >nul
            set VALID_OK=0
        ) else (
            REM Check 2: file must be at least 1KB (not truncated)
            for %%F in ("%%D\index.html") do (
                if %%~zF LSS 1024 (
                    echo   [FAIL] %%D\index.html too small ^(%%~zF bytes^) - restoring checkpoint
                    copy /Y "%%D\index.html.checkpoint" "%%D\index.html" >nul
                    set VALID_OK=0
                ) else (
                    echo   [OK] %%D\index.html valid ^(%%~zF bytes^)
                )
            )
        )
    )
)
if "!VALID_OK!"=="0" (
    echo.
    echo   [WARNING] Some HTML files were invalid and restored from checkpoint.
    echo   Build script may have bugs. Push will continue with restored files.
)
echo.

REM ============================================================
REM   Step 4: Initialize git if needed
REM ============================================================
if not exist ".git" (
    echo [4/9] Initializing git repository...
    git init
    git branch -M main
    git remote add origin https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
    echo.
) else (
    echo [4/9] Git already initialized.
    echo.
)

REM --- Step 5: Set git identity ---
echo [5/9] Setting git identity...
git config user.email "prayoonsi147@gmail.com"
git config user.name "prayoonsi147-code"
echo.

REM --- Step 5.5: Remove stale lock file (e.g. left by Cowork/other git process) ---
if exist ".git\index.lock" (
    echo   [CLEANUP] Removing stale .git/index.lock...
    del /F ".git\index.lock" >nul 2>nul
)

REM --- Step 6: Pull + squash ---
echo [6/9] Pulling latest + cleaning history...
git pull origin main --allow-unrelated-histories 2>nul
git rev-parse origin/main >nul 2>nul
if not errorlevel 1 (
    git reset --soft origin/main
)
echo.

REM --- Step 7: Stage all files (respects .gitignore) ---
if exist ".git\index.lock" ( del /F ".git\index.lock" >nul 2>nul )
echo [7/9] Staging files...
git add -A
REM Remove oversized files that may have slipped in
git rm --cached "*.sqlite" >nul 2>nul
git rm --cached "*.cache.json" >nul 2>nul
git rm --cached "*.checkpoint" >nul 2>nul
echo   Files to commit:
git status --short
echo.

REM --- Step 8: Commit + Push ---
if exist ".git\index.lock" ( del /F ".git\index.lock" >nul 2>nul )
echo [8/9] Committing and pushing...
git commit -m "Update dashboard data"
if errorlevel 1 (
    echo   No changes to commit.
    goto RESTORE
)
git push -u origin main
if errorlevel 1 (
    echo   [ERROR] Push failed!
    goto RESTORE
)
echo.

REM ============================================================
REM   Step 9: RESTORE — กลับไปใช้ index.html เดิมตาม checkpoint
REM ============================================================
:RESTORE
echo [9/9] Restoring checkpoint (original index.html)...
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html.checkpoint" (
        copy /Y "%%D\index.html.checkpoint" "%%D\index.html" >nul
        del "%%D\index.html.checkpoint" >nul 2>nul
        echo   [OK] %%D\index.html restored
    )
)
echo.

echo ============================================
if "!BUILD_OK!"=="1" if "!VALID_OK!"=="1" (
    echo   Done! All dashboards built + validated + pushed.
) else (
    echo   Done! Pushed, but some builds had warnings.
)
echo   Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
