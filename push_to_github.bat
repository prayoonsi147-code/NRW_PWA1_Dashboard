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
    if not errorlevel 1 set PHPCMD=php
)

if not defined PHPCMD (
    echo   [ABORT] PHP not found! Cannot build dashboards.
    echo   Install XAMPP or add PHP to PATH.
    pause
    exit /b 1
)
echo   PHP found: %PHPCMD%
echo.

REM ============================================================
REM   Step 1: CHECKPOINT
REM ============================================================
echo [1/9] Creating checkpoint...
set CHECKPOINT_OK=1
set DASHBOARDS=Dashboard_PR Dashboard_Leak Dashboard_GIS Dashboard_Meter
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html" (
        copy /Y "%%D\index.html" "%%D\index.html.checkpoint" >nul 2>nul
        if errorlevel 1 (
            echo   [ERROR] Failed to backup %%D\index.html
            set CHECKPOINT_OK=0
        ) else (
            echo   [OK] %%D checkpoint created
        )
    )
)
if "!CHECKPOINT_OK!"=="0" (
    echo   [ABORT] Checkpoint failed.
    pause
    exit /b 1
)
echo.

REM ============================================================
REM   Step 2: BUILD
REM ============================================================
echo [2/9] Building dashboards...
set BUILD_OK=1
set BUILD_COUNT=0
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\build_dashboard.php" (
        echo   Building %%D...
        pushd "%%D"
        %PHPCMD% build_dashboard.php
        set PHPERR=!errorlevel!
        popd
        if !PHPERR! neq 0 (
            echo   [FAIL] %%D build failed
            set BUILD_OK=0
        ) else (
            echo   [OK] %%D built
            set /a BUILD_COUNT+=1
        )
    )
)
if "!BUILD_OK!"=="0" (
    echo   [ABORT] Build failed! Restoring checkpoints...
    goto RESTORE_ABORT
)
if "!BUILD_COUNT!"=="0" (
    echo   [ABORT] No dashboards built!
    goto RESTORE_ABORT
)
echo   All !BUILD_COUNT! dashboards built.
echo.

REM ============================================================
REM   Step 3: VALIDATE (use subroutine to avoid nesting issues)
REM ============================================================
echo [3/9] Validating...
set VALID_OK=1
for %%D in (%DASHBOARDS%) do call :validate_one %%D
if "!VALID_OK!"=="0" (
    echo   [ABORT] Validation failed!
    goto RESTORE_ABORT
)
echo.

REM ============================================================
REM   Step 4: Init git
REM ============================================================
if not exist ".git" (
    echo [4/9] Initializing git...
    git init
    git branch -M main
    git remote add origin https://github.com/prayoonsi147-code/NRW_PWA1_Dashboard.git
) else (
    echo [4/9] Git OK.
)
echo.

REM --- Step 5: Identity ---
echo [5/9] Setting git identity...
git config user.email "prayoonsi147@gmail.com"
git config user.name "prayoonsi147-code"
echo.

REM --- Step 5.5: Lock cleanup ---
if exist ".git\index.lock" (
    echo   [CLEANUP] Removing stale .git/index.lock...
    del /F ".git\index.lock" >nul 2>nul
)

REM --- Step 6: Pull + squash ---
echo [6/9] Pulling latest...
git pull origin main --allow-unrelated-histories 2>nul
git rev-parse origin/main >nul 2>nul
if not errorlevel 1 git reset --soft origin/main
echo.

REM --- Step 7: Stage ---
if exist ".git\index.lock" del /F ".git\index.lock" >nul 2>nul
echo [7/9] Staging files...
git add -A
git rm --cached "*.sqlite" >nul 2>nul
git rm --cached "*.cache.json" >nul 2>nul
git rm --cached "*.checkpoint" >nul 2>nul
echo   Files to commit:
git status --short
echo.

REM --- Step 8: Commit + Push ---
if exist ".git\index.lock" del /F ".git\index.lock" >nul 2>nul
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
REM   Step 9: RESTORE
REM ============================================================
:RESTORE
echo [9/9] Restoring checkpoints...
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html.checkpoint" (
        copy /Y "%%D\index.html.checkpoint" "%%D\index.html" >nul
        del "%%D\index.html.checkpoint" >nul 2>nul
        echo   [OK] %%D restored
    )
)
echo.
echo ============================================
echo   Done! Check: https://prayoonsi147-code.github.io/NRW_PWA1_Dashboard/
echo ============================================
pause
exit /b 0

REM ============================================================
REM   RESTORE_ABORT
REM ============================================================
:RESTORE_ABORT
echo.
echo   Restoring all checkpoints (no push)...
for %%D in (%DASHBOARDS%) do (
    if exist "%%D\index.html.checkpoint" (
        copy /Y "%%D\index.html.checkpoint" "%%D\index.html" >nul
        del "%%D\index.html.checkpoint" >nul 2>nul
        echo   [OK] %%D restored
    )
)
echo.
echo ============================================
echo   [ABORTED] Push cancelled - build/validation failed.
echo   Local files are safe.
echo ============================================
pause
exit /b 1

REM ============================================================
REM   Subroutine: validate one dashboard (avoids nested for/if)
REM ============================================================
:validate_one
set "_DD=%~1"
if not exist "%_DD%\index.html" goto :eof

REM Check DOCTYPE
findstr /i "DOCTYPE" "%_DD%\index.html" >nul 2>nul
if errorlevel 1 (
    echo   [FAIL] %_DD% missing DOCTYPE - restoring
    copy /Y "%_DD%\index.html.checkpoint" "%_DD%\index.html" >nul
    set VALID_OK=0
    goto :eof
)

REM Get sizes
set "_SZ_BUILT=0"
set "_SZ_CHECK=0"
for %%F in ("%_DD%\index.html") do set "_SZ_BUILT=%%~zF"
for %%C in ("%_DD%\index.html.checkpoint") do set "_SZ_CHECK=%%~zC"

REM Check minimum size
if !_SZ_BUILT! LSS 1024 (
    echo   [FAIL] %_DD% too small - restoring
    copy /Y "%_DD%\index.html.checkpoint" "%_DD%\index.html" >nul
    set VALID_OK=0
    goto :eof
)

REM Check built is not drastically smaller than checkpoint
set /a "_MIN_SZ=!_SZ_CHECK! / 2"
if !_SZ_BUILT! LSS !_MIN_SZ! (
    echo   [FAIL] %_DD% much smaller than checkpoint [!_SZ_BUILT! vs !_SZ_CHECK! bytes]
    echo          Build likely did not embed data!
    copy /Y "%_DD%\index.html.checkpoint" "%_DD%\index.html" >nul
    set VALID_OK=0
    goto :eof
)

echo   [OK] %_DD% valid [!_SZ_BUILT! bytes, was !_SZ_CHECK!]
goto :eof
