@echo off
title Dashboard PWA-R1
echo ========================================
echo   Dashboard PWA Region 1 (XAMPP)
echo ========================================
echo.

:: Check XAMPP
if not exist "C:\xampp\apache\bin\httpd.exe" (
    echo [ERROR] XAMPP not found at C:\xampp
    pause
    exit /b 1
)
echo [OK] XAMPP found

:: Auto-setup: check DocumentRoot
set "CONF=C:\xampp\apache\conf\httpd.conf"
set "USER_HOME=%USERPROFILE%"
set "APACHE_PATH=%USER_HOME:\=/%"
findstr /C:"DocumentRoot \"%APACHE_PATH%\"" "%CONF%" >NUL 2>&1
if %errorlevel% neq 0 call :do_setup

:: Kill old Apache
tasklist /FI "IMAGENAME eq httpd.exe" 2>NUL | find /I "httpd.exe" >NUL 2>&1
if %errorlevel%==0 taskkill /F /IM httpd.exe >NUL 2>&1
timeout /t 1 /nobreak >NUL

:: Start Apache
echo [..] Starting Apache...
start "Apache-HTTPD" /min "C:\xampp\apache\bin\httpd.exe"
timeout /t 3 /nobreak >NUL

tasklist /FI "IMAGENAME eq httpd.exe" 2>NUL | find /I "httpd.exe" >NUL 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Apache failed to start!
    "C:\xampp\apache\bin\httpd.exe" -t 2>&1
    pause
    exit /b 1
)
echo [OK] Apache is running

:: Open browser
start "" "http://localhost/Claude%%20Test%%20Cowork/index.html"

echo.
echo ========================================
echo   http://localhost/Claude Test Cowork/
echo   Press any key to stop Apache
echo ========================================
pause

taskkill /F /IM httpd.exe >NUL 2>&1
echo Stopped.
timeout /t 1 /nobreak >NUL
goto :eof

:do_setup
echo [Setup] First time setup...
copy /Y "%CONF%" "%CONF%.backup" >NUL 2>&1
powershell -Command "(Get-Content '%CONF%') -replace 'DocumentRoot \"C:/[^\"]*\"', 'DocumentRoot \"%APACHE_PATH%\"' | Set-Content '%CONF%'"
powershell -Command "(Get-Content '%CONF%') -replace '<Directory \"C:/[^\"]*\">', '<Directory \"%APACHE_PATH%\">' | Set-Content '%CONF%'"
echo   [1/3] DocumentRoot = %USER_HOME%
set "PHPINI=C:\xampp\php\php.ini"
if not exist "%PHPINI%" goto :skip_php
copy /Y "%PHPINI%" "%PHPINI%.backup" >NUL 2>&1
powershell -Command "(Get-Content '%PHPINI%') -replace '^;extension=gd', 'extension=gd' | Set-Content '%PHPINI%'"
powershell -Command "(Get-Content '%PHPINI%') -replace '^;extension=zip', 'extension=zip' | Set-Content '%PHPINI%'"
powershell -Command "(Get-Content '%PHPINI%') -replace '^;extension=sqlite3', 'extension=sqlite3' | Set-Content '%PHPINI%'"
echo   [2/3] PHP extensions enabled
:skip_php
if exist "%USER_HOME%\Claude Test Cowork\vendor\autoload.php" (
    echo   [3/3] vendor/ OK
) else (
    echo   [3/3] WARNING: vendor/ not found. Run: composer require phpoffice/phpspreadsheet
)
echo [Setup] Done!
echo.
goto :eof
