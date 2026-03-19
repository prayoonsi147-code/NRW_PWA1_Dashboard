@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo   PWA Region 1 Dashboard - Dev Server
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Download at: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install flask openpyxl xlrd >nul 2>&1

echo.
echo ========================================
echo   Select Dashboard to Run
echo ========================================
echo 1. Dashboard PR        (port 5000)
echo 2. Dashboard Leak      (port 5001)
echo 3. Dashboard GIS       (port 5002)
echo 4. Dashboard Meter     (port 5003)
echo 5. All servers
echo 6. Main + All servers
echo.

set /p choice="Enter your choice (1-6): "

if "%choice%"=="1" goto pr_only
if "%choice%"=="2" goto leak_only
if "%choice%"=="3" goto gis_only
if "%choice%"=="4" goto meter_only
if "%choice%"=="5" goto all
if "%choice%"=="6" goto main_all
goto end

:pr_only
echo Starting Dashboard PR on http://localhost:5000
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5000"
cd /d "%~dp0Dashboard_PR"
python server.py
goto end

:leak_only
echo Starting Dashboard Leak on http://localhost:5001
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5001"
cd /d "%~dp0Dashboard_Leak"
python server.py
goto end

:gis_only
echo Starting Dashboard GIS on http://localhost:5002
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5002"
cd /d "%~dp0Dashboard_GIS"
python server.py
goto end

:meter_only
echo Starting Dashboard Meter on http://localhost:5003
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5003"
cd /d "%~dp0Dashboard_Meter"
python server.py
goto end

:all
echo Starting all servers...
start "Dashboard PR" cmd /c "cd /d "%~dp0Dashboard_PR" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard Leak" cmd /c "cd /d "%~dp0Dashboard_Leak" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard GIS" cmd /c "cd /d "%~dp0Dashboard_GIS" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard Meter" cmd /c "cd /d "%~dp0Dashboard_Meter" && python server.py"
timeout /t 2 /nobreak >nul
echo.
echo All servers running:
echo   PR:    http://localhost:5000
echo   Leak:  http://localhost:5001
echo   GIS:   http://localhost:5002
echo   Meter: http://localhost:5003
echo.
echo Press any key to stop...
pause >nul
taskkill /FI "WINDOWTITLE eq Dashboard PR" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard Leak" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard GIS" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard Meter" >nul 2>&1
goto end

:main_all
echo Starting all servers + Main page...
start "Dashboard PR" cmd /c "cd /d "%~dp0Dashboard_PR" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard Leak" cmd /c "cd /d "%~dp0Dashboard_Leak" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard GIS" cmd /c "cd /d "%~dp0Dashboard_GIS" && python server.py"
timeout /t 1 /nobreak >nul
start "Dashboard Meter" cmd /c "cd /d "%~dp0Dashboard_Meter" && python server.py"
timeout /t 2 /nobreak >nul
start "" "%~dp0index.html"
echo.
echo ========================================
echo   All servers running:
echo   - Dashboard PR:    http://localhost:5000
echo   - Dashboard Leak:  http://localhost:5001
echo   - Dashboard GIS:   http://localhost:5002
echo   - Dashboard Meter: http://localhost:5003
echo   - Main page: opened in browser
echo ========================================
echo.
echo Press any key to stop all servers...
pause >nul
taskkill /FI "WINDOWTITLE eq Dashboard PR" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard Leak" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard GIS" >nul 2>&1
taskkill /FI "WINDOWTITLE eq Dashboard Meter" >nul 2>&1
goto end

:end
