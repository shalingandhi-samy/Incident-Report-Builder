@echo off
:: RCA Incident Report Builder - Launcher
:: Requires: Blank RCA Deck.pptx and WMW-738   new template.xlsx in this folder

echo.
echo ===================================================
echo   Incident Report Builder - Safety RCA Generator
echo ===================================================
echo.

:: Check for template files
if not exist "Blank RCA Deck.pptx" (
    echo [WARNING] Missing: "Blank RCA Deck.pptx"
    echo           Place the RCA PowerPoint template in this folder first!
    echo.
)

if not exist "WMW-738   new template.xlsx" (
    echo [WARNING] Missing: "WMW-738   new template.xlsx"
    echo           Place the Excel template in this folder first!
    echo           Note: filename has THREE spaces before "new"
    echo.
)

:: Get local IP for team sharing
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set LOCAL_IP=%%a
    goto :found_ip
)
:found_ip
set LOCAL_IP=%LOCAL_IP: =%

echo -----------------------------------------------
echo   YOUR LINK:   http://localhost:8000
echo   TEAM LINK:   http://%LOCAL_IP%:8000
echo -----------------------------------------------
echo   Share the TEAM LINK with anyone on the same
echo   network / Walmart VPN!
echo.
echo Press Ctrl+C to stop the server.
echo.

cd rca_app
"%~dp0.venv\Scripts\python.exe" main_v2.py

pause
