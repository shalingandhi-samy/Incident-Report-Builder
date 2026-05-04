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

echo Starting server at http://127.0.0.1:8000
echo Press Ctrl+C to stop.
echo.

cd rca_app
"%~dp0.venv\Scripts\python.exe" main_v2.py

pause
