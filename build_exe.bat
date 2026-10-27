@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM ============================================================
REM  Builds a standalone Peer_Screener.exe with PyInstaller.
REM  Run once when you need a portable .exe (e.g. for a demo on
REM  a PC without Python). Output lands in: dist\Peer_Screener.exe
REM ============================================================

set "PYEXE="
py -3 -c "import sys" >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE python -c "import sys" >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
    for %%D in ("C:\Python314" "C:\Python313" "C:\Python312" "C:\Python311") do (
        if not defined PYEXE if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)
if not defined PYEXE (
    echo Kein Python gefunden. Bitte Python 3.11+ installieren.
    pause & exit /b 1
)

echo Python: %PYEXE%
echo.
echo [1/2] Installiere Build-Abhaengigkeiten ^(PyInstaller + Pakete^)...
%PYEXE% -m pip install --user pyinstaller -r requirements.txt

echo.
echo [2/2] Baue Peer_Screener.exe ^(dauert 2-3 Minuten^)...
REM --windowed = kein Konsolenfenster; --add-data bundelt die Datendateien
%PYEXE% -m PyInstaller --onefile --windowed --name Peer_Screener ^
    --add-data "universes.py;." ^
    --add-data "transactions.csv;." ^
    --add-data "transactions_template.csv;." ^
    --hidden-import pandas --hidden-import openpyxl --hidden-import yfinance ^
    --hidden-import pptx --hidden-import numpy ^
    --collect-all yfinance --collect-all curl_cffi ^
    gui.py

echo.
if exist "dist\Peer_Screener.exe" (
    echo ============================================================
    echo  FERTIG:  dist\Peer_Screener.exe
    echo.
    echo  WICHTIG - daneben legen ^(in denselben Ordner wie die .exe^):
    echo    - .env                  ^(Gemini API-Key^)
    echo    - transactions.csv      ^(M&A-Deals^)
    echo    - einen "output" Ordner ^(wird sonst automatisch erstellt^)
    echo ============================================================
) else (
    echo Build fehlgeschlagen. Siehe Meldungen oben.
)
pause
