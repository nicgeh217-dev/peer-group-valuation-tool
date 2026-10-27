@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "LOG=%~dp0last_run.log"
set PYTHONNOUSERSITE=
set PYTHONPATH=

REM ============================================================
REM  Auto-detect a working Python (with tkinter) — portable,
REM  no hard-coded path. Works on any PC that has Python 3.
REM ============================================================
set "PYEXE="

REM 1) Python launcher (most reliable on Windows)
py -3 -c "import tkinter" >nul 2>&1 && set "PYEXE=py -3"

REM 2) python on PATH (skips the Microsoft Store stub, which lacks tkinter)
if not defined PYEXE (
    python -c "import tkinter" >nul 2>&1 && set "PYEXE=python"
)

REM 3) Scan common install locations
if not defined PYEXE (
    for %%D in (
        "C:\Python314" "C:\Python313" "C:\Python312" "C:\Python311" "C:\Python310"
        "%LOCALAPPDATA%\Programs\Python\Python314"
        "%LOCALAPPDATA%\Programs\Python\Python313"
        "%LOCALAPPDATA%\Programs\Python\Python312"
        "%LOCALAPPDATA%\Programs\Python\Python311"
    ) do (
        if not defined PYEXE (
            if exist "%%~D\python.exe" (
                "%%~D\python.exe" -c "import tkinter" >nul 2>&1 && set "PYEXE=%%~D\python.exe"
            )
        )
    )
)

if not defined PYEXE (
    echo ============================================================
    echo  Kein passendes Python gefunden.
    echo  Bitte Python 3.11+ installieren von https://www.python.org/downloads/
    echo  ^(beim Setup "Add Python to PATH" anhaken^)
    echo ============================================================
    pause
    exit /b 1
)

echo Python gefunden: %PYEXE%
echo.

REM ---- Dependencies check + auto-install ----
%PYEXE% -c "import pandas, openpyxl, yfinance, requests, numpy, pptx" >nul 2>&1
if errorlevel 1 (
    echo Installiere benoetigte Pakete ^(einmalig, ca. 1-2 Min^)...
    %PYEXE% -m pip install --user -r requirements.txt
)

REM ---- Launch GUI ----
%PYEXE% gui.py
if errorlevel 1 (
    echo.
    echo Fehler beim Start. Details siehe oben.
    pause
)
