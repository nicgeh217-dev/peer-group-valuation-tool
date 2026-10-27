@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
set PYTHONNOUSERSITE=
set PYTHONPATH=

REM Auto-detect Python
set "PYEXE="
py -3 -c "import streamlit" >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE python -c "import streamlit" >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
    for %%D in ("C:\Python314" "C:\Python313" "C:\Python312" "C:\Python311") do (
        if not defined PYEXE if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)
if not defined PYEXE (
    py -3 -c "import sys" >nul 2>&1 && set "PYEXE=py -3"
    if not defined PYEXE python -c "import sys" >nul 2>&1 && set "PYEXE=python"
)
if not defined PYEXE ( echo Kein Python gefunden. Python 3.11+ installieren. & pause & exit /b 1 )

REM Skip Streamlit's first-run email prompt (otherwise the console blocks here)
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
)

REM Ensure web deps
%PYEXE% -c "import streamlit, plotly" >nul 2>&1 || %PYEXE% -m pip install --user streamlit plotly

echo.
echo ============================================================
echo  Web-App startet. Browser oeffnet sich automatisch.
echo  Falls nicht: http://localhost:8501  manuell oeffnen.
echo  Beenden: dieses Fenster schliessen.
echo ============================================================
echo.
%PYEXE% -m streamlit run app.py --server.headless false
pause
