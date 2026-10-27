@echo off
setlocal
cd /d "%~dp0"

set "PYEXE="
py -3 -c "import sys" >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE python -c "import sys" >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
    for %%D in ("C:\Python314" "C:\Python313" "C:\Python312" "C:\Python311") do (
        if not defined PYEXE if exist "%%~D\python.exe" set "PYEXE=%%~D\python.exe"
    )
)
if not defined PYEXE ( echo Kein Python gefunden. & pause & exit /b 1 )

%PYEXE% -c "import pytest" >nul 2>&1 || %PYEXE% -m pip install --user pytest

echo Starte Test-Suite...
echo.
%PYEXE% -m pytest tests/ -v
echo.
pause
