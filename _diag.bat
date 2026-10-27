@echo off
cd /d "%~dp0"
set "PYEXE=C:\Python314\python.exe"
set "LOG=%~dp0_diag.log"

(
    echo === ENV ===
    echo PYTHONNOUSERSITE=%PYTHONNOUSERSITE%
    echo PYTHONPATH=%PYTHONPATH%
    echo PYTHONHOME=%PYTHONHOME%
    echo VIRTUAL_ENV=%VIRTUAL_ENV%
    echo PATH=%PATH%
    echo.
    echo === PYTHON ===
    "%PYEXE%" -c "import sys, site; print('exe:', sys.executable); print('user site enabled:', site.ENABLE_USER_SITE); print('user site:', site.getusersitepackages()); print('sys.path:'); [print(' ', p) for p in sys.path]"
    echo.
    echo === IMPORTS ===
    "%PYEXE%" -c "import pandas; print('pandas:', pandas.__file__)"
    "%PYEXE%" -c "import yfinance; print('yfinance:', yfinance.__file__)" 2^>^&1
    echo.
    echo === PIP LIST ^(top of list^) ===
    "%PYEXE%" -m pip list 2^>^&1 | findstr /i "yfinance pandas openpyxl curl"
) > "%LOG%" 2>&1

type "%LOG%"
echo.
echo Log: %LOG%
pause
