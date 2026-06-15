@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "PYTHON_CMD="

if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python313\python.exe"
)

where py >nul 2>nul
if not defined PYTHON_CMD if not errorlevel 1 (
    py -3.13 -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py -3.13"
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        for /f %%V in ('python -c "import sys; print(sys.version_info.major * 100 + sys.version_info.minor)"') do set "PYTHON_VERSION=%%V"
        if defined PYTHON_VERSION (
            if !PYTHON_VERSION! GEQ 310 if !PYTHON_VERSION! LEQ 313 set "PYTHON_CMD=python"
        )
    )
)

if not defined PYTHON_CMD goto :python_error

echo Creating the application environment...
if exist ".tkvenv" rmdir /s /q ".tkvenv"
%PYTHON_CMD% -m venv .tkvenv
if errorlevel 1 goto :install_error

echo Installing Microsoft MarkItDown and file converters...
".tkvenv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :install_error
".tkvenv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :install_error

".tkvenv\Scripts\python.exe" -c "from PySide6.QtWidgets import QApplication; from markitdown import MarkItDown"
if errorlevel 1 goto :install_error

echo.
echo Installation complete. Run run.bat to open MarkItDown Converter.
pause
exit /b 0

:python_error
echo.
echo A supported Python version was not found.
echo MarkItDown currently requires Python 3.10 through 3.13.
echo Python 3.14 is not supported by all required packages yet.
echo Install Python 3.13 from https://www.python.org/downloads/latest/python3.13/
echo and enable "Add Python to PATH" and "tcl/tk and IDLE".
pause
exit /b 1

:install_error
echo.
echo Installation did not complete. Review the error shown above.
echo Make sure the internet connection is available, then run install.bat again.
pause
exit /b 1
