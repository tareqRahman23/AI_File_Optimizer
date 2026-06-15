@echo off
setlocal
cd /d "%~dp0"

".tkvenv\Scripts\python.exe" -c "from PySide6.QtWidgets import QApplication; from markitdown import MarkItDown" >nul 2>nul
if errorlevel 1 (
    echo The app environment is missing or incomplete. Running setup...
    call install.bat
    if errorlevel 1 exit /b 1
)

".tkvenv\Scripts\python.exe" -c "from PySide6.QtWidgets import QApplication; from markitdown import MarkItDown" >nul 2>nul
if errorlevel 1 (
    echo The application environment failed its startup check.
    echo Run install.bat again and review any installation errors.
    pause
    exit /b 1
)

start "" ".tkvenv\Scripts\pythonw.exe" app.py
