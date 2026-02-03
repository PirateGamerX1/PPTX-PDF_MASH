@echo off
REM Run script for Windows

setlocal enabledelayedexpansion

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0

REM Check if venv exists
if not exist "%SCRIPT_DIR%venv" (
    echo Virtual environment not found. Creating one...
    cd /d "%SCRIPT_DIR%"
    python -m venv venv
    echo Virtual environment created.
)

REM Activate virtual environment
call "%SCRIPT_DIR%venv\Scripts\activate.bat"

REM Check if dependencies are installed
python -c "import PyPDF2, PIL" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    python -m pip install --upgrade pip
    pip install -r "%SCRIPT_DIR%requirements.txt"
    echo Dependencies installed.
)

REM Run the application
cd /d "%SCRIPT_DIR%"
python pptx_pdf_merger.py %*

REM Keep window open if there's an error
if errorlevel 1 pause
