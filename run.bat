@echo off
REM Run script for Windows
REM This script sets up the environment and runs the PPTX-PDF Merger application

setlocal enabledelayedexpansion

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0

echo ================================
echo PPTX-PDF Merger - Setup and Run
echo ================================
echo.

REM Check for Python
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed!
    echo Please install Python 3.7 or higher from https://www.python.org/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [OK] Python %PYTHON_VERSION% found
echo.

REM Check/create virtual environment
echo Checking virtual environment...
if not exist "%SCRIPT_DIR%venv" (
    echo Creating virtual environment...
    cd /d "%SCRIPT_DIR%"
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
) else (
    echo [OK] Virtual environment exists
)

REM Activate virtual environment
echo Activating virtual environment...
call "%SCRIPT_DIR%venv\Scripts\activate.bat"
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)
echo [OK] Virtual environment activated
echo.

REM Check for tkinter
echo Checking tkinter installation...
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo [WARN] tkinter is not installed
    echo tkinter comes with Python, but may need to be installed separately
    echo.
    echo Please ensure tkinter is installed:
    echo 1. Reinstall Python and check "tcl/tk and IDLE" during installation
    echo 2. Or run: python -m pip install tk
    echo.
    echo Attempting to install tk via pip...
    pip install tk
    if errorlevel 1 (
        echo ERROR: Failed to install tkinter
        echo Please reinstall Python with tkinter support
        pause
        exit /b 1
    )
    echo [OK] tkinter installed
) else (
    echo [OK] tkinter is installed
)
echo.

REM Check and install Python dependencies
echo Checking Python dependencies...

python -c "import PyPDF2" >nul 2>&1
if errorlevel 1 (
    echo Installing PyPDF2...
    python -m pip install --upgrade pip >nul 2>&1
    pip install PyPDF2>=3.0.0
    if errorlevel 1 (
        echo ERROR: Failed to install PyPDF2
        pause
        exit /b 1
    )
    echo [OK] PyPDF2 installed
) else (
    echo [OK] PyPDF2 is installed
)

python -c "import PIL" >nul 2>&1
if errorlevel 1 (
    echo Installing Pillow...
    pip install Pillow>=10.0.0
    if errorlevel 1 (
        echo ERROR: Failed to install Pillow
        pause
        exit /b 1
    )
    echo [OK] Pillow installed
) else (
    echo [OK] Pillow is installed
)
echo.

REM Check for LibreOffice (optional but recommended)
echo Checking for LibreOffice (optional)...
where soffice >nul 2>&1
if errorlevel 1 (
    REM Try common LibreOffice paths
    if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
        echo [OK] LibreOffice found at C:\Program Files\LibreOffice
    ) else if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
        echo [OK] LibreOffice found at C:\Program Files ^(x86^)\LibreOffice
    ) else (
        echo [WARN] LibreOffice is not installed (optional, needed for PowerPoint conversion^)
        echo Install from: https://www.libreoffice.org/download/
        echo Or via: winget install TheDocumentFoundation.LibreOffice
    )
) else (
    echo [OK] LibreOffice is installed
)
echo.

REM All checks passed
echo ================================
echo [OK] All requirements verified!
echo ================================
echo.
echo Launching PPTX-PDF Merger...
echo.

REM Run the application
cd /d "%SCRIPT_DIR%"
python pptx_pdf_merger.py %*

REM Keep window open if there's an error
if errorlevel 1 (
    echo.
    echo Application exited with an error. Press any key to close...
    pause >nul
)

