#!/bin/bash
# Run script for Unix-like systems (macOS, Linux)
# This script sets up the environment and runs the PPTX-PDF Merger application

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
OS_TYPE=$(uname -s)

# Color codes for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo ""
echo "================================"
echo "PPTX-PDF Merger - Setup & Run"
echo "================================"
echo ""

# Function to print colored status
print_status() {
    if [ $1 -eq 0 ]; then
        echo -e "${GREEN}✓${NC} $2"
    else
        echo -e "${RED}✗${NC} $2"
    fi
}

# Check for Python
echo "Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}ERROR: Python 3 is not installed!${NC}"
    echo "Please install Python 3.7 or higher from https://www.python.org/"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${GREEN}✓${NC} Python ${PYTHON_VERSION} found"
echo ""

# Check/create virtual environment
echo "Setting up virtual environment..."
if [ ! -d "$SCRIPT_DIR/venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$SCRIPT_DIR/venv"
    if [ $? -ne 0 ]; then
        echo -e "${RED}ERROR: Failed to create virtual environment${NC}"
        exit 1
    fi
fi

# Activate virtual environment
source "$SCRIPT_DIR/venv/bin/activate" 2>/dev/null
if [ $? -ne 0 ]; then
    echo -e "${RED}ERROR: Failed to activate virtual environment${NC}"
    exit 1
fi
echo -e "${GREEN}✓${NC} Virtual environment ready"
echo ""

# Upgrade pip
echo "Upgrading pip..."
python3 -m pip install --upgrade pip --quiet 2>/dev/null
echo -e "${GREEN}✓${NC} pip upgraded"
echo ""

# Install requirements from requirements.txt
echo "Installing Python dependencies from requirements.txt..."
if [ -f "$SCRIPT_DIR/requirements.txt" ]; then
    pip install -q -r "$SCRIPT_DIR/requirements.txt"
    if [ $? -eq 0 ]; then
        echo -e "${GREEN}✓${NC} PyPDF2 and Pillow installed"
    else
        echo -e "${RED}WARNING: Failed to install some dependencies from requirements.txt${NC}"
        echo "Attempting to install individually..."
        pip install -q PyPDF2>=3.0.0
        pip install -q Pillow>=10.0.0
    fi
else
    echo -e "${RED}WARNING: requirements.txt not found${NC}"
    echo "Installing default packages..."
    pip install -q PyPDF2>=3.0.0 Pillow>=10.0.0
fi
echo ""

# Verify tkinter
echo "Verifying tkinter..."
if python3 -c "import tkinter" 2>/dev/null; then
    echo -e "${GREEN}✓${NC} tkinter is available"
else
    echo -e "${YELLOW}⚠${NC}  tkinter is missing but appears in venv"
    echo "Attempting alternative approach..."
    if [ "$OS_TYPE" = "Darwin" ]; then
        echo "Try running: brew install python-tk"
    fi
fi
echo ""

# Check for LibreOffice (optional but recommended)
echo "Checking for LibreOffice (optional)..."
if command -v soffice &> /dev/null || [ -f "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] 2>/dev/null; then
    echo -e "${GREEN}✓${NC} LibreOffice found"
else
    echo -e "${YELLOW}⚠${NC}  LibreOffice not installed (optional, for PowerPoint conversion)"
fi
echo ""

# All checks passed
echo "================================"
echo -e "${GREEN}✓ Setup complete!${NC}"
echo "================================"
echo ""
echo "Launching PPTX-PDF Merger..."
echo ""

# Run the application
cd "$SCRIPT_DIR"
python3 pptx_pdf_merger.py "$@"

EXIT_CODE=$?

# Show any errors
if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo -e "${RED}Application exited with error code: $EXIT_CODE${NC}"
    echo "Try running: python3 check_requirements.py"
fi

exit $EXIT_CODE


