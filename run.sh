#!/bin/bash
# Run script for Unix-like systems (macOS, Linux)

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Check if venv exists
if [ ! -d "$SCRIPT_DIR/venv" ]; then
    echo "Virtual environment not found. Creating one..."
    cd "$SCRIPT_DIR"
    python3 -m venv venv
    echo "Virtual environment created."
fi

# Activate virtual environment
source "$SCRIPT_DIR/venv/bin/activate"

# Check if dependencies are installed
if ! python3 -c "import PyPDF2, PIL" 2>/dev/null; then
    echo "Installing dependencies..."
    pip install --upgrade pip
    pip install -r "$SCRIPT_DIR/requirements.txt"
    echo "Dependencies installed."
fi

# Run the application
cd "$SCRIPT_DIR"
python3 pptx_pdf_merger.py "$@"
