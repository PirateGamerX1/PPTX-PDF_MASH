#!/usr/bin/env python3
"""
Requirements Checker for PPTX-PDF Merger
Verifies all system and Python dependencies are installed.
"""

import sys
import subprocess
import platform
from pathlib import Path
import os

class Colors:
    """ANSI color codes."""
    RED = '\033[0;31m'
    GREEN = '\033[0;32m'
    YELLOW = '\033[1;33m'
    BLUE = '\033[0;34m'
    NC = '\033[0m'  # No Color

def check_python_version():
    """Check Python version."""
    print(f"Checking Python version...", end=" ")
    version = sys.version_info
    if version.major == 3 and version.minor >= 7:
        print(f"{Colors.GREEN}✓{Colors.NC} Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"{Colors.RED}✗{Colors.NC} Python {version.major}.{version.minor} (requires 3.7+)")
        return False

def check_package(package_name, import_name=None):
    """Check if a Python package is installed."""
    if import_name is None:
        import_name = package_name
    
    print(f"Checking {package_name}...", end=" ")
    try:
        __import__(import_name)
        print(f"{Colors.GREEN}✓{Colors.NC}")
        return True
    except ImportError:
        # Try with pip as well
        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "show", package_name],
                capture_output=True,
                timeout=5
            )
            if result.returncode == 0:
                print(f"{Colors.GREEN}✓{Colors.NC} (installed via pip)")
                return True
        except:
            pass
        print(f"{Colors.RED}✗{Colors.NC}")
        return False

def check_soffice():
    """Check if LibreOffice (soffice) is installed."""
    print(f"Checking LibreOffice...", end=" ")
    
    os_type = platform.system()
    possible_paths = []
    
    if os_type == "Darwin":
        possible_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/local/bin/soffice",
        ]
    elif os_type == "Windows":
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    else:  # Linux
        possible_paths = [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
        ]
    
    # Check PATH
    try:
        result = subprocess.run(["soffice", "--version"], 
                              capture_output=True, 
                              timeout=5)
        if result.returncode == 0:
            print(f"{Colors.GREEN}✓{Colors.NC} (in PATH)")
            return True
    except:
        pass
    
    # Check specific paths
    for path in possible_paths:
        if Path(path).exists():
            print(f"{Colors.GREEN}✓{Colors.NC} (at {path})")
            return True
    
    print(f"{Colors.YELLOW}⚠{Colors.NC} (optional - needed for PowerPoint conversion)")
    return False

def get_install_instructions():
    """Get installation instructions for missing dependencies."""
    os_type = platform.system()
    instructions = []
    
    print(f"\n{Colors.BLUE}Installation Instructions:{Colors.NC}")
    print("=" * 70)
    
    if os_type == "Darwin":
        print(f"\n{Colors.YELLOW}macOS:{Colors.NC}")
        print("1. Install Homebrew (if not already installed):")
        print("   /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"")
        print("\n2. Install Python dependencies:")
        print("   pip install -r requirements.txt")
        print("   (or use: ./run.sh which sets up everything automatically)")
        print("\n3. Install tkinter (if missing):")
        PYTHON_SHORT = f"{sys.version_info.major}.{sys.version_info.minor}"
        print(f"   brew install python-tk@{PYTHON_SHORT}")
        print("\n4. Install LibreOffice (optional, for PowerPoint conversion):")
        print("   brew install --cask libreoffice")
        
    elif os_type == "Windows":
        print(f"\n{Colors.YELLOW}Windows:{Colors.NC}")
        print("1. Install Python dependencies:")
        print("   pip install -r requirements.txt")
        print("   (or use: run.bat which sets up everything automatically)")
        print("\n2. Install tkinter (comes with Python, but may need reinstallation):")
        print("   pip install tk")
        print("\n3. If tkinter still fails, reinstall Python with tkinter selected:")
        print("   - Download from https://www.python.org/")
        print("   - Check 'tcl/tk and IDLE' during installation")
        print("\n4. Install LibreOffice (optional, for PowerPoint conversion):")
        print("   - Download from https://www.libreoffice.org/download/")
        print("   - Or: winget install TheDocumentFoundation.LibreOffice")
        
    else:  # Linux
        print(f"\n{Colors.YELLOW}Linux:{Colors.NC}")
        print("1. Install Python dependencies:")
        print("   pip install -r requirements.txt")
        print("   (or use: ./run.sh which sets up everything automatically)")
        print("\n2. Install tkinter:")
        print("   Ubuntu/Debian: sudo apt-get install python3-tk")
        print("   Fedora: sudo dnf install python3-tkinter")
        print("   Arch: sudo pacman -S tk")
        print("\n3. Install LibreOffice (optional, for PowerPoint conversion):")
        print("   Ubuntu/Debian: sudo apt-get install libreoffice")
        print("   Fedora: sudo dnf install libreoffice")
        print("   Arch: sudo pacman -S libreoffice-fresh")
    
    print("=" * 70)

def main():
    """Check all requirements."""
    print(f"\n{Colors.BLUE}PPTX-PDF Merger - Requirements Checker{Colors.NC}")
    print("=" * 70)
    print()
    
    checks = {
        "Python Version": check_python_version(),
        "PyPDF2": check_package("PyPDF2"),
        "Pillow": check_package("Pillow", "PIL"),
        "tkinter": check_package("tkinter"),
        "LibreOffice": check_soffice(),
    }
    
    print()
    print("=" * 70)
    
    # Summary
    required_ok = all([checks["Python Version"], checks["PyPDF2"], 
                       checks["Pillow"], checks["tkinter"]])
    
    if required_ok:
        print(f"{Colors.GREEN}✓ All required packages are installed!{Colors.NC}")
        if checks["LibreOffice"]:
            print(f"{Colors.GREEN}✓ LibreOffice is also installed (PowerPoint support enabled){Colors.NC}")
        else:
            print(f"{Colors.YELLOW}⚠ LibreOffice is not installed (PowerPoint files will be skipped){Colors.NC}")
        print("=" * 70)
        print()
        print("You can now run the application with:")
        print("  macOS/Linux: ./run.sh")
        print("  Windows: run.bat")
        return 0
    else:
        print(f"{Colors.RED}✗ Some required packages are missing!{Colors.NC}")
        get_install_instructions()
        print()
        print(f"{Colors.YELLOW}Alternative: Just run ./run.sh (or run.bat on Windows){Colors.NC}")
        print("The run scripts will automatically set up and install everything!")
        print()
        return 1

if __name__ == "__main__":
    sys.exit(main())

