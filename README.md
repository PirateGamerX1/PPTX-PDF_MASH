# PPTX-PDF Merger

A cross-platform Python application to merge PowerPoint presentations (PPTX/PPT), images, and PDFs into a single PDF file. Featuring an intuitive GUI and automatic dependency detection.

## Features

- **Multi-format support**: Merge PowerPoint files (.pptx, .ppt), images (.png, .jpg, .jpeg, .gif, .bmp, .tiff), and PDF files
- **Cross-platform**: Works on Windows, macOS, and Linux with automatic OS detection
- **Intuitive GUI**: tkinter-based graphical interface for easy file and folder selection
- **Smart dependency detection**: Checks for LibreOffice and provides installation guidance if missing
- **Automatic conversion**: Converts presentations and images to PDF before merging
- **Customizable output**: Choose custom output filename or use default "merged.pdf"
- **Easy to run**: Simple run scripts for Windows, macOS, and Linux
- **CLI mode**: Also supports command-line mode for scripting

## Requirements

### System Requirements
- **Python 3.7+**
- **LibreOffice/OpenOffice**: Required for PowerPoint conversion (automatically detected; app will guide installation if missing)

### Python Packages
- PyPDF2 >= 3.0.0
- Pillow >= 10.0.0

## Quick Start

### Windows
Double-click `run.bat` or run in terminal:
```bash
run.bat
```

### macOS / Linux
Open terminal in the project directory and run:
```bash
./run.sh
```

The script will automatically:
1. Create a Python virtual environment (if not already created)
2. Install required Python packages
3. Launch the application with GUI

## Installation (Manual)

If you prefer to set up manually:

1. Clone this repository:
   ```bash
   git clone https://github.com/PirateGamerX1/PPTX-PDF_MASH.git
   cd PPTX-PDF_MASH
   ```

2. Create a virtual environment:
   ```bash
   # On macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   
   # On Windows
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Install LibreOffice (optional, for PowerPoint support):
   - **macOS**: `brew install --cask libreoffice`
   - **Ubuntu/Debian**: `sudo apt-get install libreoffice`
   - **Fedora/RHEL**: `sudo dnf install libreoffice`
   - **Windows**: Download from [libreoffice.org](https://www.libreoffice.org/download/)

5. Run the application:
   ```bash
   python pptx_pdf_merger.py
   ```

## Usage

### GUI Mode (Default)
Simply run the application - the GUI will open automatically.

1. **Select Input Folder**: Click "Browse..." to choose the folder containing files to merge
2. **Select Output Folder**: Click "Browse..." to choose where to save the PDF
3. **Enter Output Filename** (optional): Specify a custom filename, or leave empty for "merged.pdf"
4. **Click "Merge Files"**: The application will process all files and display progress

### CLI Mode
For scripting or batch processing, use CLI mode:

```bash
python pptx_pdf_merger.py --cli
```

This will use the default `input/` and `output/` folders.

## Supported File Types

| Format | Extensions | Requirements |
|--------|-----------|--------------|
| PowerPoint | .pptx, .ppt | LibreOffice |
| Images | .png, .jpg, .jpeg, .gif, .bmp, .tiff | Pillow (included) |
| PDF | .pdf | PyPDF2 (included) |

## Project Structure

```
PPTX-PDF_MASH/
├── pptx_pdf_merger.py    # Main application (GUI + CLI modes)
├── run.sh                 # Launch script for macOS/Linux
├── run.bat                # Launch script for Windows
├── requirements.txt       # Python dependencies
├── input/                 # Place files to merge here
├── output/                # Merged PDF will be saved here
├── README.md              # This file
└── .gitignore             # Git ignore rules
```

## How It Works

1. Scans the input folder for supported files
2. Identifies file types (PowerPoint, images, PDFs)
3. Converts PowerPoint files to PDF using LibreOffice
4. Converts images to PDF using Pillow
5. Merges all PDFs in alphabetical order
6. Saves the final merged PDF to the output folder

## Cross-Platform Details

### Windows
- Uses `soffice.exe` from LibreOffice Program Files directory
- `run.bat` handles virtual environment setup
- Detects `winget` for automatic LibreOffice installation

### macOS
- Uses LibreOffice.app from Applications folder
- `run.sh` compatible with both Intel and Apple Silicon
- Detects `brew` for installation guidance

### Linux
- Uses system `soffice` command
- Supports Ubuntu/Debian, Fedora/RHEL, and Arch Linux
- `run.sh` provides distribution-specific installation commands

## Troubleshooting

### LibreOffice Not Found
The application will detect if LibreOffice is missing and provide installation instructions. You can still merge images and PDFs without LibreOffice, but PowerPoint files will be skipped.

### Virtual Environment Issues
Delete the `venv` folder and run the script again. It will automatically recreate and reinstall dependencies.

### Permission Denied (macOS/Linux)
Make sure the run script is executable:
```bash
chmod +x run.sh
```

## License

MIT License - Feel free to use and modify as needed.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Repository

GitHub: [github.com/PirateGamerX1/PPTX-PDF_MASH](https://github.com/PirateGamerX1/PPTX-PDF_MASH)
