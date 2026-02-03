# PPTX-PDF Merger

A cross-platform Python application to merge PowerPoint presentations (PPTX/PPT), images, and PDFs into a single PDF file.

## Features

- **Multi-format support**: Merge PowerPoint files (.pptx, .ppt), images (.png, .jpg, .jpeg, .gif, .bmp, .tiff), and PDF files
- **Cross-platform**: Works on Windows, macOS, and Linux
- **Easy to use**: Simple command-line interface
- **Automatic conversion**: Converts presentations and images to PDF before merging
- **Smart dependency detection**: Checks for required software and provides installation guidance

## Requirements

### Software Dependencies
- **Python 3.7+**
- **LibreOffice**: Required for PowerPoint conversion (automatically detected)

### Python Packages
- PyPDF2 >= 3.0.0
- Pillow >= 10.0.0

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/PPTX-PDF_MASH.git
   cd PPTX-PDF_MASH
   ```

2. Create and activate a virtual environment (recommended):
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

4. Install LibreOffice (if not already installed):
   - **macOS**: `brew install --cask libreoffice`
   - **Ubuntu/Debian**: `sudo apt-get install libreoffice`
   - **Windows**: Download from [libreoffice.org](https://www.libreoffice.org/download/)

## Usage

1. Place all files you want to merge in the `input/` folder
2. Run the script:
   ```bash
   python merge_pptx_to_pdf.py
   ```
3. Find the merged PDF in the `output/` folder as `merged.pdf`

### Supported File Types
- PowerPoint: `.pptx`, `.ppt`
- Images: `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff`
- PDF: `.pdf`

## Project Structure

```
PPTX-PDF_MASH/
├── merge_pptx_to_pdf.py    # Main script
├── requirements.txt         # Python dependencies
├── input/                   # Place files to merge here
├── output/                  # Merged PDF will be saved here
└── README.md               # This file
```

## How It Works

1. The script scans the `input/` folder for supported files
2. PowerPoint files are converted to PDF using LibreOffice
3. Images are converted to PDF using Pillow (PIL)
4. All PDFs (converted and original) are merged into a single file
5. The final merged PDF is saved to the `output/` folder

## License

MIT License - Feel free to use and modify as needed.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
