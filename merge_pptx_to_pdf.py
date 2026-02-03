#!/usr/bin/env python3
"""
Script to merge all files (PPTX/PPT, images, PDFs) from input folder into a single PDF in output folder.
Supports: PowerPoint files, images (PNG, JPG, JPEG, GIF, BMP, TIFF), and existing PDFs.
Uses LibreOffice for presentation conversion and PIL for image conversion.
"""

import os
import subprocess
import tempfile
from pathlib import Path
from PyPDF2 import PdfMerger
import shutil
import sys
import platform
from PIL import Image

def has_soffice():
    """Check if soffice (LibreOffice/OpenOffice) is installed."""
    # Try various common paths
    possible_paths = [
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice"
    ]
    for path in possible_paths:
        if shutil.which(path) or Path(path).exists():
            return path
    return None

def convert_presentation_to_pdf(presentation_file, temp_dir, soffice_path):
    """Convert a presentation file to PDF using soffice."""
    try:
        cmd = [
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(temp_dir),
            str(presentation_file)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        
        # LibreOffice converts the file to the same name but with .pdf extension
        pdf_name = presentation_file.stem + ".pdf"
        temp_pdf = temp_dir / pdf_name
        
        if temp_pdf.exists():
            return temp_pdf
        else:
            print(f"  Warning: Expected PDF not found at {temp_pdf}")
            if result.stderr:
                print(f"  Error output: {result.stderr}")
            return None
            
    except subprocess.TimeoutExpired:
        print(f"  Error: Conversion timeout for {presentation_file.name}")
        return None
    except Exception as e:
        print(f"  Error converting {presentation_file.name}: {e}")
        return None

def convert_image_to_pdf(image_file, temp_dir):
    """Convert an image file to PDF using PIL."""
    try:
        pdf_name = image_file.stem + ".pdf"
        temp_pdf = temp_dir / pdf_name
        
        # Open and convert image
        img = Image.open(image_file)
        
        # Convert to RGB if necessary (for formats like PNG with transparency)
        if img.mode in ('RGBA', 'LA', 'P'):
            # Create white background
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            if img.mode in ('RGBA', 'LA'):
                background.paste(img, mask=img.split()[-1])
            else:
                background.paste(img)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Save as PDF
        img.save(str(temp_pdf), 'PDF', resolution=100.0)
        
        return temp_pdf
        
    except Exception as e:
        print(f"  Error converting {image_file.name}: {e}")
        return None

def copy_pdf_to_temp(pdf_file, temp_dir):
    """Copy an existing PDF to the temp directory."""
    try:
        temp_pdf = temp_dir / pdf_file.name
        shutil.copy2(pdf_file, temp_pdf)
        return temp_pdf
    except Exception as e:
        print(f"  Error copying {pdf_file.name}: {e}")
        return None

def merge_pptx_files_to_pdf(input_folder="input", output_folder="output", output_name="merged.pdf"):
    """Merge all supported files (PPTX/PPT, images, PDFs) from input folder into a single PDF."""
    
    # Define supported file extensions
    PRESENTATION_EXTS = ['.pptx', '.ppt']
    IMAGE_EXTS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']
    PDF_EXTS = ['.pdf']
    
    # Check for LibreOffice/soffice (only needed if presentations are present)
    soffice_path = has_soffice()
    
    # Convert to absolute paths
    script_dir = Path(__file__).parent
    input_path = Path(input_folder) if Path(input_folder).is_absolute() else script_dir / input_folder
    output_path = Path(output_folder) if Path(output_folder).is_absolute() else script_dir / output_folder
    
    if not input_path.exists():
        print(f"Error: Input folder '{input_folder}' does not exist")
        return False
    
    if not output_path.exists():
        output_path.mkdir(parents=True, exist_ok=True)
    
    # Find all supported files
    all_files = []
    for file in sorted(input_path.iterdir()):
        if file.is_file():
            ext = file.suffix.lower()
            if ext in PRESENTATION_EXTS or ext in IMAGE_EXTS or ext in PDF_EXTS:
                all_files.append(file)
    
    if not all_files:
        print(f"No supported files found in {input_folder}")
        print(f"Supported formats: PowerPoint ({', '.join(PRESENTATION_EXTS)}), ")
        print(f"                   Images ({', '.join(IMAGE_EXTS)}), ")
        print(f"                   PDF ({', '.join(PDF_EXTS)})")
        return False
    
    # Categorize files
    presentations = [f for f in all_files if f.suffix.lower() in PRESENTATION_EXTS]
    images = [f for f in all_files if f.suffix.lower() in IMAGE_EXTS]
    pdfs = [f for f in all_files if f.suffix.lower() in PDF_EXTS]
    
    # Check if LibreOffice is needed
    if presentations and not soffice_path:
        print("Warning: LibreOffice or OpenOffice is not installed.")
        print("PowerPoint files will be skipped. To convert them, install LibreOffice:")
        print("  macOS: brew install --cask libreoffice")
        print("  Ubuntu/Debian: sudo apt-get install libreoffice")
        print("  Windows: Download from https://www.libreoffice.org/download/")
        presentations = []  # Skip presentations if no LibreOffice
    
    print(f"\nFound {len(all_files)} file(s):")
    if presentations:
        print(f"  - {len(presentations)} PowerPoint file(s)")
    if images:
        print(f"  - {len(images)} image file(s)")
    if pdfs:
        print(f"  - {len(pdfs)} PDF file(s)")
    
    if soffice_path and presentations:
        print(f"\nUsing LibreOffice: {soffice_path}")
    
    # Create temporary directory for intermediate PDFs
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        pdf_files = []
        total_files = len(presentations) + len(images) + len(pdfs)
        current = 0
        
        # Convert presentations to PDF
        for pptx_file in presentations:
            current += 1
            print(f"\n[{current}/{total_files}] Converting presentation: {pptx_file.name}")
            pdf_file = convert_presentation_to_pdf(pptx_file, temp_path, soffice_path)
            if pdf_file:
                pdf_files.append(pdf_file)
                print(f"  ✓ Converted successfully")
            else:
                print(f"  ✗ Failed to convert")
        
        # Convert images to PDF
        for image_file in images:
            current += 1
            print(f"\n[{current}/{total_files}] Converting image: {image_file.name}")
            pdf_file = convert_image_to_pdf(image_file, temp_path)
            if pdf_file:
                pdf_files.append(pdf_file)
                print(f"  ✓ Converted successfully")
            else:
                print(f"  ✗ Failed to convert")
        
        # Copy existing PDFs
        for pdf_file in pdfs:
            current += 1
            print(f"\n[{current}/{total_files}] Adding PDF: {pdf_file.name}")
            temp_pdf = copy_pdf_to_temp(pdf_file, temp_path)
            if temp_pdf:
                pdf_files.append(temp_pdf)
                print(f"  ✓ Added successfully")
            else:
                print(f"  ✗ Failed to add")
        
        if not pdf_files:
            print("\nError: No files were successfully converted to PDF")
            return False
        
        # Merge PDFs
        print(f"\nMerging {len(pdf_files)} PDF(s)...")
        output_file = output_path / output_name
        
        try:
            merger = PdfMerger()
            for pdf_file in sorted(pdf_files, key=lambda x: x.stem):
                merger.append(str(pdf_file))
            
            merger.write(str(output_file))
            merger.close()
            
            print(f"\n✓ PDF created successfully: {output_file}")
            print(f"✓ Total: {len(pdf_files)} file(s) merged")
            return True
            
        except Exception as e:
            print(f"\nError merging PDFs: {e}")
            return False

if __name__ == "__main__":
    merge_pptx_files_to_pdf()
