#!/usr/bin/env python3
"""
Cross-platform PPTX/PDF Merger with GUI
Merges PowerPoint presentations, images, and PDFs into a single PDF file.
Supports: PowerPoint files, images (PNG, JPG, JPEG, GIF, BMP, TIFF), and existing PDFs.
Uses LibreOffice for presentation conversion and PIL for image conversion.
"""

import os
import subprocess
import tempfile
from pathlib import Path
import shutil
import sys
import platform

# Check for required Python modules
def check_python_requirements():
    """Check and verify all required Python packages are installed."""
    missing_packages = []
    
    # Check PyPDF2
    try:
        from PyPDF2 import PdfMerger
    except ImportError:
        missing_packages.append("PyPDF2")
    
    # Check Pillow
    try:
        from PIL import Image
    except ImportError:
        missing_packages.append("Pillow")
    
    # Check tkinter (built-in, but may be missing on some systems)
    try:
        import tkinter as tk
    except ImportError:
        missing_packages.append("tkinter")
    
    return missing_packages

# Check requirements before proceeding
missing = check_python_requirements()
if missing:
    print("=" * 70)
    print("ERROR: Missing Python packages!")
    print("=" * 70)
    print(f"\nThe following packages are not installed:")
    for pkg in missing:
        print(f"  - {pkg}")
    print("\nTo fix this, please run:")
    print("  pip install -r requirements.txt")
    
    if "tkinter" in missing:
        os_type = platform.system()
        if os_type == "Darwin":
            print("\nFor tkinter on macOS, you may need to:")
            print("  1. Use homebrew: brew install python-tk@3.13 (or your Python version)")
            print("  2. Or reinstall Python with: brew install python-tk")
        elif os_type == "Linux":
            print("\nFor tkinter on Linux, run:")
            print("  Ubuntu/Debian: sudo apt-get install python3-tk")
            print("  Fedora: sudo dnf install python3-tkinter")
            print("  Arch: sudo pacman -S tk")
    
    print("\n" + "=" * 70)
    sys.exit(1)

# Now import remaining modules
from PyPDF2 import PdfMerger
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import webbrowser

def get_os_type():
    """Detect the operating system."""
    system = platform.system()
    if system == "Darwin":
        return "macos"
    elif system == "Windows":
        return "windows"
    elif system == "Linux":
        return "linux"
    else:
        return "unknown"

def get_libreoffice_install_command():
    """Get the installation command for LibreOffice based on OS."""
    os_type = get_os_type()
    
    if os_type == "macos":
        return "brew install --cask libreoffice", "https://www.libreoffice.org/download/"
    elif os_type == "windows":
        return "winget install TheDocumentFoundation.LibreOffice", "https://www.libreoffice.org/download/"
    elif os_type == "linux":
        # Try to detect the distribution
        try:
            with open('/etc/os-release', 'r') as f:
                os_release = f.read().lower()
                if 'ubuntu' in os_release or 'debian' in os_release:
                    return "sudo apt-get update && sudo apt-get install libreoffice", "https://www.libreoffice.org/download/"
                elif 'fedora' in os_release or 'rhel' in os_release or 'centos' in os_release:
                    return "sudo dnf install libreoffice", "https://www.libreoffice.org/download/"
                elif 'arch' in os_release:
                    return "sudo pacman -S libreoffice-fresh", "https://www.libreoffice.org/download/"
        except:
            pass
        return "sudo apt-get install libreoffice", "https://www.libreoffice.org/download/"
    else:
        return None, "https://www.libreoffice.org/download/"

def offer_libreoffice_installation():
    """Offer to install LibreOffice if not found."""
    install_cmd, download_url = get_libreoffice_install_command()
    os_type = get_os_type()
    
    print("\n" + "="*70)
    print("LibreOffice NOT FOUND")
    print("="*70)
    print("LibreOffice is required to convert PowerPoint files to PDF.")
    print()
    
    if install_cmd:
        print(f"To install LibreOffice on {os_type.upper()}, run:")
        print(f"  {install_cmd}")
        print()
        
        response = input("Would you like to install it now? (y/n): ").strip().lower()
        if response == 'y':
            print(f"\nAttempting to install LibreOffice...")
            try:
                if os_type == "windows":
                    # On Windows, try winget
                    result = subprocess.run(install_cmd, shell=True, capture_output=True, text=True)
                else:
                    # On Unix-like systems
                    result = subprocess.run(install_cmd, shell=True)
                
                if result.returncode == 0:
                    print("LibreOffice installed successfully!")
                    print("Please restart the application.")
                    return True
                else:
                    print("Installation failed. Please install manually.")
                    print(f"Download from: {download_url}")
                    return False
            except Exception as e:
                print(f"Installation failed: {e}")
                print(f"Please install manually from: {download_url}")
                return False
    else:
        print(f"Please install LibreOffice manually from:")
        print(f"  {download_url}")
    
    print("="*70)
    return False

def has_soffice():
    """Check if soffice (LibreOffice/OpenOffice) is installed."""
    os_type = get_os_type()
    
    # Platform-specific paths
    if os_type == "macos":
        possible_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/local/bin/soffice",
            "soffice",
            "libreoffice"
        ]
    elif os_type == "windows":
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\Program Files\LibreOffice 7\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice 7\program\soffice.exe",
            "soffice",
            "soffice.exe"
        ]
    else:  # Linux and others
        possible_paths = [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "soffice",
            "libreoffice"
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

def merge_pptx_files_to_pdf(input_folder="input", output_folder="output", output_name="merged.pdf", progress_callback=None):
    """Merge all supported files (PPTX/PPT, images, PDFs) from input folder into a single PDF."""
    
    def log(message):
        """Print and optionally call progress callback."""
        print(message)
        if progress_callback:
            progress_callback(message)
    
    # Define supported file extensions
    PRESENTATION_EXTS = ['.pptx', '.ppt']
    IMAGE_EXTS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']
    PDF_EXTS = ['.pdf']
    
    # Check for LibreOffice/soffice
    soffice_path = has_soffice()
    
    # Convert to absolute paths
    script_dir = Path(__file__).parent
    input_path = Path(input_folder) if Path(input_folder).is_absolute() else script_dir / input_folder
    output_path = Path(output_folder) if Path(output_folder).is_absolute() else script_dir / output_folder
    
    if not input_path.exists():
        log(f"Error: Input folder '{input_folder}' does not exist")
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
        log(f"No supported files found in {input_folder}")
        log(f"Supported formats: PowerPoint ({', '.join(PRESENTATION_EXTS)}), ")
        log(f"                   Images ({', '.join(IMAGE_EXTS)}), ")
        log(f"                   PDF ({', '.join(PDF_EXTS)})")
        return False
    
    # Categorize files
    presentations = [f for f in all_files if f.suffix.lower() in PRESENTATION_EXTS]
    images = [f for f in all_files if f.suffix.lower() in IMAGE_EXTS]
    pdfs = [f for f in all_files if f.suffix.lower() in PDF_EXTS]
    
    # Check if LibreOffice is needed and offer installation
    if presentations and not soffice_path:
        log("\n" + "="*70)
        log("WARNING: LibreOffice is not installed!")
        log("="*70)
        log(f"Found {len(presentations)} PowerPoint file(s) that require LibreOffice.")
        log("")
        
        if not progress_callback:  # Only offer installation in CLI mode
            if offer_libreoffice_installation():
                # Re-check for soffice after installation
                soffice_path = has_soffice()
        
        if not soffice_path:
            install_cmd, download_url = get_libreoffice_install_command()
            if install_cmd:
                log(f"To install LibreOffice, run: {install_cmd}")
            log(f"Or download from: {download_url}")
            log("")
            log("PowerPoint files will be SKIPPED.")
            log("="*70)
            presentations = []
    
    log(f"\nFound {len(all_files)} file(s):")
    if presentations:
        log(f"  - {len(presentations)} PowerPoint file(s)")
    if images:
        log(f"  - {len(images)} image file(s)")
    if pdfs:
        log(f"  - {len(pdfs)} PDF file(s)")
    
    if soffice_path and presentations:
        log(f"\nUsing LibreOffice: {soffice_path}")
    
    # Create temporary directory for intermediate PDFs
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        pdf_files = []
        total_files = len(presentations) + len(images) + len(pdfs)
        current = 0
        
        # Convert presentations to PDF
        for pptx_file in presentations:
            current += 1
            log(f"\n[{current}/{total_files}] Converting presentation: {pptx_file.name}")
            pdf_file = convert_presentation_to_pdf(pptx_file, temp_path, soffice_path)
            if pdf_file:
                pdf_files.append(pdf_file)
                log(f"  ✓ Converted successfully")
            else:
                log(f"  ✗ Failed to convert")
        
        # Convert images to PDF
        for image_file in images:
            current += 1
            log(f"\n[{current}/{total_files}] Converting image: {image_file.name}")
            pdf_file = convert_image_to_pdf(image_file, temp_path)
            if pdf_file:
                pdf_files.append(pdf_file)
                log(f"  ✓ Converted successfully")
            else:
                log(f"  ✗ Failed to convert")
        
        # Copy existing PDFs
        for pdf_file in pdfs:
            current += 1
            log(f"\n[{current}/{total_files}] Adding PDF: {pdf_file.name}")
            temp_pdf = copy_pdf_to_temp(pdf_file, temp_path)
            if temp_pdf:
                pdf_files.append(temp_pdf)
                log(f"  ✓ Added successfully")
            else:
                log(f"  ✗ Failed to add")
        
        if not pdf_files:
            log("\nError: No files were successfully converted to PDF")
            return False
        
        # Merge PDFs
        log(f"\nMerging {len(pdf_files)} PDF(s)...")
        output_file = output_path / output_name
        
        try:
            merger = PdfMerger()
            for pdf_file in sorted(pdf_files, key=lambda x: x.stem):
                merger.append(str(pdf_file))
            
            merger.write(str(output_file))
            merger.close()
            
            log(f"\n✓ PDF created successfully: {output_file}")
            log(f"✓ Total: {len(pdf_files)} file(s) merged")
            return True
            
        except Exception as e:
            log(f"\nError merging PDFs: {e}")
            return False

class PDFMergerGUI:
    """Simple GUI for PPTX/PDF Merger."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX-PDF Merger")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)
        
        # Variables
        self.input_folder = tk.StringVar(value=str(Path.cwd() / "input"))
        self.output_folder = tk.StringVar(value=str(Path.cwd() / "output"))
        self.output_filename = tk.StringVar(value="merged.pdf")
        
        self.create_widgets()
        
        # Check for LibreOffice on startup
        self.root.after(100, self.check_libreoffice)
    
    def check_libreoffice(self):
        """Check if LibreOffice is installed and show warning if not."""
        soffice_path = has_soffice()
        if not soffice_path:
            install_cmd, download_url = get_libreoffice_install_command()
            message = "LibreOffice is not installed!\n\n"
            message += "It's required to convert PowerPoint files.\n\n"
            if install_cmd:
                message += f"To install, run:\n{install_cmd}\n\n"
            message += f"Or download from:\n{download_url}\n\n"
            message += "You can still merge images and PDFs without LibreOffice."
            
            response = messagebox.askquestion(
                "LibreOffice Not Found",
                message + "\n\nWould you like to open the download page?",
                icon='warning'
            )
            if response == 'yes':
                webbrowser.open(download_url)
    
    def create_widgets(self):
        """Create the GUI widgets."""
        # Title
        title_label = tk.Label(
            self.root,
            text="PPTX-PDF Merger",
            font=("Arial", 18, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Description
        desc_label = tk.Label(
            self.root,
            text="Merge PowerPoint presentations, images, and PDFs into a single PDF file",
            font=("Arial", 10),
            fg="gray"
        )
        desc_label.pack(pady=(0, 10))
        
        # Main frame - this will grow with the window
        main_frame = tk.Frame(self.root, padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input folder selection
        input_frame = tk.LabelFrame(main_frame, text="Input Folder", padx=10, pady=10)
        input_frame.pack(fill=tk.X, pady=10)
        
        tk.Entry(input_frame, textvariable=self.input_folder, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(input_frame, text="Browse...", command=self.browse_input).pack(side=tk.LEFT)
        
        # Output folder selection
        output_frame = tk.LabelFrame(main_frame, text="Output Folder", padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=10)
        
        tk.Entry(output_frame, textvariable=self.output_folder, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(output_frame, text="Browse...", command=self.browse_output).pack(side=tk.LEFT)
        
        # Output filename
        filename_frame = tk.LabelFrame(main_frame, text="Output Filename", padx=10, pady=10)
        filename_frame.pack(fill=tk.X, pady=10)
        
        tk.Entry(filename_frame, textvariable=self.output_filename, width=50).pack(side=tk.LEFT, padx=5)
        tk.Label(filename_frame, text="(leave empty for 'merged.pdf')").pack(side=tk.LEFT)
        
        # Progress text area
        progress_frame = tk.LabelFrame(main_frame, text="Progress", padx=10, pady=10)
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.progress_text = tk.Text(progress_frame, height=10, width=65, state=tk.DISABLED, wrap=tk.WORD)
        scrollbar = tk.Scrollbar(progress_frame, command=self.progress_text.yview)
        self.progress_text.config(yscrollcommand=scrollbar.set)
        self.progress_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons frame
        button_frame = tk.Frame(self.root, pady=15, bg=self.root.cget('bg'))
        button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.merge_button = tk.Button(
            button_frame,
            text="Merge Files",
            command=self.start_merge,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=25,
            pady=12,
            cursor="hand2"
        )
        self.merge_button.pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            button_frame,
            text="Exit",
            command=self.root.quit,
            font=("Arial", 12),
            padx=25,
            pady=12,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=10)
    
    def browse_input(self):
        """Browse for input folder."""
        folder = filedialog.askdirectory(
            title="Select Input Folder",
            initialdir=self.input_folder.get()
        )
        if folder:
            self.input_folder.set(folder)
    
    def browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_folder.get()
        )
        if folder:
            self.output_folder.set(folder)
    
    def log_progress(self, message):
        """Add message to progress text area."""
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.progress_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def start_merge(self):
        """Start the merge process in a separate thread."""
        # Validate inputs
        input_path = Path(self.input_folder.get())
        if not input_path.exists():
            messagebox.showerror("Error", f"Input folder does not exist:\n{input_path}")
            return
        
        output_filename = self.output_filename.get().strip()
        if not output_filename:
            output_filename = "merged.pdf"
        elif not output_filename.endswith('.pdf'):
            output_filename += '.pdf'
        
        # Clear progress
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.config(state=tk.DISABLED)
        
        # Disable merge button
        self.merge_button.config(state=tk.DISABLED, text="Merging...")
        
        # Run merge in thread
        thread = threading.Thread(
            target=self.run_merge,
            args=(
                self.input_folder.get(),
                self.output_folder.get(),
                output_filename
            ),
            daemon=True
        )
        thread.start()
    
    def run_merge(self, input_folder, output_folder, output_filename):
        """Run the merge process."""
        try:
            success = merge_pptx_files_to_pdf(
                input_folder=input_folder,
                output_folder=output_folder,
                output_name=output_filename,
                progress_callback=self.log_progress
            )
            
            if success:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Success",
                    f"PDF created successfully!\n\nOutput: {Path(output_folder) / output_filename}"
                ))
            else:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    "Failed to create PDF. Check the progress log for details."
                ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror(
                "Error",
                f"An error occurred:\n{str(e)}"
            ))
        finally:
            self.root.after(0, lambda: self.merge_button.config(state=tk.NORMAL, text="Merge Files"))

def run_gui():
    """Launch the GUI application."""
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    # Check if GUI mode should be used
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        # CLI mode
        merge_pptx_files_to_pdf()
    else:
        # GUI mode (default)
        run_gui()
