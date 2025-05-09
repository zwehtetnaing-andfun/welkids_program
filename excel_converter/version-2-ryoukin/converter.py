import os
import xlwings as xw
import logging
from pathlib import Path
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import time
import calendar
import re

class ConversionError(Exception):
    """Custom exception for conversion errors"""
    pass

def setup_logging(output_dir):
    """Configure logging to track conversion process and any errors"""
    log_file = os.path.join(output_dir, f'conversion_log_{time.strftime("%Y%m%d_%H%M%S")}.txt')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return log_file

def select_folder(root, title="Select Folder"):
    """
    Open folder selection dialog with a custom title.
    Args:
        root (tk.Tk): The root Tkinter window.
        title (str): Title of the folder selection dialog.
    Returns:
        str: Selected folder path or None if cancelled.
    """
    try:
        folder_path = filedialog.askdirectory(title=title, parent=root)
        return folder_path if folder_path else None
    except Exception as e:
        logging.error(f"Error showing folder dialog: {str(e)}")
        return None
        
def create_root():
    """Create and configure the Tkinter root window."""
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window
    root.attributes("-topmost", True)  # Bring the dialog to the front
    return root

def check_excel_installation():
    """
    Check if Excel or a compatible application is installed.
    Returns:
        bool: True if Excel or compatible app is available.
    """
    system = platform.system()
    if system == "Darwin":  # macOS
        excel_path = "/Applications/Microsoft Excel.app"
        if not os.path.exists(excel_path):
            try:
                subprocess.run(['which', 'soffice'], check=True, capture_output=True)
                return True
            except subprocess.CalledProcessError:
                return False
        return True
    elif system == "Linux":
        try:
            subprocess.run(['which', 'soffice'], check=True, capture_output=True)
            return True
        except subprocess.CalledProcessError:
            return False
    else:  # Windows
        try:
            xw.apps.keys()
            return True
        except:
            return False

def validate_folders(input_dir, output_dir):
    """
    Validate input and output directories.
    
    Args:
        input_dir (str): Input directory path.
        output_dir (str): Output directory path.
        
    Returns:
        tuple: (bool, str) - (is_valid, error_message).
    """
    if not os.path.exists(input_dir):
        return False, "Input directory does not exist"
    
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception as e:
            return False, f"Could not create output directory: {str(e)}"
    
    if not os.access(output_dir, os.W_OK):
        return False, "Output directory is not writable"
    
    return True, ""

def get_xls_files(directory):
    """
    Find all .xls files in the specified directory.
    
    Args:
        directory (str): Path to search for .xls files.
        
    Returns:
        list: List of Path objects for .xls files.
    """
    xls_files = []
    skipped_files = []
    
    for file_path in Path(directory).rglob('*'):
        if file_path.is_file():
            if file_path.suffix.lower() == '.xls':
                # Check if it's not a temporary or hidden file
                if not file_path.name.startswith('~$') and not file_path.name.startswith('.'):
                    xls_files.append(file_path)
            else:
                skipped_files.append(file_path.name)
    
    if skipped_files:
        logging.info(f"Skipped non-XLS files: {', '.join(skipped_files)}")
    
    return xls_files

def check_output_file(output_path):
    """
    Check if output file already exists and handle accordingly.
    
    Args:
        output_path (Path): Intended output file path.
        
    Returns:
        Path: Final output path to use.
    """
    if output_path.exists():
        base = output_path.stem
        ext = output_path.suffix
        counter = 1
        while output_path.exists():
            output_path = output_path.with_name(f"{base}_{counter}{ext}")
            counter += 1
    return output_path

def convert_xls_to_xlsx(xls_file, output_dir):
    """
    Convert a single .xls file to .xlsx format while preserving all formatting.
    
    Args:
        xls_file (Path): Path object pointing to the .xls file.
        output_dir (str): Directory to save the converted file.
        
    Returns:
        bool: True if conversion successful, False otherwise.
    """
    app = None
    try:
        # Check if file exists and is readable
        if not xls_file.exists():
            raise ConversionError("Source file does not exist")
        if not os.access(xls_file, os.R_OK):
            raise ConversionError("Source file is not readable")
            
        # Create output filename with same name but in output directory
        xlsx_file = Path(output_dir) / xls_file.name
        xlsx_file = xlsx_file.with_suffix('.xlsx')
        
         # Match pattern for year and month (e.g., 2024年10月)
        match = re.match(r'(\d{4})年(\d{1,2})月', xls_file.stem)
        if match:
            year, month = match.groups()
            month = int(month)  # Convert to integer
            # last_day = calendar.monthrange(int(year), month)[1]  # Get last day of month
            new_filename = f"{year}年{month}月_通常_保育費請求書（月単位）.xlsx"
            xlsx_file = Path(output_dir) / new_filename
        
        # Check if output file exists and handle duplicates
        xlsx_file = check_output_file(xlsx_file)
        
        # Start Excel application in the background
        app = xw.App(visible=False)
        
        # Open workbook
        wb = app.books.open(str(xls_file.absolute()))
        
        # Save as xlsx
        wb.save(str(xlsx_file.absolute()))
        
        # Close workbook
        wb.close()
        
        logging.info(f"Successfully converted: {xls_file.name} -> {xlsx_file.name}")
        return True
    except Exception as e:
        error_msg = str(e)
        if "Cannot find running instance of Excel" in error_msg:
            error_msg = "Excel is not responding. Please check if it's running properly."
        elif "Permission denied" in error_msg:
            error_msg = "Permission denied. Please check if the file is open in another program."
            
        logging.error(f"Error converting {xls_file.name}: {error_msg}")
        return False
    finally:
        if app:
            try:
                app.quit()
            except:
                pass

def main():
    """Main function to handle the conversion process."""
    root = create_root()  # Create and configure the Tkinter root window
    try:
        # Select source folder
        source_folder = select_folder(root, "Select the source folder")
        if not source_folder:
            logging.info("No source folder selected. Exiting...")
            messagebox.showinfo(
                "No Folder Selected",
                "You must select a source folder to proceed.",
                parent=root
            )
            return

        # Find all V1 subfolders
        v1_folders = [str(p) for p in Path(source_folder).rglob("*/V1") if p.is_dir()]

        if not v1_folders:
            logging.info("No V1 folders found in the source directory.")
            messagebox.showinfo(
                "No V1 Folders Found",
                "No V1 folders found in the selected source directory.",
                parent=root
            )
            return

        logging.info(f"Found {len(v1_folders)} V1 folders. Starting conversion...")

        successful, failed = 0, 0

        for v1_folder in v1_folders:
            xls_files = get_xls_files(v1_folder)
            if not xls_files:
                logging.info(f"No .xls files found in {v1_folder}")
                continue

            for xls_file in xls_files:
                try:
                    success = convert_xls_to_xlsx(xls_file, v1_folder)  # Save in the same location
                    if success:
                        successful += 1
                        xls_file.unlink()  # Delete the original .xls file
                    else:
                        failed += 1
                except Exception as e:
                    logging.error(f"Unexpected error processing {xls_file.name}: {str(e)}")
                    failed += 1

        summary = (f"\nConversion Summary:\n"
                   f"Total successfully converted: {successful}\n"
                   f"Total failed conversions: {failed}")

        logging.info(summary)
        messagebox.showinfo("Conversion Complete", summary, parent=root)
    
    except Exception as e:
        logging.error(f"An unexpected error occurred: {str(e)}")
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}", parent=root)
    
    finally:
        if root:
            root.quit()
            root.destroy()

if __name__ == "__main__":
    main()