"""
================================================================================
                    UNIFIED FILE MANAGER SYSTEM
================================================================================
This script combines three file management functionalities into one unified system:

1. OOC FILE UPLOADER (formerly Moving_Occ1.py)
   - Monitors the Upload_ooc folder for PDF files
   - Extracts job number (IR_xxxxx) from filename
   - Moves files to corresponding job folders
   
2. JOB FOLDER MOVER (formerly auto_move_job_foldersAll.py)
   - Monitors job folders for OOC/OUT OF CHARGE trigger PDFs
   - Extracts importer name from PDF content
   - Moves entire job folder to appropriate billing folder
   
3. LOOSE FILE ORGANIZER (formerly MOVING_loose_files.py)
   - Monitors billing company folders for loose files
   - Matches files to job folders based on job number in filename
   - Moves loose files into their matching job folders

NOTE: File renaming feature is a separate script (auto_rename_on_startup.pyw)
      that runs silently on Windows startup.

================================================================================
                        CONFIGURATION SECTION
================================================================================
Modify these settings according to your folder structure:

SOURCE_BASE     : Main folder containing job folders (e.g., Y:\E SANCHIT 25-26)
UPLOAD_OOC      : Folder where OOC files are uploaded for processing
BILLING_BASE    : Base billing folder (e.g., Z:\BILLING 2025-2026)
LOG_DIR         : Folder where all logs are stored
SCAN_INTERVAL   : How often to scan folders (in seconds)

================================================================================
                        HOW IT WORKS
================================================================================

WORKFLOW 1: OOC File Upload
   Upload_ooc/file.pdf → Extract IR_xxxxx → Move to SOURCE_BASE/IRxxxxx/

WORKFLOW 2: Job Folder Move
   SOURCE_BASE/IRxxxxx/ (with OOC PDF) → Extract Importer → Move to BILLING_BASE/Company/

WORKFLOW 3: Loose File Organization
   BILLING_BASE/Company/loose_file.pdf → Extract job number → Move to Company/IRxxxxx/

================================================================================
"""

import os
import time
import shutil
import csv
import re
import threading
import sys
from datetime import datetime

try:
    import fitz  # PyMuPDF
except ImportError:
    print("ERROR: PyMuPDF not installed. Run: pip install PyMuPDF")
    sys.exit(1)

try:
    from PIL import Image, ImageTk
except ImportError:
    print("WARNING: Pillow (PIL) not installed. Logo quality may be lower. Run: pip install Pillow")
    Image = None
    ImageTk = None

import tkinter as tk
from tkinter import messagebox, ttk

# ================================================================================
#                           CONFIGURATION
# ================================================================================

# Main source folder containing job folders
SOURCE_BASE = r"Y:\E SANCHIT 25-26"

# Folder where OOC files are uploaded (inside SOURCE_BASE)
UPLOAD_OOC = os.path.join(SOURCE_BASE, "Upload_ooc")

# Billing base folder
BILLING_BASE = r"Z:\BILLING 2025-2026"

# Log directory - stores all logs for the system
LOG_DIR = r"Z:\BILLING 2025-2026\Automation Logs"

# Log files
MOVE_LOG = os.path.join(LOG_DIR, "job_move_log.csv")
REVERT_LOG = os.path.join(LOG_DIR, "revert_log.csv")  # Internal tracking file for undo capability
REVERT_HISTORY_LOG = os.path.join(LOG_DIR, "revert_history_log.csv")  # Records actual revert operations
OOC_UPLOAD_LOG = os.path.join(LOG_DIR, "ooc_upload_log.csv")
LOOSE_FILE_LOG = os.path.join(LOG_DIR, "loose_file_log.csv")


# Scan intervals (seconds)
OOC_UPLOAD_INTERVAL = 15       # How often to check Upload_ooc folder
JOB_MOVE_INTERVAL = 15         # How often to check for OOC trigger files
LOOSE_FILE_INTERVAL = 15       # How often to check for loose files

# Trigger keywords for OOC detection in PDF filenames
# ONLY files with this exact pattern will trigger job folder moves
# Files like INBOND_OOC, OUBOND_OOC, etc. will be IGNORED
# Supports both underscore and space variants
TRIGGER_KEYWORDS = [
    "OUT_OF_CHARGE_IR_",
    "OUT OF CHARGE_IR_"
]

# File extensions to monitor for loose file organization
MONITORED_EXTENSIONS = ['.pdf', '.docx', '.xlsx', '.jpg', '.png', '.zip']

# Job folder pattern (IR or ER followed by digits)
JOB_FOLDER_PATTERN = re.compile(r"^(IR|ER)\d+", re.IGNORECASE)
JOB_NUMBER_REGEX = re.compile(r"(IR|ER)[\s_-]?(\d{4,5})", re.IGNORECASE)

# Importer name to billing folder mapping
# Format: "IMPORTER NAME IN PDF" : "BILLING FOLDER NAME"
IMPORTER_MAP = {
    "ABBOTT HEALTHCARE PRIVATE LIMITED": "ABBOTT HEALTHCARE ( IMPORT )",
    "ANANDA BALAJI FOODS PRIVATE LIMITED": "ANANDA BALAJI FOODS PV LTD",
    "APOLLO HOSPITALS ENTERPRISE LIMITED": "APPOLLO HOSPITAL",
    "AVIAT NETWORKS (INDIA) PRIVATE LIMITED": "AVIAT NETWORKS (INDIA) PRIVATE LIMITED",
    "BALAJI WAFERS PRIVATE LIMITED": "BALAJI WAFERS ( IMPORT )",
    "BESTEX MM INDIA PRIVATE LIMITED": "BESTEX MM INDIA PVT LTD",
    "BLUE STAR ENGINEERING & ELECTRONICS LIMITED": "BLUE STAR",
    "BRISTOL-MYERS SQUIBB INDIA PRIVATE LIMITED": "BRISTOL MYERS",
    "EDIFICE MEDICAL TECHNOLOGIES": "EDIFICE MEDICAL SYSTEM",
    "GATEWAY TERMINALS INDIA PRIVATE LIMITED": "GATEWAY TERMINALS INDIA PVT LTD",
    "GUJARAT PIPAVAV PORT LIMITED": "GUJARAT PIPAVAV",
    "JAY MA INTERNATIONAL": "JAY MA INTERNATIONAL",
    "KANTILAL CHHOTALAL": "KANTILAL CHHOTALAL",
    "MEDLINE HEALTHCARE INDUSTRIES PRIVATE LIMITED": "MEDLINE HEALTHCARE  PVT LTD ( IMPORT )",
    "MUSASHI AUTO PARTS INDIA PRIVATE LIMITED": "MUSASHI AUTO PARTS",
    "NILKANTH AGRO TECH": "NILKANTH AGRO TECH",
    "SAKET TEX-DYE PVT.LTD.": "SAKET TEX DYE PVT LTD",
    "SANCO BESAN MILL": "Sanco Besan Mill",
    "SIDDHAYU LIFE SCIENCES PRIVATE LIMITED": "SIDDHAYU LIFE SCIENCES PVT ITD",
    "SVAAR PROCESS SOLUTIONS PRIVATE LIMITED": "SVAAR PROCESS SOLUTIONS PRIVATE LIMITED",
    "URSCHEL INDIA TRADING PRIVATE LIMITED": "URSCHEL INDIA",
    "ADVICS INDIA PRIVATE LIMITED": "ADVICS INDIA ( IMPORT )",
    "ANSELL INDIA PROTECTIVE PRODUCTS PRIVATE LIMITED": "ANSELL INDIA PVT LTD ( IMPORT )",
    "ARJOHUNTLEIGH HEALTHCARE INDIA PRIVATE LIMITED": "ARJOHUNTLEIGH HEALTHCARE INDIA PVT LTD",
    "B.BRAUN MEDICAL (INDIA) PRIVATE LIMITED": "B BRAUN MEDICAL ( IMPORT )",
    "BMW INDIA PRIVATE LIMITED": "BMW INDIA PVT LTD",
    "ESSITY INDIA PRIVATE LIMITED": "ESSITY INDIA ( IMPORT )",
    "MAXHILL TECHNOLOGIES": "MAXHILL TECHNOLOGIES",
    "SKODA AUTO VOLKSWAGEN INDIA PRIVATE LIMITED": "SKODA AUTO ( IMPORT )",
    "TESLA INDIA MOTORS AND ENERGY PRIVATE LIMITED": "TESLA INDIA",
    "BHARTI AIRTEL LIMITED": "BHARTI AIRTEL",
    "DUCATI INDIA PRIVATE LIMITED": "DUCATI INDIA PVT LTD",
    "BHARTI HEXACOM LIMITED": "BHARTI HEXACOM LIMITED"
}

# ================================================================================
#                           INITIALIZATION
# ================================================================================

def ensure_directories_and_logs():
    """Create necessary directories and initialize log files if they don't exist"""
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(UPLOAD_OOC, exist_ok=True)
    
    # Initialize job move log
    if not os.path.exists(MOVE_LOG):
        with open(MOVE_LOG, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "Timestamp", "Job No", "Importer", "Billing Folder",
                "Trigger File", "Action", "Comments"
            ])
    
    # Initialize revert log
    if not os.path.exists(REVERT_LOG):
        with open(REVERT_LOG, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "Job No", "Original Path", "Moved Path", "Timestamp"
            ])
    
    # Initialize OOC upload log
    if not os.path.exists(OOC_UPLOAD_LOG):
        with open(OOC_UPLOAD_LOG, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "Timestamp", "Filename", "Job No", "Destination Path", "Status"
            ])
    
    # Initialize loose file log
    if not os.path.exists(LOOSE_FILE_LOG):
        with open(LOOSE_FILE_LOG, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "Timestamp", "Company Folder", "Filename", "Source Path",
                "Destination Path", "Status"
            ])

# ================================================================================
#                           UTILITY FUNCTIONS
# ================================================================================

def extract_text_from_pdf(pdf_path):
    """Extract text from first page of PDF using PyMuPDF"""
    try:
        doc = fitz.open(pdf_path)
        text = doc[0].get_text()
        doc.close()
        return text.upper()
    except Exception as e:
        print(f"[ERROR] Failed to extract text from PDF: {e}")
        return ""


def find_importer(text):
    """Find importer name in PDF text"""
    for key in IMPORTER_MAP.keys():
        if key in text:
            return key
    return None


def is_trigger_file(filename):
    """Check if filename STARTS WITH trigger pattern (Out_of_Charge_IR_)"""
    name = filename.upper()
    return any(name.startswith(k) for k in TRIGGER_KEYWORDS)


def extract_job_number(text):
    """Extract job number (IRxxxxx or ERxxxxx) from text"""
    match = JOB_NUMBER_REGEX.search(text)
    if match:
        prefix = match.group(1).upper()
        number = match.group(2)
        return f"{prefix}{number}"
    return None


def unique_filename(dest_path):
    """Generate unique filename if file already exists"""
    if not os.path.exists(dest_path):
        return dest_path
    
    base, ext = os.path.splitext(dest_path)
    counter = 1
    
    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def is_job_folder(folder_name):
    """Check if folder name matches job folder pattern"""
    return JOB_FOLDER_PATTERN.match(folder_name) is not None


def find_matching_job_folder(company_path, job_number):
    """Find job folder that matches the job number within a company folder"""
    try:
        for item in os.listdir(company_path):
            item_path = os.path.join(company_path, item)
            if os.path.isdir(item_path):
                if job_number.upper() in item.upper():
                    return item_path
        return None
    except Exception as e:
        print(f"[ERROR] Error searching for job folder: {e}")
        return None



# ================================================================================
#                       WATCHER 1: OOC FILE UPLOADER
# ================================================================================

ooc_upload_running = False
ooc_upload_stats = {"moved": 0, "skipped": 0, "errors": 0}

def log_ooc_upload(filename, job_no, dest_path, status):
    """Log OOC upload operation"""
    try:
        with open(OOC_UPLOAD_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                filename, job_no, dest_path, status
            ])
    except Exception as e:
        print(f"[ERROR] Failed to write OOC upload log: {e}")


def watcher_ooc_upload():
    """
    Watches Upload_ooc folder for PDF files
    Extracts job number (IR_xxxxx) from filename
    Moves file to corresponding job folder
    """
    global ooc_upload_running, ooc_upload_stats
    print(">>> OOC Upload Watcher Started <<<")
    
    if not os.path.exists(UPLOAD_OOC):
        print(f"[ERROR] Upload_ooc folder does not exist: {UPLOAD_OOC}")
        ooc_upload_running = False
        return
    
    # Pattern to extract IR_xxxxx from filename
    ir_pattern = re.compile(r"IR[_\s-]?(\d{5})", re.IGNORECASE)
    
    while ooc_upload_running:
        try:
            for filename in os.listdir(UPLOAD_OOC):
                file_path = os.path.join(UPLOAD_OOC, filename)
                
                # Skip directories
                if os.path.isdir(file_path):
                    continue
                
                # Only process PDF files
                if not filename.lower().endswith(".pdf"):
                    continue
                
                # Extract job number from filename
                match = ir_pattern.search(filename)
                if not match:
                    print(f"[SKIP] No IR_xxxxx pattern in: {filename}")
                    ooc_upload_stats["skipped"] += 1
                    continue
                
                # Format as IRxxxxx
                job_no = f"IR{match.group(1)}"
                job_folder = os.path.join(SOURCE_BASE, job_no)
                
                print(f"[DEBUG] Extracted job: {job_no} from: {filename}")
                
                # Check if job folder exists
                if not os.path.exists(job_folder):
                    # Try to find folder that contains the job number
                    found = False
                    for folder in os.listdir(SOURCE_BASE):
                        folder_path = os.path.join(SOURCE_BASE, folder)
                        if os.path.isdir(folder_path) and job_no.upper() in folder.upper():
                            job_folder = folder_path
                            found = True
                            break
                    
                    if not found:
                        print(f"[ERROR] Job folder not found: {job_no}")
                        log_ooc_upload(filename, job_no, "", "ERROR: Folder not found")
                        ooc_upload_stats["errors"] += 1
                        continue
                
                # Move file to job folder
                dest_path = os.path.join(job_folder, filename)
                dest_path = unique_filename(dest_path)
                
                try:
                    shutil.move(file_path, dest_path)
                    log_ooc_upload(filename, job_no, dest_path, "MOVED")
                    ooc_upload_stats["moved"] += 1
                    print(f"[MOVED] {filename} → {job_folder}")
                except Exception as e:
                    print(f"[ERROR] Failed to move {filename}: {e}")
                    log_ooc_upload(filename, job_no, "", f"ERROR: {e}")
                    ooc_upload_stats["errors"] += 1
        
        except Exception as e:
            print(f"[ERROR] OOC Upload watcher error: {e}")
            ooc_upload_stats["errors"] += 1
        
        time.sleep(OOC_UPLOAD_INTERVAL)
    
    print(">>> OOC Upload Watcher Stopped <<<")

# ================================================================================
#                       WATCHER 2: JOB FOLDER MOVER
# ================================================================================

job_move_running = False
job_move_stats = {"moved": 0, "skipped": 0, "errors": 0}

def log_job_move(job_no, importer, billing_folder, trigger_file, action, comments):
    """Log job move operation"""
    try:
        with open(MOVE_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                job_no, importer, billing_folder, trigger_file, action, comments
            ])
    except Exception as e:
        print(f"[ERROR] Failed to write job move log: {e}")


def log_revert(job_no, original_path, moved_path):
    """Log revert information for undo capability"""
    try:
        with open(REVERT_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                job_no, original_path, moved_path,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ])
    except Exception as e:
        print(f"[ERROR] Failed to write revert log: {e}")


def move_folder(src, dest_parent):
    """Move folder to destination, handling duplicates"""
    job = os.path.basename(src)
    dest = os.path.join(dest_parent, job)
    
    os.makedirs(dest_parent, exist_ok=True)
    
    if os.path.exists(dest):
        dest = dest + "_" + datetime.now().strftime("%Y%m%d%H%M%S")
    
    shutil.move(src, dest)
    log_revert(job, src, dest)
    
    return dest


def log_revert_history(job_no, reverted_from, reverted_to, status):
    """Log successful revert operation to history log"""
    try:
        # Create file with header if it doesn't exist
        if not os.path.exists(REVERT_HISTORY_LOG):
            with open(REVERT_HISTORY_LOG, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(["Timestamp", "Job", "Reverted From", "Reverted To", "Status"])
        
        with open(REVERT_HISTORY_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                job_no, reverted_from, reverted_to, status
            ])
    except Exception as e:
        print(f"[ERROR] Failed to write revert history log: {e}")


def revert_job(job_no):
    """Revert a moved job folder to original location"""
    try:
        # Read the revert log
        if not os.path.exists(REVERT_LOG):
            print(f"[REVERT] ERROR: Revert log does not exist: {REVERT_LOG}")
            return "Revert log file not found."
        
        with open(REVERT_LOG, "r", encoding="utf-8") as f:
            all_rows = list(csv.reader(f))
        
        # Check if there's a header
        if len(all_rows) == 0:
            return "Revert log is empty."
        
        # Skip header (first row)
        rows = all_rows[1:]
        
        # Find all matches for this job number
        matches = [r for r in rows if r and len(r) >= 3 and r[0] == job_no]
        
        if not matches:
            print(f"[REVERT] No entry found for job: {job_no}")
            return f"No revert entry found for {job_no}."
        
        # Get the last (most recent) entry
        last = matches[-1]
        original = last[1]  # Original path (where it was moved FROM)
        moved = last[2]     # Moved path (where it was moved TO)
        
        print(f"[REVERT] Job: {job_no}")
        print(f"[REVERT] Original path: {original}")
        print(f"[REVERT] Moved path: {moved}")
        
        # Check if folder exists at moved location
        if not os.path.exists(moved):
            print(f"[REVERT] ERROR: Folder not found at moved location: {moved}")
            log_revert_history(job_no, moved, original, "FAILED: Folder not at moved location")
            return f"Folder no longer exists at: {moved}"
        
        # Check if original parent directory exists
        original_parent = os.path.dirname(original)
        if not os.path.exists(original_parent):
            print(f"[REVERT] ERROR: Original parent directory doesn't exist: {original_parent}")
            log_revert_history(job_no, moved, original, "FAILED: Original directory missing")
            return f"Original directory doesn't exist: {original_parent}"
        
        # Store original destination for logging
        final_destination = original
        
        # Check if something already exists at original location
        if os.path.exists(original):
            print(f"[REVERT] WARNING: Folder already exists at original location: {original}")
            # Move with timestamp suffix
            final_destination = original + "_" + datetime.now().strftime("%Y%m%d%H%M%S")
            print(f"[REVERT] Using alternate name: {final_destination}")
        
        # Perform the move
        print(f"[REVERT] Moving: {moved} -> {final_destination}")
        shutil.move(moved, final_destination)
        print(f"[REVERT] SUCCESS: Reverted {job_no} to {final_destination}")
        
        # Log successful revert to history
        log_revert_history(job_no, moved, final_destination, "SUCCESS")
        
        return f"Reverted successfully to {final_destination}"
    
    except PermissionError as e:
        print(f"[REVERT] PERMISSION ERROR: {e}")
        log_revert_history(job_no, "Unknown", "Unknown", f"FAILED: Permission denied - {e}")
        return f"Permission denied: {e}"
    except Exception as e:
        print(f"[REVERT] ERROR: {e}")
        log_revert_history(job_no, "Unknown", "Unknown", f"FAILED: {e}")
        return f"Error reverting: {e}"


def watcher_job_move():
    """
    Watches job folders for OOC trigger files
    Uses DIRECT FILE SEARCH for speed (instead of scanning all folders)
    Extracts importer from PDF content
    Moves entire job folder to billing folder
    """
    global job_move_running, job_move_stats
    print(">>> Job Folder Mover Watcher Started <<<")
    print("[INFO] Using optimized direct file search method")
    
    if not os.path.exists(SOURCE_BASE):
        print(f"[ERROR] Source path does not exist: {SOURCE_BASE}")
        job_move_running = False
        return
    
    if not os.path.exists(BILLING_BASE):
        print(f"[ERROR] Billing path does not exist: {BILLING_BASE}")
        job_move_running = False
        return
    
    # Build glob patterns for trigger files (both underscore and space variants)
    trigger_patterns = [
        os.path.join(SOURCE_BASE, "*", "Out_of_Charge_IR_*.pdf"),
        os.path.join(SOURCE_BASE, "*", "Out of Charge_IR_*.pdf"),
    ]
    
    import glob
    
    while job_move_running:
        try:
            # Direct search for trigger files (FAST!)
            trigger_files = []
            for pattern in trigger_patterns:
                trigger_files.extend(glob.glob(pattern))
            
            for trigger_path in trigger_files:
                try:
                    # Get job folder path (parent of trigger file)
                    job_path = os.path.dirname(trigger_path)
                    job = os.path.basename(job_path)
                    trigger = os.path.basename(trigger_path)
                    
                    # Skip Upload_ooc folder
                    if job.lower() == "upload_ooc":
                        continue
                    
                    # Skip if folder no longer exists (already moved)
                    if not os.path.exists(job_path):
                        continue
                    
                    print(f"[FOUND] Trigger file: {trigger} in {job}")
                    
                    # Extract text from PDF
                    text = extract_text_from_pdf(trigger_path)
                    
                    if not text:
                        print(f"[SKIP] Could not extract text from: {trigger}")
                        log_job_move(job, "", "", trigger, "SKIPPED", 
                                   "Could not extract text from PDF")
                        job_move_stats["skipped"] += 1
                        continue
                    
                    # Find importer
                    importer = find_importer(text)
                    
                    if not importer:
                        print(f"[SKIP] No matching importer found in: {trigger}")
                        log_job_move(job, "", "", trigger, "SKIPPED", 
                                   "No matching importer found in PDF")
                        job_move_stats["skipped"] += 1
                        continue
                    
                    # Get billing folder name
                    billing_folder = IMPORTER_MAP.get(importer)
                    if not billing_folder:
                        print(f"[SKIP] No billing folder mapping for: {importer}")
                        log_job_move(job, importer, "", trigger, "SKIPPED", 
                                   "No billing folder mapping found")
                        job_move_stats["skipped"] += 1
                        continue
                    
                    # Move folder to billing
                    dest_parent = os.path.join(BILLING_BASE, billing_folder)
                    
                    try:
                        final_path = move_folder(job_path, dest_parent)
                        log_job_move(job, importer, billing_folder, trigger, "MOVED",
                                   f"Moved to {final_path}")
                        job_move_stats["moved"] += 1
                        print(f"[MOVED] {job} → {billing_folder}")
                    
                    except PermissionError as e:
                        print(f"[ERROR] Permission denied moving {job}: {e}")
                        log_job_move(job, importer, billing_folder, trigger, "ERROR",
                                   f"Permission denied: {e}")
                        job_move_stats["errors"] += 1
                    
                    except Exception as e:
                        print(f"[ERROR] Failed to move {job}: {e}")
                        log_job_move(job, importer, billing_folder, trigger, "ERROR",
                                   f"Move failed: {e}")
                        job_move_stats["errors"] += 1
                
                except Exception as e:
                    print(f"[ERROR] Error processing {trigger_path}: {e}")
                    log_job_move(os.path.basename(os.path.dirname(trigger_path)), 
                               "", "", os.path.basename(trigger_path), "ERROR",
                               f"Processing error: {e}")
                    job_move_stats["errors"] += 1
        
        except Exception as e:
            print(f"[ERROR] Job mover watcher error: {e}")
            job_move_stats["errors"] += 1
        
        time.sleep(JOB_MOVE_INTERVAL)
    
    print(">>> Job Folder Mover Watcher Stopped <<<")

# ================================================================================
#                       WATCHER 3: LOOSE FILE ORGANIZER
# ================================================================================

loose_file_running = False
loose_file_stats = {"moved": 0, "skipped": 0, "errors": 0}

def log_loose_file(company, filename, source, dest, status):
    """Log loose file organization"""
    try:
        with open(LOOSE_FILE_LOG, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                company, filename, source, dest, status
            ])
    except Exception as e:
        print(f"[ERROR] Failed to write loose file log: {e}")


def watcher_loose_files():
    """
    Watches billing company folders for loose files
    Matches files to job folders based on job number in filename
    Moves loose files into their matching job folders
    """
    global loose_file_running, loose_file_stats
    print(">>> Loose File Organizer Watcher Started <<<")
    
    if not os.path.exists(BILLING_BASE):
        print(f"[ERROR] Billing path does not exist: {BILLING_BASE}")
        loose_file_running = False
        return
    
    while loose_file_running:
        try:
            for company_folder in os.listdir(BILLING_BASE):
                company_path = os.path.join(BILLING_BASE, company_folder)
                
                # Skip if not a directory or if it's a system/excluded folder
                if not os.path.isdir(company_path):
                    continue
                
                # Folders to skip (don't scan for loose files)
                skip_folders = [
                    "BILL RECEVING COPY",
                    "BHARTI AIRTEL  AIFTA",
                    "Automation Logs",
                    "AUTO_SCRIPT",
                    "File_Organization_Logs"
                ]
                if company_folder in skip_folders:
                    continue
                
                # Look for loose files
                for item in os.listdir(company_path):
                    item_path = os.path.join(company_path, item)
                    
                    # Skip directories
                    if os.path.isdir(item_path):
                        continue
                    
                    # Check file extension
                    _, ext = os.path.splitext(item)
                    if ext.lower() not in MONITORED_EXTENSIONS:
                        continue
                    
                    # Extract job number from filename
                    job_number = extract_job_number(item)
                    
                    if not job_number:
                        continue
                    
                    # Find matching job folder
                    job_folder_path = find_matching_job_folder(company_path, job_number)
                    
                    if not job_folder_path:
                        print(f"[KEEP] No matching folder for {job_number} in {company_folder}")
                        continue
                    
                    # Move file to job folder
                    dest_path = unique_filename(os.path.join(job_folder_path, item))
                    
                    try:
                        shutil.move(item_path, dest_path)
                        log_loose_file(company_folder, item, item_path, dest_path, "MOVED")
                        loose_file_stats["moved"] += 1
                        print(f"[MOVED] {item} → {os.path.basename(job_folder_path)}")
                    except Exception as e:
                        print(f"[ERROR] Failed to move {item}: {e}")
                        log_loose_file(company_folder, item, item_path, "", f"ERROR: {e}")
                        loose_file_stats["errors"] += 1
        
        except Exception as e:
            print(f"[ERROR] Loose file watcher error: {e}")
            loose_file_stats["errors"] += 1
        
        time.sleep(LOOSE_FILE_INTERVAL)
    
    print(">>> Loose File Organizer Watcher Stopped <<<")


# ================================================================================
#                           GUI: LOG VIEWER
# ================================================================================

def open_log_viewer(log_file, title, columns):
    """Generic log viewer window"""
    win = tk.Toplevel()
    win.title(title)
    win.geometry("1100x600")
    
    # Search bar
    search_frame = tk.Frame(win)
    search_frame.pack(fill="x", pady=5)
    
    tk.Label(search_frame, text="Search:").pack(side="left", padx=5)
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, width=50)
    search_entry.pack(side="left", padx=5)
    
    # Table frame
    table_frame = tk.Frame(win)
    table_frame.pack(fill="both", expand=True)
    
    scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")
    
    h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal")
    h_scrollbar.pack(side="bottom", fill="x")
    
    table = ttk.Treeview(table_frame, columns=columns, show="headings",
                         yscrollcommand=scrollbar.set,
                         xscrollcommand=h_scrollbar.set)
    
    scrollbar.config(command=table.yview)
    h_scrollbar.config(command=table.xview)
    
    for col in columns:
        table.heading(col, text=col)
        table.column(col, width=150)
    
    table.pack(side="left", fill="both", expand=True)
    
    def load_logs():
        table.delete(*table.get_children())
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                next(reader)  # Skip header
                rows = list(reader)
                rows.reverse()  # Newest first
                
                search_term = search_var.get().upper()
                
                for row in rows:
                    if search_term and search_term not in " ".join(row).upper():
                        continue
                    table.insert("", "end", values=tuple(row))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load log: {e}")
    
    load_logs()
    search_var.trace("w", lambda *args: load_logs())
    
    # Buttons
    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=10)
    
    tk.Button(btn_frame, text="Refresh", command=load_logs,
              bg="#444", fg="white", width=15).pack(side="left", padx=5)
    
    tk.Button(btn_frame, text="Open File", 
              command=lambda: os.startfile(log_file),
              bg="blue", fg="white", width=15).pack(side="left", padx=5)

# ================================================================================
#                           GUI: REVERT MANAGER
# ================================================================================

def open_revert_gui():
    """Open revert manager window"""
    win = tk.Toplevel()
    win.title("Revert Manager")
    win.geometry("1000x600")
    
    # Search and filter
    search_frame = tk.Frame(win)
    search_frame.pack(fill="x", pady=5)
    
    tk.Label(search_frame, text="Search:").pack(side="left", padx=5)
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, width=40)
    search_entry.pack(side="left", padx=5)
    
    tk.Label(search_frame, text="Filter:").pack(side="left", padx=(20, 5))
    filter_var = tk.StringVar(value="All")
    filter_dropdown = ttk.Combobox(search_frame, textvariable=filter_var,
                                   values=["All", "Not Reverted", "Already Reverted"],
                                   state="readonly", width=15)
    filter_dropdown.pack(side="left", padx=5)
    
    # Table
    table_frame = tk.Frame(win)
    table_frame.pack(fill="both", expand=True)
    
    scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")
    
    columns = ("☑", "Job", "Original", "Moved", "Time")
    table = ttk.Treeview(table_frame, columns=columns, show="tree headings",
                         selectmode="extended", yscrollcommand=scrollbar.set)
    scrollbar.config(command=table.yview)
    
    table.heading("#0", text="")
    table.column("#0", width=0, stretch=False)
    table.heading("☑", text="☑")
    table.column("☑", width=40, anchor="center")
    
    for c in columns[1:]:
        table.heading(c, text=c)
        table.column(c, width=220)
    
    table.pack(side="left", fill="both", expand=True)
    
    # Color tags
    table.tag_configure('reverted', background='#90EE90')
    table.tag_configure('not_reverted', background='#FFB6C6')
    
    checkbox_states = {}
    
    def check_status(row):
        if not row or len(row) < 3:
            return None
        if os.path.exists(row[1]):
            return True  # Reverted
        elif os.path.exists(row[2]):
            return False  # Not reverted
        return None
    
    def load_rows():
        table.delete(*table.get_children())
        checkbox_states.clear()
        
        try:
            with open(REVERT_LOG, "r", encoding="utf-8") as f:
                all_rows = list(csv.reader(f))[1:]
        except:
            all_rows = []
        
        search_term = search_var.get().upper()
        filter_option = filter_var.get()
        
        for row in all_rows:
            if not row:
                continue
            
            # Apply search
            if search_term and search_term not in " ".join(row).upper():
                continue
            
            # Apply filter
            status = check_status(row)
            if filter_option == "Not Reverted" and status is not False:
                continue
            if filter_option == "Already Reverted" and status is not True:
                continue
            
            item_id = table.insert("", "end", values=("☐",) + tuple(row))
            checkbox_states[item_id] = False
            
            if status is True:
                table.item(item_id, tags=('reverted',))
            elif status is False:
                table.item(item_id, tags=('not_reverted',))
    
    load_rows()
    
    def toggle_checkbox(event):
        item = table.identify_row(event.y)
        column = table.identify_column(event.x)
        
        if item and column == "#1":
            current = checkbox_states.get(item, False)
            checkbox_states[item] = not current
            values = list(table.item(item)["values"])
            values[0] = "☑" if not current else "☐"
            table.item(item, values=values)
    
    table.bind("<Button-1>", toggle_checkbox)
    search_var.trace("w", lambda *args: load_rows())
    filter_var.trace("w", lambda *args: load_rows())
    
    def do_revert():
        checked = [i for i, c in checkbox_states.items() if c]
        if not checked:
            messagebox.showwarning("No Selection", "Select at least one job to revert.")
            return
        
        if not messagebox.askyesno("Confirm", f"Revert {len(checked)} job(s)?"):
            return
        
        success = 0
        failures = []
        
        for item in checked:
            job_no = table.item(item)["values"][1]
            print(f"[DO_REVERT] Attempting to revert: {job_no}")
            result = revert_job(job_no)
            print(f"[DO_REVERT] Result for {job_no}: {result}")
            
            if "successfully" in result.lower():
                success += 1
            else:
                failures.append(f"{job_no}: {result}")
        
        # Show result message
        if failures:
            fail_msg = "\\n".join(failures[:5])  # Show up to 5 failures
            if len(failures) > 5:
                fail_msg += f"\\n... and {len(failures) - 5} more"
            messagebox.showwarning("Partial Success", 
                f"Reverted {success}/{len(checked)} jobs.\\n\\nFailed:\\n{fail_msg}")
        else:
            messagebox.showinfo("Success", f"Reverted {success}/{len(checked)} jobs successfully!")
        
        load_rows()
    
    def select_all():
        for item in table.get_children():
            checkbox_states[item] = True
            values = list(table.item(item)["values"])
            values[0] = "☑"
            table.item(item, values=values)
    
    def deselect_all():
        for item in table.get_children():
            checkbox_states[item] = False
            values = list(table.item(item)["values"])
            values[0] = "☐"
            table.item(item, values=values)
    
    # Buttons
    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", pady=10)
    
    # Legend
    legend = tk.Frame(btn_frame)
    legend.pack(side="left", padx=10)
    tk.Label(legend, text="●", fg="#90EE90", font=("Arial", 14)).pack(side="left")
    tk.Label(legend, text="Reverted").pack(side="left", padx=(2, 10))
    tk.Label(legend, text="●", fg="#FFB6C6", font=("Arial", 14)).pack(side="left")
    tk.Label(legend, text="Not Reverted").pack(side="left", padx=2)
    
    def delete_logs():
        checked = [i for i, c in checkbox_states.items() if c]
        if not checked:
            messagebox.showwarning("No Selection", "Select at least one log to delete.")
            return
        
        if not messagebox.askyesno("Confirm Delete", f"Delete {len(checked)} log entries?\n\nThis will NOT move folders, only remove the log record.", icon='warning'):
            return

        # Get job numbers to delete
        # Use job number (index 1) as unique identifier
        jobs_to_delete = {table.item(item)["values"][1] for item in checked}
        
        # Read all logs
        remaining_rows = []
        try:
            with open(REVERT_LOG, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                try:
                    header = next(reader)
                    remaining_rows.append(header)
                except StopIteration:
                    pass # Empty file
                
                for row in reader:
                    # Row structure: [Job, OriginalPath, MovedPath, Timestamp]
                    # Job is at index 0. Check if this job is in deletion list
                    if len(row) > 0 and row[0] not in jobs_to_delete:
                        remaining_rows.append(row)
            
            # Write back
            with open(REVERT_LOG, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(remaining_rows)
            
            messagebox.showinfo("Success", f"Deleted {len(checked)} log entries.")
            load_rows()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to modify log file: {e}")

    tk.Button(btn_frame, text="Select All", command=select_all,
              bg="#444", fg="white", width=12).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Deselect All", command=deselect_all,
              bg="#444", fg="white", width=12).pack(side="left", padx=5)
    
    tk.Button(btn_frame, text="Revert Selected Jobs", command=do_revert,
              bg="red", fg="white", width=20).pack(side="right", padx=10)
              
    tk.Button(btn_frame, text="Delete Selected Logs", command=delete_logs,
              bg="red", fg="white", width=20).pack(side="right", padx=5)

# ================================================================================
#                           MAIN GUI
# ================================================================================

def start_gui():
    global ooc_upload_running, job_move_running, loose_file_running
    
    # Initialize logs and directories
    ensure_directories_and_logs()
    
    # --- PRO ASSETS & COLORS ---
    COLORS = {
        "bg": "#f4f7f6",             # Light Grey Background
        "card_bg": "#ffffff",        # White Card Background
        "text_primary": "#333333",   # Dark Grey Text
        "text_secondary": "#666666", # Light Grey Text
        "primary": "#0056b3",        # Corporate Blue
        "success": "#28a745",        # Green
        "danger": "#dc3545",         # Red
        "warning": "#ffc107",        # Yellow
        "border": "#e0e0e0"          # Subtle Border
    }
    
    HOVER_COLORS = {
        COLORS["success"]: "#218838", # Darker Green
        COLORS["danger"]: "#c82333",  # Darker Red
        COLORS["warning"]: "#e0a800", # Darker Yellow
        COLORS["card_bg"]: "#f2f2f2", # Light Grey hover
        "#e3f2fd": "#bbdefb"          # Darker Light Blue
    }
    
    FONT_HEADER = ("Segoe UI", 18, "bold")
    FONT_TITLE = ("Segoe UI", 12, "bold")
    FONT_BODY = ("Segoe UI", 10)
    FONT_SMALL = ("Segoe UI", 9)

    root = tk.Tk()
    root.title("Unified File Manager System V3")
    root.geometry("950x700")
    root.configure(bg=COLORS["bg"])
    
    # --- HELPER: HOVER BUTTON ---
    def create_hover_button(parent, text, command, bg, fg, width=None, height=None, font=None, relief="flat", bd=0):
        # Determine hover color
        hover_bg = HOVER_COLORS.get(bg, bg)
        # If no specific mapping, try to darken slightly if it's a light color, or use standard behavior
        if bg == COLORS["card_bg"]: hover_bg = "#f2f2f2"
        
        btn = tk.Button(parent, text=text, command=command, bg=bg, fg=fg, 
                       width=width, height=height, font=font, relief=relief, bd=bd, cursor="hand2")
        
        def on_enter(e):
            btn['background'] = hover_bg
        def on_leave(e):
            btn['background'] = bg
            
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn

    # --- HELPER: CARD CREATOR ---
    def create_card(parent, title=None):
        card = tk.Frame(parent, bg=COLORS["card_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill="x", padx=20, pady=10)
        
        if title:
            lbl = tk.Label(card, text=title, font=FONT_TITLE, bg=COLORS["card_bg"], fg=COLORS["primary"])
            lbl.pack(anchor="w", padx=15, pady=(10, 5))
            tk.Frame(card, bg=COLORS["border"], height=1).pack(fill="x", padx=15, pady=5)
            
        content = tk.Frame(card, bg=COLORS["card_bg"])
        content.pack(fill="both", expand=True, padx=15, pady=10)
        return content

    # --- HEADER SECTION ---
    header_frame = tk.Frame(root, bg="white", height=90)
    header_frame.pack(fill="x", side="top")
    header_frame.pack_propagate(False) # Force height
    
    # 1. Logo (Left Aligned)
    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    try:
        # Look for logo in the same directory as the script/exe
        logo_path = resource_path("Nagarkot Logo.png")
        
        if os.path.exists(logo_path):
            if Image and ImageTk:
                # High-quality resize using Pillow
                pil_img = Image.open(logo_path)
                # Resize to target height of 70px (fits nicely in 90px header)
                base_height = 70
                w_percent = (base_height / float(pil_img.size[1]))
                w_size = int((float(pil_img.size[0]) * float(w_percent)))
                
                # Check purely for width constraint too (e.g. max 250px wide)
                if w_size > 250:
                     base_width = 250
                     h_percent = (base_width / float(pil_img.size[0]))
                     h_size = int((float(pil_img.size[1]) * float(h_percent)))
                     logo_img_pil = pil_img.resize((base_width, h_size), Image.Resampling.LANCZOS)
                else:
                     logo_img_pil = pil_img.resize((w_size, base_height), Image.Resampling.LANCZOS)
                
                logo_img = ImageTk.PhotoImage(logo_img_pil)
                logo_lbl = tk.Label(header_frame, image=logo_img, bg="white")
                logo_lbl.image = logo_img  # Keep reference
                logo_lbl.place(x=20, rely=0.5, anchor="w")
            else:
                # Fallback to standard PhotoImage if PIL not available
                logo_img = tk.PhotoImage(file=logo_path)
                if logo_img.height() > 70:
                    scale_factor = int(logo_img.height() / 70)
                    if scale_factor > 1:
                        logo_img = logo_img.subsample(scale_factor, scale_factor)
                logo_lbl = tk.Label(header_frame, image=logo_img, bg="white")
                logo_lbl.image = logo_img
                logo_lbl.place(x=20, rely=0.5, anchor="w")
    
    except Exception as e:
        print(f"Logo error: {e}")
        tk.Label(header_frame, text="NAGARKOT", font=("Arial", 20, "bold"), fg="#002b5c", bg="white").place(x=20, rely=0.5, anchor="w")

    # 2. Title & Subtitle (Centered)
    title_frame = tk.Frame(header_frame, bg="white")
    title_frame.place(relx=0.5, rely=0.5, anchor="center") # Dead center of header
    
    tk.Label(title_frame, text="Unified File Manager", font=FONT_HEADER, bg="white", fg=COLORS["text_primary"]).pack(anchor="center")
    tk.Label(title_frame, text="Automated OOC Upload • Job Moving • File Organization", font=FONT_SMALL, bg="white", fg=COLORS["text_secondary"]).pack(anchor="center")


    # --- MAIN CONTENT ---
    # --- MAIN CONTENT (SCROLLABLE) ---
    container = tk.Frame(root, bg=COLORS["bg"])
    container.pack(fill="both", expand=True)

    # Canvas and Scrollbar
    canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Inner Frame
    main_frame = tk.Frame(canvas, bg=COLORS["bg"])
    canvas.create_window((0, 0), window=main_frame, anchor="nw", tags="inner_frame")

    # Configure Scrolling
    def on_canvas_configure(event):
        # Allow the inner frame to expand to fill the canvas width
        canvas.itemconfig("inner_frame", width=event.width)
    
    def on_frame_configure(event):
        # Set the scroll region to encompass the inner frame
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", on_canvas_configure)
    main_frame.bind("<Configure>", on_frame_configure)

    # Mousewheel Binding
    def _on_mousewheel(event):
        if canvas.winfo_exists():
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_mousewheel(event):
        if container.winfo_exists():
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def _unbind_mousewheel(event):
        if container.winfo_exists():
             canvas.unbind_all("<MouseWheel>")

    container.bind('<Enter>', _bind_mousewheel)
    container.bind('<Leave>', _unbind_mousewheel)

    # --- SECTION 1: WATCHER CONTROLS & STATUS ---
    watcher_card = create_card(main_frame, "Watcher Monitor & Controls")
    
    # Grid Layout Configuration
    watcher_card.columnconfigure(0, weight=1) # Name
    watcher_card.columnconfigure(1, weight=1) # Status
    watcher_card.columnconfigure(2, weight=2) # Buttons
    
    # Header Row
    tk.Label(watcher_card, text="Service Name", font=("Segoe UI", 9, "bold"), fg=COLORS["text_secondary"], bg=COLORS["card_bg"]).grid(row=0, column=0, padx=10, sticky="w")
    tk.Label(watcher_card, text="Status", font=("Segoe UI", 9, "bold"), fg=COLORS["text_secondary"], bg=COLORS["card_bg"]).grid(row=0, column=1, padx=10, sticky="w")
    tk.Label(watcher_card, text="Actions", font=("Segoe UI", 9, "bold"), fg=COLORS["text_secondary"], bg=COLORS["card_bg"]).grid(row=0, column=2, padx=10, sticky="w")

    status_labels = {}

    def add_watcher_row(row_idx, name, start_cmd, stop_cmd):
        # Name
        tk.Label(watcher_card, text=name, font=FONT_BODY, bg=COLORS["card_bg"], fg=COLORS["text_primary"]).grid(row=row_idx, column=0, padx=10, pady=10, sticky="w")
        
        # Status
        status_frame = tk.Frame(watcher_card, bg=COLORS["card_bg"])
        status_frame.grid(row=row_idx, column=1, padx=10, sticky="w")
        dot = tk.Label(status_frame, text="●", fg=COLORS["danger"], bg=COLORS["card_bg"], font=("Arial", 14))
        dot.pack(side="left")
        text = tk.Label(status_frame, text="STOPPED", fg=COLORS["text_secondary"], bg=COLORS["card_bg"], font=FONT_SMALL)
        text.pack(side="left")
        status_labels[name] = (dot, text)
        
        # Buttons
        btn_frame = tk.Frame(watcher_card, bg=COLORS["card_bg"])
        btn_frame.grid(row=row_idx, column=2, padx=10, sticky="w")
        
        create_hover_button(btn_frame, "Start", start_cmd, bg=COLORS["card_bg"], fg=COLORS["success"], 
                           font=FONT_SMALL, width=8, relief="solid", bd=1).pack(side="left", padx=2)
        create_hover_button(btn_frame, "Stop", stop_cmd, bg=COLORS["card_bg"], fg=COLORS["danger"], 
                           font=FONT_SMALL, width=8, relief="solid", bd=1).pack(side="left", padx=2)

    def update_status_ui(name, is_running):
        dot, text = status_labels[name]
        if is_running:
            dot.config(fg=COLORS["success"])
            text.config(text="RUNNING", fg=COLORS["success"])
        else:
            dot.config(fg=COLORS["danger"])
            text.config(text="STOPPED", fg=COLORS["text_secondary"])

    # Define Start/Stop wrappers
    def start_ooc():
        global ooc_upload_running
        if not ooc_upload_running:
            ooc_upload_running = True
            threading.Thread(target=watcher_ooc_upload, daemon=True).start()
            update_status_ui("OOC Upload", True)
    def stop_ooc():
        global ooc_upload_running
        ooc_upload_running = False
        update_status_ui("OOC Upload", False)

    def start_job():
        global job_move_running
        if not job_move_running:
            job_move_running = True
            threading.Thread(target=watcher_job_move, daemon=True).start()
            update_status_ui("Job Mover", True)
    def stop_job():
        global job_move_running
        job_move_running = False
        update_status_ui("Job Mover", False)

    def start_loose():
        global loose_file_running
        if not loose_file_running:
            loose_file_running = True
            threading.Thread(target=watcher_loose_files, daemon=True).start()
            update_status_ui("Loose Files", True)
    def stop_loose():
        global loose_file_running
        loose_file_running = False
        update_status_ui("Loose Files", False)

    add_watcher_row(1, "OOC Upload", start_ooc, stop_ooc)
    add_watcher_row(2, "Job Mover", start_job, stop_job)
    add_watcher_row(3, "Loose Files", start_loose, stop_loose)
    
    # Global Controls
    global_frame = tk.Frame(watcher_card, bg=COLORS["card_bg"])
    global_frame.grid(row=4, column=0, columnspan=3, pady=(15, 5))
    
    def start_all(): start_ooc(); start_job(); start_loose()
    def stop_all(): stop_ooc(); stop_job(); stop_loose()
        
    create_hover_button(global_frame, "▶ Start All Services", start_all, bg=COLORS["success"], fg="white", font=FONT_BODY, width=20, relief="flat").pack(side="left", padx=10)
    create_hover_button(global_frame, "⏹ Stop All Services", stop_all, bg=COLORS["danger"], fg="white", font=FONT_BODY, width=20, relief="flat").pack(side="left", padx=10)


    # --- SECTION 2: STATISTICS ---
    stats_card = create_card(main_frame, "Live Statistics")
    stat_labels = {}
    
    def add_stat_box(parent, title, col):
        frame = tk.Frame(parent, bg="#f8f9fa", highlightbackground="#e9ecef", highlightthickness=1)
        frame.grid(row=0, column=col, padx=10, sticky="ew")
        parent.grid_columnconfigure(col, weight=1)
        
        tk.Label(frame, text=title, font=("Segoe UI", 9, "bold"), fg=COLORS["text_secondary"], bg="#f8f9fa").pack(pady=(10, 5))
        val_lbl = tk.Label(frame, text="0", font=("Segoe UI", 24), fg=COLORS["primary"], bg="#f8f9fa")
        val_lbl.pack()
        err_lbl = tk.Label(frame, text="0 errors", font=("Segoe UI", 8), fg=COLORS["danger"], bg="#f8f9fa")
        err_lbl.pack(pady=(0, 10))
        stat_labels[title] = (val_lbl, err_lbl)

    add_stat_box(stats_card, "OOC Uploads", 0)
    add_stat_box(stats_card, "Jobs Moved", 1)
    add_stat_box(stats_card, "Loose File Org", 2)
    
    def update_stats_ui():
        stat_labels["OOC Uploads"][0].config(text=str(ooc_upload_stats['moved']))
        stat_labels["OOC Uploads"][1].config(text=f"{ooc_upload_stats['errors']} errors")
        
        stat_labels["Jobs Moved"][0].config(text=str(job_move_stats['moved']))
        stat_labels["Jobs Moved"][1].config(text=f"{job_move_stats['errors']} errors")
        
        stat_labels["Loose File Org"][0].config(text=str(loose_file_stats['moved']))
        stat_labels["Loose File Org"][1].config(text=f"{loose_file_stats['errors']} errors")
        
        root.after(2000, update_stats_ui)
    
    update_stats_ui()


    # --- SECTION 3: LOGS & MANAGEMENT ---
    logs_card = create_card(main_frame, "Logs & Management")
    
    btn_style = {"bg": "#e3f2fd", "fg": COLORS["primary"], "font": FONT_BODY, "width": 20, "relief": "flat"}
    
    create_hover_button(logs_card, "📄 OOC Upload Log", 
        lambda: open_log_viewer(OOC_UPLOAD_LOG, "OOC Upload Log", ("Timestamp", "Filename", "Job No", "Destination", "Status")),
        **btn_style).grid(row=0, column=0, padx=10, pady=5)

    create_hover_button(logs_card, "📄 Job Move Log", 
        lambda: open_log_viewer(MOVE_LOG, "Job Move Log", ("Timestamp", "Job", "Importer", "Folder", "Trigger", "Action", "Comments")),
        **btn_style).grid(row=0, column=1, padx=10, pady=5)

    create_hover_button(logs_card, "📄 Loose File Log", 
        lambda: open_log_viewer(LOOSE_FILE_LOG, "Loose File Log", ("Timestamp", "Company", "Filename", "Source", "Destination", "Status")),
        **btn_style).grid(row=0, column=2, padx=10, pady=5)

    tk.Frame(logs_card, height=10, bg=COLORS["card_bg"]).grid(row=1, column=0) 

    create_hover_button(logs_card, "🕒 Revert History", 
        lambda: open_log_viewer(REVERT_HISTORY_LOG, "Revert History Log", ("Timestamp", "Job", "Reverted From", "Reverted To", "Status")),
        **btn_style).grid(row=2, column=0, padx=10, pady=5)

    create_hover_button(logs_card, "🔄 Open Revert Manager", open_revert_gui,
              bg=COLORS["warning"], fg="#000", font=("Segoe UI", 10, "bold"), relief="flat", width=25).grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

    # --- FOOTER ---
    footer_frame = tk.Frame(root, bg=COLORS["bg"])
    footer_frame.pack(side="bottom", fill="x", pady=10)
    
    path_text = f"Source: {SOURCE_BASE}  |  Billing: {BILLING_BASE}"
    tk.Label(footer_frame, text=path_text, font=("Consolas", 8), bg=COLORS["bg"], fg=COLORS["text_secondary"]).pack()
    tk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", font=("Segoe UI", 8), bg=COLORS["bg"], fg="#999").pack(pady=(2,0))

    root.mainloop()

# ================================================================================
#                           RUN PROGRAM
# ================================================================================

if __name__ == "__main__":
    start_gui()
