import os
import logging
from pathlib import Path

# Configure logging for better visibility
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def clone_directory_structure(source_path: str, target_path: str) -> None:
    """
    Recursively clones the directory structure from source to target without copying files.
    
    Args:
        source_path: The root directory to copy structure from.
        target_path: The directory where the structure should be recreated.
    """
    src = Path(source_path)
    dst = Path(target_path)

    if not src.exists():
        logging.error(f"Source directory does not exist: {src}")
        return

    if not src.is_dir():
        logging.error(f"Source path is not a directory: {src}")
        return

    logging.info(f"Cloning structure from: {src}")
    logging.info(f"Target destination:    {dst}")
    print("-" * 50)

    folder_count = 0
    
    # Create the root destination folder if it doesn't exist
    dst.mkdir(parents=True, exist_ok=True)

    try:
        # Loop ONLY through the immediate children of the source (Top Level)
        for item in src.iterdir():
            if item.is_dir():
                target_dir = dst / item.name
                
                # Create the directory in the new location
                if not target_dir.exists():
                    target_dir.mkdir(parents=True, exist_ok=True)
                    print(f"[+] Created Top-Level: {item.name}")
                    folder_count += 1
                else:
                    print(f"[.] Exists:           {item.name}")

        print("-" * 50)
        logging.info(f"Process complete. Created {folder_count} top-level directories.")
        
    except PermissionError as e:
        logging.error(f"Permission denied: {e}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    # CONFIGURATION: Change these paths as needed
    SOURCE_BILLING = r"Z:\BILLING 2025-2026"
    TARGET_BILLING = r"Z:\BILLING 2026-2027"

    clone_directory_structure(SOURCE_BILLING, TARGET_BILLING)
