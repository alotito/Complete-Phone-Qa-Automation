# PhoneQACopyTool.py

import sys
import os
import datetime
import shutil
import fnmatch
import re
import logging
import argparse
import traceback
import hashlib
import configparser
from typing import List, Tuple, Optional, Dict

# --- Determine Script Directory ---
try:
    script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
except NameError:
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# --- Static Configuration ---
CONFIG_FILE_NAME: str = "config.ini"
CONFIG_FILE_PATH: str = os.path.join(script_dir, CONFIG_FILE_NAME)
EXT_LIST_FILE_PATH: str = os.path.join(script_dir, "ExtList.data")

# --- Logger Setup (Global Instance) ---
logger: logging.Logger = logging.getLogger("PhoneQACopyTool")

def setup_logger(base_path_for_logs: str, is_debug: bool):
    """Configures the global logger instance."""
    try:
        log_dir = os.path.join(base_path_for_logs, "logs", "PhoneQACopyTool")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"FileCopy_{datetime.datetime.now():%Y-%m-%d_%H%M%S}.log")

        logger.setLevel(logging.DEBUG)
        if logger.hasHandlers():
            for handler in logger.handlers[:]: logger.removeHandler(handler); handler.close()

        fh = logging.FileHandler(log_file, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.DEBUG if is_debug else logging.INFO)
        ch_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        ch.setFormatter(ch_formatter)
        logger.addHandler(ch)

        logger.info(f"Logger initialized. Log file: {log_file}")
    except Exception as e:
        print(f"FATAL: Failed to configure logging to '{base_path_for_logs}\\logs': {e}", file=sys.stderr)
        sys.exit(1)

def calculate_sha256(filepath: str) -> str:
    """Calculates the SHA256 hash of a file."""
    sha256_hash = hashlib.sha256()
    with open(filepath, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def verify_extlist_integrity(config: configparser.ConfigParser) -> bool:
    """Verifies the integrity of the ExtList.data file using a stored SHA256 hash."""
    try:
        expected_hash = config.get('ExtListChecksum', 'hash', fallback=None)
        if not expected_hash:
            logger.warning("Hash for ExtList.data not found in config.ini. Integrity check SKIPPED.")
            return True

        current_hash = calculate_sha256(EXT_LIST_FILE_PATH)
        if current_hash == expected_hash:
            logger.info(f"Integrity check PASSED for '{os.path.basename(EXT_LIST_FILE_PATH)}'.")
            return True
        else:
            logger.critical(f"CRITICAL: Integrity check FAILED for ExtList.data. File may have been tampered with.")
            return False
    except FileNotFoundError:
        logger.critical(f"'{os.path.basename(EXT_LIST_FILE_PATH)}' not found at '{EXT_LIST_FILE_PATH}'. Cannot verify integrity.")
        return False
    except Exception as e:
        logger.critical(f"Error during integrity check for ExtList.data: {e}", exc_info=True)
        return False

def read_extension_list(file_path: str) -> List[str]:
    """Reads a tab-separated file, extracting the first column as the extension."""
    extensions: List[str] = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not (s_line := line.strip()) or s_line.startswith('#'): continue
                if ext := s_line.split('\t')[0].strip():
                    extensions.append(ext)
        if not extensions:
            logger.warning(f"Extension list file '{file_path}' is empty or contains no valid entries.")
        return extensions
    except Exception as e:
        logger.critical(f"Error reading extension list file '{file_path}': {e}", exc_info=True)
        sys.exit(1)

def parse_arguments() -> Tuple[Optional[datetime.date], bool]:
    """Parses command-line arguments for date override and debug mode."""
    parser = argparse.ArgumentParser(description="Copies call recordings based on config.ini settings.")
    parser.add_argument("-D", "--date", type=str, help="Specify Sunday date (YYYY-MM-DD) to define processing week.")
    parser.add_argument('--debug', action='store_true', help='Enable detailed debug console messages.')
    args = parser.parse_args()
    
    parsed_date_override = None
    if args.date:
        try:
            parsed_date_override = datetime.datetime.strptime(args.date, '%Y-%m-%d').date()
            if parsed_date_override.weekday() != 6: # Sunday is 6
                print(f"WARNING: Provided date --date '{args.date}' is not a Sunday.", file=sys.stderr)
        except ValueError:
            print(f"ERROR: Invalid date format for --date: '{args.date}'. Ignoring.", file=sys.stderr)
            
    return parsed_date_override, args.debug

def main(source_root: str, base_dest_path: str, files_to_copy_count: int, file_patterns: List[str], cli_date_override: Optional[datetime.date]):
    """Main application logic for finding and copying files."""
    script_start_time = datetime.datetime.now()
    logger.info(f"Copy Tool Started. Will copy {files_to_copy_count} largest files matching {file_patterns}.")

    effective_ref_date = cli_date_override or datetime.date.today()
    logger.info(f"Using reference date: {effective_ref_date.strftime('%Y-%m-%d')}")

    days_since_sunday = (effective_ref_date.weekday() + 1) % 7
    target_week_start = effective_ref_date - datetime.timedelta(days=days_since_sunday + 7)
    target_week_end = target_week_start + datetime.timedelta(days=6)
    
    date_str_for_folder = target_week_start.strftime('%Y-%m-%d')
    logger.info(f"Targeting files for calendar week: {target_week_start.strftime('%Y-%m-%d')} to {target_week_end.strftime('%Y-%m-%d')}")

    target_extensions = read_extension_list(EXT_LIST_FILE_PATH)
    if not target_extensions:
        logger.warning("No extensions to process. Exiting."); return

    logger.info(f"Found {len(target_extensions)} extensions to process.")
    
    week_destination_base = os.path.join(base_dest_path, f"Week of {date_str_for_folder}")
    os.makedirs(week_destination_base, exist_ok=True)
    logger.info(f"Base weekly destination folder: '{week_destination_base}'")

    total_files_copied = 0
    for i, extension in enumerate(target_extensions, 1):
        logger.info(f"--- Processing Ext ({i}/{len(target_extensions)}): {extension} ---")
        current_source_folder = os.path.join(source_root, extension)

        if not os.path.isdir(current_source_folder):
            logger.warning(f"Source folder not found, skipping: '{current_source_folder}'"); continue
        
        matching_files: List[Tuple[int, str]] = [] # (size, path)
        try:
            for entry in os.scandir(current_source_folder):
                if entry.is_file() and any(fnmatch.fnmatch(entry.name.lower(), p.lower()) for p in file_patterns):
                    mod_time = datetime.date.fromtimestamp(entry.stat().st_mtime)
                    if target_week_start <= mod_time <= target_week_end:
                        matching_files.append((entry.stat().st_size, entry.path))
        except Exception as e_scan:
            logger.error(f"Could not scan folder '{current_source_folder}': {e_scan}"); continue

        if not matching_files:
            logger.info(f"No files matching criteria found for extension {extension}."); continue

        largest_files_to_copy = sorted(matching_files, key=lambda x: x[0], reverse=True)[:files_to_copy_count]
        dest_folder_for_ext = os.path.join(week_destination_base, extension)
        os.makedirs(dest_folder_for_ext, exist_ok=True)
        
        logger.info(f"Found {len(matching_files)} matching files. Selecting {len(largest_files_to_copy)} largest for copy.")
        
        for size, source_path in largest_files_to_copy:
            try:
                dest_path = os.path.join(dest_folder_for_ext, os.path.basename(source_path))
                if os.path.exists(dest_path):
                    logger.info(f"Skipping, already exists: '{os.path.basename(dest_path)}'"); continue
                shutil.copy2(source_path, dest_path)
                logger.info(f"Copied '{os.path.basename(source_path)}'")
                total_files_copied += 1
            except Exception as e_copy:
                logger.error(f"Failed to copy '{source_path}': {e_copy}")

    logger.info("-" * 80)
    duration = datetime.datetime.now() - script_start_time
    logger.info(f"Script finished. Total files copied: {total_files_copied}. Duration: {duration}.")


if __name__ == "__main__":
    cli_date, is_debug_mode = parse_arguments()
    
    if not os.path.exists(CONFIG_FILE_PATH):
        print(f"FATAL: Configuration file '{CONFIG_FILE_PATH}' not found. Cannot continue.", file=sys.stderr)
        sys.exit(1)
        
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE_PATH)
    
    try:
        source_dir = config.get('Paths', 'CopySourceRoot')
        dest_dir = config.get('Paths', 'ImporterSourceRoot')
        files_per_ext = config.getint('CopyTool', 'FilesToCopyPerExtension', fallback=10)
        file_patterns_str = config.get('CopyTool', 'FilePatterns', fallback='*.wav')
        file_patterns = [p.strip() for p in file_patterns_str.split(',') if p.strip()]
    except (configparser.Error, KeyError) as e:
        print(f"FATAL: Missing required setting in config.ini: {e}", file=sys.stderr); sys.exit(1)

    setup_logger(dest_dir, is_debug_mode)
    logger.info(f"Configuration loaded. Source: '{source_dir}', Destination: '{dest_dir}'.")
    
    if not verify_extlist_integrity(config):
        logger.critical("Exiting due to failed integrity check."); sys.exit(1)

    if not os.path.isdir(source_dir):
        logger.critical(f"Source Root '{source_dir}' from config.ini is not a valid directory."); sys.exit(1)

    try:
        main(
            source_root=source_dir,
            base_dest_path=dest_dir,
            files_to_copy_count=files_per_ext,
            file_patterns=file_patterns,
            cli_date_override=cli_date
        )
    except Exception as e:
        logger.critical("An unhandled exception occurred during main execution.", exc_info=True); sys.exit(1)

    logger.info("Script execution process completed.")