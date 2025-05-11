import datetime
import os
import sys
import traceback
from pathlib import Path
import re # For log file pattern matching
from typing import Optional # Import Optional
from constants import LOG_LEVELS, DEFAULT_LOGGING_LEVEL, LOG_FILE_PREFIX, DEFAULT_MAX_LOG_FILES

LOG_LEVEL_ORDER = {level.lower(): i for i, level in enumerate(LOG_LEVELS)}

class AppLogger:
    def __init__(self, base_path: Path):
        self.log_level_str: str = DEFAULT_LOGGING_LEVEL
        self.log_to_file_enabled: bool = False
        self.log_file_path: Path | None = None
        self.log_file_handle = None
        self.base_path = base_path # This is user_config_path / "logs" effectively
        self.log_dir = self.base_path / "logs" # Specific subdirectory for logs
        self.max_log_files: int = DEFAULT_MAX_LOG_FILES # Initialize with default

    def configure(self, level: str, log_to_file: bool, max_log_files: Optional[int] = None):
        self.log_level_str = level
        self.log_to_file_enabled = log_to_file
        if max_log_files is not None and max_log_files > 0:
            self.max_log_files = max_log_files
        else: # Use default if None or invalid
            self.max_log_files = DEFAULT_MAX_LOG_FILES


        if not self.log_to_file_enabled and self.log_file_handle:
            try:
                self.log_file_handle.close()
            except Exception:
                pass
            self.log_file_handle = None
            self.log_file_path = None
        elif self.log_to_file_enabled: # If enabling or already enabled
            if not self.log_file_handle: # Open if not already open
                self._open_log_file()
            else: # If already open, still run management in case max_log_files changed
                self._manage_log_files()


    def _open_log_file(self):
        if self.log_file_handle: # Already open
            return
        try:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            timestamp_file = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            log_filename = f"{LOG_FILE_PREFIX}{timestamp_file}.txt"
            self.log_file_path = self.log_dir / log_filename
            self.log_file_handle = open(self.log_file_path, 'a', encoding='utf-8')
            self.log_message_internal("INFO", f"Logging to file: {self.log_file_path}", force_print=True)
            self._manage_log_files() # Manage logs after successfully opening a new one
        except Exception as e:
            self.log_message_internal("ERROR", f"Error opening log file: {e}", force_print=True)
            self.log_to_file_enabled = False # Disable if opening fails
            self.log_file_handle = None

    def _manage_log_files(self):
        if not self.log_to_file_enabled or self.max_log_files <= 0:
            return

        try:
            log_files = []
            # Regex to match log files: whisperr_log_YYYYMMDD_HHMMSS.txt
            log_pattern = re.compile(rf"^{re.escape(LOG_FILE_PREFIX)}\d{{8}}_\d{{6}}\.txt$")
            
            for item in self.log_dir.iterdir():
                if item.is_file() and log_pattern.match(item.name):
                    try:
                        log_files.append((item, item.stat().st_mtime))
                    except OSError: # File might be deleted between iterdir and stat
                        continue
            
            log_files.sort(key=lambda x: x[1]) # Sort by modification time (oldest first)

            num_to_delete = len(log_files) - self.max_log_files
            if num_to_delete > 0:
                for f_path, _ in log_files[:num_to_delete]:
                    try:
                        f_path.unlink()
                        self.log_message_internal("INFO", f"Deleted old log file: {f_path.name}", force_print=True)
                    except Exception as e_del:
                        self.log_message_internal("ERROR", f"Error deleting old log file {f_path.name}: {e_del}", force_print=True)
        except Exception as e_manage:
            self.log_message_internal("ERROR", f"Error managing log files: {e_manage}", force_print=True)


    def log_message(self, level: str, message: str, exc_info=False):
        self.log_message_internal(level, message, exc_info)

    def log_message_internal(self, level: str, message: str, exc_info=False, force_print=False):
        try:
            msg_level_val = LOG_LEVEL_ORDER.get(level.lower())
            current_level_val = LOG_LEVEL_ORDER.get(self.log_level_str.lower())

            if msg_level_val is None or current_level_val is None:
                print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [WhisperR] [LOGGER_ERROR] Invalid log level: msg='{level}', current='{self.log_level_str}' for message: {message[:100]}", file=sys.stderr)
                return

            if current_level_val == LOG_LEVEL_ORDER["none"] and not force_print:
                return
            if msg_level_val > current_level_val and not force_print:
                return
        except Exception as e:
            print(f"Logger configuration/level check error: {e}", file=sys.stderr)
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        log_line = f"[{timestamp}] [WhisperR] [{level.upper()}] {message}"
        if exc_info:
            log_line += f"\n{traceback.format_exc()}"

        print(log_line) # Always print to console based on level check

        if self.log_to_file_enabled and self.log_file_handle:
            try:
                self.log_file_handle.write(log_line + "\n")
                self.log_file_handle.flush()
            except Exception as e:
                # Avoid recursive logging if log_message_internal is called from here
                print(f"[{timestamp}] [WhisperR] [LOGGER_ERROR] Error writing to log file: {e}", file=sys.stderr)
                try:
                    self.log_file_handle.close() # Attempt to close
                except: pass
                self.log_file_handle = None # Mark as closed to prevent further write attempts this session
                self.log_to_file_enabled = False # Disable file logging for this session

    def get_log_file_path(self) -> Path | None:
        return self.log_file_path

    def close(self):
        if self.log_file_handle:
            try:
                self.log_message_internal("INFO", "Closing log file.", force_print=True)
                self.log_file_handle.close()
            except Exception as e:
                print(f"Error closing log file: {e}", file=sys.stderr)
            finally:
                self.log_file_handle = None

LOGGER: AppLogger | None = None

def get_logger() -> AppLogger:
    global LOGGER
    if LOGGER is None:
        print("WARNING: Logger accessed before full initialization. Creating temporary console logger.", file=sys.stderr)
        # Create a temporary base_path for the fallback logger
        temp_base_path = Path(".") 
        try:
            # Attempt to use a user-specific temp location if possible, even for fallback
            temp_base_path = Path(os.getenv('TEMP', Path.home())) / "WhisperR_fallback_logs"
            temp_base_path.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass # Stick to Path(".") if user temp fails

        temp_logger = AppLogger(temp_base_path)
        temp_logger.configure(level="DEBUG", log_to_file=False) # Default to console for fallback
        return temp_logger
    return LOGGER

def log_error(message: str, exc_info=True):
    get_logger().log_message("ERROR", message, exc_info=exc_info)

def log_warning(message: str):
    get_logger().log_message("WARNING", message)

def log_essential(message: str):
    get_logger().log_message("ESSENTIAL", message)

def log_extended(message: str):
    get_logger().log_message("EXTENDED", message)

def log_debug(message: str):
    get_logger().log_message("DEBUG", message)
