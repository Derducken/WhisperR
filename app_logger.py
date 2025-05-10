import datetime
import os
import sys
import traceback
from pathlib import Path
from constants import LOG_LEVELS, DEFAULT_LOGGING_LEVEL, LOG_FILE_PREFIX

LOG_LEVEL_ORDER = {level.lower(): i for i, level in enumerate(LOG_LEVELS)}

class AppLogger:
    def __init__(self, base_path: Path):
        self.log_level_str: str = DEFAULT_LOGGING_LEVEL
        self.log_to_file_enabled: bool = False
        self.log_file_path: Path | None = None
        self.log_file_handle = None
        self.base_path = base_path

    def configure(self, level: str, log_to_file: bool):
        self.log_level_str = level
        self.log_to_file_enabled = log_to_file

        if not self.log_to_file_enabled and self.log_file_handle:
            try:
                self.log_file_handle.close()
            except Exception:
                pass
            self.log_file_handle = None
            self.log_file_path = None
        elif self.log_to_file_enabled and not self.log_file_handle:
            self._open_log_file()

    def _open_log_file(self):
        if self.log_file_handle:
            return
        try:
            log_dir = self.base_path
            log_dir.mkdir(parents=True, exist_ok=True)
            timestamp_file = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            log_filename = f"{LOG_FILE_PREFIX}{timestamp_file}.txt"
            self.log_file_path = log_dir / log_filename
            self.log_file_handle = open(self.log_file_path, 'a', encoding='utf-8')
            self.log_message_internal("INFO", f"Logging to file: {self.log_file_path}", force_print=True)
        except Exception as e:
            self.log_message_internal("ERROR", f"Error opening log file: {e}", force_print=True)
            self.log_to_file_enabled = False
            self.log_file_handle = None

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

        print(log_line)

        if self.log_to_file_enabled and self.log_file_handle:
            try:
                self.log_file_handle.write(log_line + "\n")
                self.log_file_handle.flush()
            except Exception as e:
                print(f"[{timestamp}] [WhisperR] [LOGGER_ERROR] Error writing to log file: {e}", file=sys.stderr)
                try:
                    self.log_file_handle.close()
                except: pass
                self.log_file_handle = None

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
        temp_logger = AppLogger(Path("."))
        temp_logger.log_level_str = "DEBUG" 
        temp_logger.log_to_file_enabled = False
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