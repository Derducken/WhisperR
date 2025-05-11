import json
import os
import shutil
import datetime
import sys
import platform # For OS-specific paths
from pathlib import Path
from dataclasses import dataclass, field, asdict, fields, is_dataclass
from typing import List, Dict, Any, TypeVar, Type

# Import your helper functions directly
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 

from constants import (
    CONFIG_FILE_NAME, PROMPT_FILE_NAME, COMMANDS_FILE_NAME,
    DEFAULT_HOTKEY_TOGGLE, DEFAULT_HOTKEY_SHOW, DEFAULT_STATUS_BAR_POSITION,
    DEFAULT_STATUS_BAR_SIZE, DEFAULT_LANGUAGE, DEFAULT_MODEL, DEFAULT_WHISPER_EXECUTABLE,
    DEFAULT_SILENCE_THRESHOLD_SECONDS, DEFAULT_VAD_ENERGY_THRESHOLD, DEFAULT_EXPORT_FOLDER,
    DEFAULT_BACKUP_FOLDER, DEFAULT_MAX_BACKUPS, CloseBehavior, DEFAULT_CLOSE_BEHAVIOR,
    LOG_LEVELS, DEFAULT_LOGGING_LEVEL, DEFAULT_WHISPER_ENGINE, WHISPER_ENGINES,
    DEFAULT_AUDIO_FORMAT, AUDIO_FORMATS,
    DEFAULT_MAX_MEMORY_SEGMENT_DURATION_SECONDS,
    DEFAULT_THEME, UI_THEMES,
    ALT_INDICATOR_POSITIONS, DEFAULT_ALT_INDICATOR_POSITION,
    DEFAULT_ALT_INDICATOR_SIZE, DEFAULT_ALT_INDICATOR_OFFSET,
    DEFAULT_MAX_LOG_FILES, DEFAULT_AUTO_ADD_SPACE, DEFAULT_HOTKEY_PUSH_TO_TALK
)


T = TypeVar('T')

def get_user_config_dir(app_name: str = "WhisperR") -> Path:
    """Returns a user-specific directory for application configuration files."""
    if platform.system() == "Windows":
        # APPDATA is typically C:\Users\<username>\AppData\Roaming
        path = Path(os.getenv('APPDATA', Path.home() / "AppData" / "Roaming")) / app_name
    elif platform.system() == "Darwin": # macOS
        path = Path.home() / "Library" / "Application Support" / app_name
    else: # Linux and other XDG-based systems
        path = Path(os.getenv('XDG_CONFIG_HOME', Path.home() / ".config")) / app_name
    
    try:
        path.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        log_error(f"Could not create user config directory {path}: {e}. Falling back to local directory.")
        # Fallback to a directory in the same location as the executable or script (less ideal for persistence)
        if hasattr(sys, '_MEIPASS'):
            path = Path(sys._MEIPASS) / app_name # For PyInstaller bundle
            path.mkdir(parents=True, exist_ok=True) # Try again in MEIPASS
        else:
            path = Path(".") / app_name # For development, relative to script
            path.mkdir(parents=True, exist_ok=True) # Try again locally
    return path

def _ensure_type(value: Any, target_type: Type[T], default_value: T) -> T:
    """Helper to ensure value is of target_type, with robust conversion for common cases."""
    if isinstance(value, target_type):
        if target_type == bool and not isinstance(value, bool):
             pass
        else:
            return value
    try:
        if target_type == bool:
            if isinstance(value, str):
                return value.lower() in ['true', '1', 'yes', 'on']
            return bool(value)
        elif target_type == int:
            return int(float(value))
        elif target_type == float:
            return float(value)
        elif target_type == str:
            return str(value)
        elif target_type == list and isinstance(default_value, list):
            return default_value if not isinstance(value, list) else value
        elif target_type == dict and isinstance(default_value, dict):
             return default_value if not isinstance(value, dict) else value
        else:
            return target_type(value)
    except (ValueError, TypeError):
        # Use the imported helper function correctly
        log_extended(f"Type conversion failed for value '{value}' to {target_type}. Using default '{default_value}'.")
        return default_value


@dataclass
class CommandEntry:
    voice: str = ""
    action: str = ""

@dataclass
class AppSettings:
    # Main App
    versioning_enabled: bool = True
    whisper_executable: str = DEFAULT_WHISPER_EXECUTABLE
    language: str = DEFAULT_LANGUAGE
    model: str = DEFAULT_MODEL
    translation_enabled: bool = False
    command_mode: bool = False
    timestamps_disabled: bool = False
    clear_text_output: bool = False
    export_folder: str = DEFAULT_EXPORT_FOLDER
    clear_audio_on_exit: bool = False
    clear_text_on_exit: bool = False
    logging_level: str = DEFAULT_LOGGING_LEVEL
    log_to_file: bool = False
    backup_folder: str = DEFAULT_BACKUP_FOLDER
    max_backups: int = DEFAULT_MAX_BACKUPS
    close_behavior: str = DEFAULT_CLOSE_BEHAVIOR
    ui_theme: str = DEFAULT_THEME

    # Hotkeys
    hotkey_toggle_record: str = DEFAULT_HOTKEY_TOGGLE
    hotkey_show_window: str = DEFAULT_HOTKEY_SHOW
    hotkey_push_to_talk: str = DEFAULT_HOTKEY_PUSH_TO_TALK

    # Audio & VAD
    selected_audio_device_index: int | None = None
    silence_threshold_seconds: float = DEFAULT_SILENCE_THRESHOLD_SECONDS
    vad_energy_threshold: int = DEFAULT_VAD_ENERGY_THRESHOLD
    max_memory_segment_duration_seconds: int = DEFAULT_MAX_MEMORY_SEGMENT_DURATION_SECONDS
    audio_segment_format: str = DEFAULT_AUDIO_FORMAT

    # Notifications & Output Actions
    beep_on_transcription: bool = False
    beep_on_save_audio_segment: bool = False
    whisper_cli_beeps_enabled: bool = False
    auto_paste: bool = False
    auto_paste_delay: float = 1.0

    # Status Bar (Windows Edge)
    status_bar_enabled: bool = True
    status_bar_position: str = DEFAULT_STATUS_BAR_POSITION
    status_bar_size: int = DEFAULT_STATUS_BAR_SIZE

    # Alternative Status Indicator (Cross-platform corner icon)
    alt_status_indicator_enabled: bool = False
    alt_status_indicator_position: str = DEFAULT_ALT_INDICATOR_POSITION
    alt_status_indicator_size: int = DEFAULT_ALT_INDICATOR_SIZE
    alt_status_indicator_offset: int = DEFAULT_ALT_INDICATOR_OFFSET

    # Whisper Engine Choice
    whisper_engine_type: str = DEFAULT_WHISPER_ENGINE # Will always be "Executable"

    # Scratchpad
    scratchpad_append_mode: bool = False

    # Log Management
    max_log_files: int = DEFAULT_MAX_LOG_FILES

    # Transcription Behavior
    auto_add_space: bool = DEFAULT_AUTO_ADD_SPACE


    @classmethod
    def from_dict(cls: Type[T], data: Dict[str, Any]) -> T:
        default_instance = cls()
        default_values = asdict(default_instance)
        init_data = {}
        for f in fields(cls):
            if f.name in data:
                init_data[f.name] = _ensure_type(data[f.name], f.type, default_values[f.name])
            else:
                init_data[f.name] = default_values[f.name]
        
        if "disable_whisper_native_beep" in data and "whisper_cli_beeps_enabled" not in data:
            init_data["whisper_cli_beeps_enabled"] = not _ensure_type(
                data["disable_whisper_native_beep"], bool, True
            )
            # Use the imported helper function correctly
            log_extended("Migrated 'disable_whisper_native_beep' to 'whisper_cli_beeps_enabled'.")
        return cls(**init_data)


class SettingsManager:
    def __init__(self, config_base_path: Path): # Changed base_path to config_base_path
        self.base_path = config_base_path # This is now the user-specific config dir
        self.base_path.mkdir(parents=True, exist_ok=True)

        self.config_file = self.base_path / CONFIG_FILE_NAME
        self.prompt_file = self.base_path / PROMPT_FILE_NAME
        self.commands_file = self.base_path / COMMANDS_FILE_NAME

        self.settings = AppSettings()
        self.prompt: str = ""
        self.commands: List[CommandEntry] = []

        self.load_all()

    def _load_json_file(self, file_path: Path, default_content: Any = None) -> Any:
        if default_content is None:
            default_content = {}
        if file_path.exists():
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                # Use the imported helper function correctly
                log_error(f"Error decoding JSON from {file_path}: {e}. Using defaults/empty.")
                self._backup_corrupted_file(file_path, "decode_error")
            except Exception as e:
                # Use the imported helper function correctly
                log_error(f"Error loading {file_path}: {e}. Using defaults/empty.")
        else:
            # Use the imported helper function correctly
            log_essential(f"{file_path.name} not found at {file_path}, using defaults/empty.") # Log full path
        return default_content

    def _save_json_file(self, data: Any, file_path: Path, perform_backup: bool = True):
        if perform_backup and self.settings.versioning_enabled:
            self._create_backup(file_path)
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            # Use the imported helper function correctly
            log_essential(f"Saved data to {file_path}") # Log full path
        except Exception as e:
            # Use the imported helper function correctly
            log_error(f"Error saving {file_path.name}: {e}")

    def load_settings(self):
        settings_data = self._load_json_file(self.config_file)
        self.settings = AppSettings.from_dict(settings_data)
        # Use the imported helper function correctly
        log_essential(f"Core settings loaded. UI Theme: {self.settings.ui_theme}")


    def save_settings(self):
        # Resolve paths relative to a user-writable directory if they are not absolute
        # For export_folder and backup_folder, if they are relative, they should be relative to user documents or similar
        # For whisper_executable, it's usually a system path or a path provided by user.
        
        # Ensure whisper_executable is an absolute path or a command findable in PATH
        if self.settings.whisper_executable and not Path(self.settings.whisper_executable).is_absolute():
            # If it's just a name like "whisper", assume it's in PATH. Otherwise, it might be problematic.
            # No change needed here if it's just a name. Path.resolve() might fail if not in CWD.
            pass
        elif self.settings.whisper_executable:
             self.settings.whisper_executable = str(Path(self.settings.whisper_executable).resolve())
        else:
            self.settings.whisper_executable = DEFAULT_WHISPER_EXECUTABLE

        # For export and backup folders, ensure they are absolute or make them relative to user config dir
        # This behavior might need refinement based on desired UX (e.g. always user documents for export)
        for folder_attr in ['export_folder', 'backup_folder']:
            folder_val = getattr(self.settings, folder_attr)
            default_folder_val = DEFAULT_EXPORT_FOLDER if folder_attr == 'export_folder' else DEFAULT_BACKUP_FOLDER
            
            if folder_val:
                path_obj = Path(folder_val)
                if not path_obj.is_absolute():
                    # If relative, make it relative to the config base path (user data dir)
                    # This ensures backups and default exports go to a known user location
                    setattr(self.settings, folder_attr, str(self.base_path / path_obj)) 
                else:
                    setattr(self.settings, folder_attr, str(path_obj.resolve()))
            else: # Is empty or None
                # Set to default, relative to config base path
                setattr(self.settings, folder_attr, str(self.base_path / default_folder_val))
            
            # Ensure the directory exists after resolving
            try:
                Path(getattr(self.settings, folder_attr)).mkdir(parents=True, exist_ok=True)
            except Exception as e:
                log_error(f"Could not create directory for {folder_attr} at {getattr(self.settings, folder_attr)}: {e}")


        self._save_json_file(asdict(self.settings), self.config_file)

    def load_prompt(self):
        prompt_data = self._load_json_file(self.prompt_file, default_content={'prompt': ''})
        self.prompt = prompt_data.get('prompt', '')
        # Use the imported helper function correctly
        log_essential("Prompt loaded.")

    def save_prompt(self):
        self._save_json_file({'prompt': self.prompt}, self.prompt_file)

    def load_commands(self):
        commands_data = self._load_json_file(self.commands_file, default_content={'commands': []})
        loaded_raw_commands = commands_data.get('commands', [])
        self.commands = []
        if isinstance(loaded_raw_commands, list):
            for cmd_dict in loaded_raw_commands:
                if isinstance(cmd_dict, dict) and "voice" in cmd_dict and "action" in cmd_dict:
                    self.commands.append(CommandEntry(voice=str(cmd_dict["voice"]), action=str(cmd_dict["action"])))
                else:
                    # Use the imported helper function correctly
                    log_extended(f"Skipping invalid command entry: {cmd_dict}")
        # Use the imported helper function correctly
        log_essential(f"Commands loaded: {len(self.commands)} entries.")


    def save_commands(self):
        commands_to_save = [asdict(cmd) for cmd in self.commands]
        self._save_json_file({'commands': commands_to_save}, self.commands_file)

    def load_all(self):
        self.load_settings()
        self.load_prompt()
        self.load_commands()

    def save_all(self):
        self.save_settings()
        self.save_prompt()
        self.save_commands()

    def _create_backup(self, file_path: Path):
        # Ensure backup_folder path is resolved correctly before use
        current_backup_folder_str = self.settings.backup_folder
        if not current_backup_folder_str: # Should have been defaulted in save_settings
            current_backup_folder_str = str(self.base_path / DEFAULT_BACKUP_FOLDER)
        
        backup_dir = Path(current_backup_folder_str)
        if not backup_dir.is_absolute(): # Should have been made absolute in save_settings
            backup_dir = self.base_path / backup_dir
            
        if not self.settings.versioning_enabled or not file_path.exists() or self.settings.max_backups <= 0:
            return

        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            # Use the imported helper function correctly
            log_error(f"Error creating backup folder '{backup_dir}': {e}")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_filepath = backup_dir / backup_filename

        try:
            shutil.copy2(file_path, backup_filepath)
            # Use the imported helper function correctly
            log_extended(f"Created backup: {backup_filepath}")
            self._manage_backups(backup_dir, file_path.stem, file_path.suffix)
        except Exception as e:
            # Use the imported helper function correctly
            log_error(f"Error creating backup for '{file_path.name}': {e}")

    def _backup_corrupted_file(self, file_path: Path, reason: str):
        # Corrupted backups go into a subfolder of the main config dir
        backup_dir = self.base_path / "corrupted_backups"
        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{file_path.stem}_{reason}_{timestamp}{file_path.suffix}"
            backup_filepath = backup_dir / backup_filename
            if file_path.exists():
                shutil.move(str(file_path), str(backup_filepath))
                # Use the imported helper function correctly
                log_error(f"Corrupted file {file_path.name} backed up to {backup_filepath}")
        except Exception as e:
            # Use the imported helper function correctly
            log_error(f"Could not back up corrupted file {file_path.name}: {e}")


    def _manage_backups(self, backup_dir: Path, base_filename_stem: str, file_extension: str):
        if self.settings.max_backups <= 0:
            return
        try:
            import re
            pattern_str = rf"^{re.escape(base_filename_stem)}_\d{{8}}_\d{{6}}{re.escape(file_extension)}$"
            pattern = re.compile(pattern_str)
            backups = []
            for f in backup_dir.iterdir():
                if f.is_file() and pattern.match(f.name):
                    try:
                        backups.append((f, f.stat().st_mtime))
                    except OSError:
                        continue
            
            backups.sort(key=lambda x: x[1])

            num_to_delete = len(backups) - self.settings.max_backups
            if num_to_delete > 0:
                for f_path, _ in backups[:num_to_delete]:
                    try:
                        f_path.unlink()
                        # Use the imported helper function correctly
                        log_extended(f"Deleted old backup: {f_path.name}")
                    except Exception as e:
                        # Use the imported helper function correctly
                        log_error(f"Error deleting old backup {f_path.name}: {e}")
        except Exception as e:
            # Use the imported helper function correctly
            log_error(f"Error managing backups in '{backup_dir}': {e}")

# This function is for asset loading (like icons) from bundle or dev location
def get_app_asset_path() -> Path:
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent
