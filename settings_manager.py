import json
import os
import shutil
import datetime
import sys
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
    DEFAULT_FW_MODEL, DEFAULT_AUDIO_FORMAT, AUDIO_FORMATS,
    DEFAULT_MAX_MEMORY_SEGMENT_DURATION_SECONDS,
    DEFAULT_THEME, UI_THEMES,
    ALT_INDICATOR_POSITIONS, DEFAULT_ALT_INDICATOR_POSITION,
    DEFAULT_ALT_INDICATOR_SIZE, DEFAULT_ALT_INDICATOR_OFFSET
)


T = TypeVar('T')

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
    whisper_engine_type: str = DEFAULT_WHISPER_ENGINE
    faster_whisper_model_name: str = DEFAULT_FW_MODEL

    # Scratchpad
    scratchpad_append_mode: bool = False


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
    def __init__(self, base_path: Path):
        self.base_path = base_path
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
            log_essential(f"{file_path.name} not found, using defaults/empty.")
        return default_content

    def _save_json_file(self, data: Any, file_path: Path, perform_backup: bool = True):
        if perform_backup and self.settings.versioning_enabled:
            self._create_backup(file_path)
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            # Use the imported helper function correctly
            log_essential(f"Saved data to {file_path.name}")
        except Exception as e:
            # Use the imported helper function correctly
            log_error(f"Error saving {file_path.name}: {e}")

    def load_settings(self):
        settings_data = self._load_json_file(self.config_file)
        self.settings = AppSettings.from_dict(settings_data)
        # Use the imported helper function correctly
        log_essential(f"Core settings loaded. UI Theme: {self.settings.ui_theme}")


    def save_settings(self):
        self.settings.whisper_executable = str(Path(self.settings.whisper_executable).resolve()) if self.settings.whisper_executable else DEFAULT_WHISPER_EXECUTABLE
        self.settings.export_folder = str(Path(self.settings.export_folder).resolve()) if self.settings.export_folder else DEFAULT_EXPORT_FOLDER
        self.settings.backup_folder = str(Path(self.settings.backup_folder).resolve()) if self.settings.backup_folder else DEFAULT_BACKUP_FOLDER
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
        if not self.settings.versioning_enabled or not file_path.exists() or self.settings.max_backups <= 0:
            return

        backup_dir = Path(self.settings.backup_folder)
        try:
            if not backup_dir.is_absolute():
                backup_dir = self.base_path / backup_dir
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

    def get_app_data_folder_path(self) -> Path:
        """Returns the base path where config files are stored."""
        # This function is not typically part of SettingsManager itself,
        # but rather a utility function. Keeping it here as per original structure for now.
        # Consider moving to a general utils.py or main_app.py if it makes more sense.
        if hasattr(sys, '_MEIPASS'): 
            return Path(sys._MEIPASS)
        return Path(os.path.dirname(os.path.abspath(__file__)))


def get_app_base_path() -> Path:
    """Determines the application's base directory for data/configs."""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent