import enum

# --- File Names ---
CONFIG_FILE_NAME = "config.json"
PROMPT_FILE_NAME = "prompt.json"
COMMANDS_FILE_NAME = "commands.json"
APP_ICON_NAME = "WhisperR_icon.png" # Changed to actual icon file name
LOG_FILE_PREFIX = "whisperr_log_"

# --- Default Hotkeys ---
DEFAULT_HOTKEY_TOGGLE = "ctrl+alt+space"
DEFAULT_HOTKEY_SHOW = "ctrl+alt+shift+space"
DEFAULT_HOTKEY_PUSH_TO_TALK = "" # Default to empty, user must set

# --- Status Bar ---
DEFAULT_STATUS_BAR_POSITION = "Top"
DEFAULT_STATUS_BAR_SIZE = 5 # pixels
STATUS_BAR_POSITIONS = ["Top", "Bottom", "Left", "Right"]

# --- Alternative Status Indicator ---
ALT_INDICATOR_POSITIONS = ["Top-Left", "Top-Right", "Bottom-Left", "Bottom-Right"]
DEFAULT_ALT_INDICATOR_POSITION = "Bottom-Right"
DEFAULT_ALT_INDICATOR_SIZE = 64 # px
DEFAULT_ALT_INDICATOR_OFFSET = 10 # px
MIN_ALT_INDICATOR_SIZE = 32
MAX_ALT_INDICATOR_SIZE = 128
MIN_ALT_INDICATOR_OFFSET = 0
MAX_ALT_INDICATOR_OFFSET = 100

# --- Logging ---
LOG_LEVELS = ["None", "ERROR", "WARNING", "Essential", "Extended", "Everything", "Debug"] # Ensure WARNING is here
DEFAULT_LOGGING_LEVEL = "Essential"
DEFAULT_MAX_LOG_FILES = 10 # Max number of log files to keep
# LOG_LEVEL_ORDER is generated in app_logger.py based on this list

# --- Transcription Behavior ---
DEFAULT_AUTO_ADD_SPACE = True

# --- Audio ---
AUDIO_QUEUE_SENTINEL = None
DEFAULT_SILENCE_THRESHOLD_SECONDS = 3.0
DEFAULT_VAD_ENERGY_THRESHOLD = 300
DEFAULT_MAX_MEMORY_SEGMENT_DURATION_SECONDS = 60
MIN_MAX_MEMORY_SEGMENT_DURATION = 10
MAX_MAX_MEMORY_SEGMENT_DURATION = 5400
AUDIO_FORMATS = ["WAV", "MP3", "AAC"]
DEFAULT_AUDIO_FORMAT = "WAV"
AUDIO_FORMAT_TOOLTIPS = {
    "WAV": "Lossless, best quality, largest file size. Universally compatible.",
    "MP3": "Lossy compression, good balance of quality and size. Widely compatible.",
    "AAC": "Lossy compression, often better quality than MP3 at similar bitrates, smaller files. Good compatibility.",
}
AUDIO_SAMPLE_RATE = 44100
AUDIO_CHANNELS = 1
AUDIO_BLOCKSIZE = 1024
AUDIO_DTYPE = 'int16'

# --- Whisper Engine ---
WHISPER_ENGINES = ["Executable"] # Now only one option
DEFAULT_WHISPER_ENGINE = "Executable" # This setting might become redundant
DEFAULT_WHISPER_EXECUTABLE = "whisper"

# --- Models ---
# FASTER_WHISPER_MODELS = [ # REMOVE
#     "tiny", "tiny.en", "base", "base.en", "small", "small.en", # REMOVE
#     "medium", "medium.en", "large-v1", "large-v2", "large-v3", # REMOVE
#     "distil-large-v2", "distil-medium.en", "distil-small.en" # REMOVE
# ] # REMOVE
CLI_MODEL_OPTIONS = ["tiny", "base", "small", "medium", "large",
                     "tiny.en", "base.en", "small.en", "medium.en"]
# EXTENDED_MODEL_OPTIONS are effectively just CLI_MODEL_OPTIONS now for model dropdowns if they use it.
# Or, the model dropdown in UI should only show CLI_MODEL_OPTIONS.
EXTENDED_MODEL_OPTIONS = sorted(list(set(CLI_MODEL_OPTIONS)))


DEFAULT_LANGUAGE = "en"
DEFAULT_MODEL = "large" # For CLI
# DEFAULT_FW_MODEL = "base" # REMOVE - For faster-whisper lib

# --- UI Themes ---
class Theme(enum.Enum):
    LIGHT = "Light"
    DARK = "Dark"
    SYSTEM = "System"

DEFAULT_THEME = Theme.LIGHT.value
UI_THEMES = [t.value for t in Theme]

# --- App Behavior ---
class CloseBehavior(enum.Enum):
    TRAY = "Minimize to tray"
    EXIT = "Exit app"

DEFAULT_CLOSE_BEHAVIOR = CloseBehavior.TRAY.value
CLOSE_BEHAVIORS = [cb.value for cb in CloseBehavior]

DEFAULT_BACKUP_FOLDER = "OldVersions"
DEFAULT_MAX_BACKUPS = 10
DEFAULT_EXPORT_FOLDER = "."

COLOR_STATUS_RECORDING_VAD_ACTIVE = "#FF0000"
COLOR_STATUS_RECORDING_VAD_WAITING = "#A9A9A9"
COLOR_STATUS_RECORDING_CONTINUOUS = "#FF0000"
COLOR_STATUS_TRANSCRIBING = "#0000FF"
COLOR_STATUS_RECORDING_AND_TRANSCRIBING = "#800080"
COLOR_STATUS_IDLE_NOT_RECORDING = "#808080"
