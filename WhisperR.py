import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import Menu
import sounddevice as sd
import keyboard
import pystray
from PIL import Image
import win32gui
import win32con
import win32api
import wave
import os
import datetime
import subprocess
import shutil
import threading
import time
import numpy as np
import re
import json
from collections import deque # deque is no longer used for audio_buffer
import queue
import winsound

CONFIG_FILE = "config.json"
PROMPT_FILE = "prompt.json"
COMMANDS_FILE = "commands.json"
AUDIO_QUEUE_SENTINEL = None

# --- Constants ---
DEFAULT_HOTKEY_TOGGLE = "ctrl+alt+space"
DEFAULT_HOTKEY_SHOW = "ctrl+alt+shift+space"


class WhisperRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhisperR")
        self.root.minsize(600, 600) # Slightly adjusted min height

        # --- Default Settings ---
        self.versioning_enabled = True
        self.whisper_executable = "whisper"
        self.language = "en"
        self.model = "large"
        self.translation_enabled = False
        self.command_mode = False
        self.timestamps_disabled = False
        self.clear_text_output = False
        self.prompt = ""
        self.selected_audio_device_index = None
        self.export_folder = "."
        self.silence_threshold_seconds = 5.0
        self.vad_energy_threshold = 300
        self.clear_audio_on_exit = False
        self.clear_text_on_exit = False
        self.loaded_commands = []
        self.scratchpad_append_mode = False
        self.logging_level = "Everything"  # None, Essential, Extended, Everything
        self.log_to_file = False
        self.beep_on_transcription = False
        self.beep_on_save_audio_segment = False # <<< NEW SETTING
        self.auto_paste = False
        self.auto_paste_delay = 1.0
        self.backup_folder = "OldVersions"
        self.max_backups = 10
        self.close_behavior = "Minimize to tray"
        self.hotkey_toggle_record = DEFAULT_HOTKEY_TOGGLE # New setting
        self.hotkey_show_window = DEFAULT_HOTKEY_SHOW   # New setting

        # --- Runtime State ---
        self.recording = False
        # self.audio_buffer = deque() # ***FIX: Removed unused audio_buffer***
        self.current_segment = []
        self.is_speaking = False
        self.silence_start_time = None
        self.green_line = None
        self.tray_icon = None
        self.commands = []
        self.vad_enabled = False
        self.config_window = None
        self.command_config_window = None
        self.scratchpad_window = None
        self.scratchpad_text_widget = None
        self.audio_stream = None
        self.recording_thread = None
        self.transcription_queue = queue.Queue()
        self.transcription_worker_thread = None
        self.log_file = None
        self.queue_processing_paused = False
        self.clear_queue_flag = False
        self.queue_indicator_var = tk.StringVar(value="Queue: 0")
        self._registered_hotkeys = {} # Store currently registered hotkeys {key_string: callback}
        self.shortcut_display_var = tk.StringVar() # For dynamic shortcut display

        # --- Style Configuration ---
        self.configure_styles()

        # --- Load, Create, Initialize ---
        self.load_settings()
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.handle_close_button)
        self.update_audio_device()
        self.update_hotkeys() # Register initial hotkeys
        self.start_transcription_worker()
        self.update_queue_indicator()

    def configure_styles(self):
        """Configures ttk styles."""
        self.style = ttk.Style(self.root)
        # Get default background for consistency
        try:
            self.default_bg = self.root.cget('bg')
            # Configure Label and Checkbutton background for main window consistency
            self.style.configure('TLabel', background=self.default_bg)
            self.style.configure('TCheckbutton', background=self.default_bg)
            # Configure Frame background (needed for frames holding labels/checkbuttons)
            self.style.configure('TFrame', background=self.default_bg)
        except tk.TclError:
            self.default_bg = 'SystemButtonFace' # Fallback color
            print("Warning: Could not get root background color. Using fallback.")

        # Custom style for LabelFrames in the configuration window
        self.style.configure(
            "Config.TLabelframe",
            padding=5, # Internal padding around content
            borderwidth=1,
            relief=tk.SOLID, # Use SOLID for a clear border
            # Add bordercolor if theme supports it (might need specific theme)
            # Example: bordercolor='gray70'
        )
        self.style.map( # Make border slightly darker on focus/hover if possible
             "Config.TLabelframe",
             bordercolor=[('active', 'gray50')]
        )

        self.style.configure(
            "Config.TLabelframe.Label",
            padding=(5, 2), # Padding for the label text within the frame border
            # background=... # Usually inherits LabelFrame background
            # font=... # Can customize font here too
        )
        # Note: Rounded corners are generally not supported by ttk LabelFrame borders.

    def log_message(self, level, message):
        level_map = {"None": 0, "Essential": 1, "Extended": 2, "Everything": 3}
        current_level = level_map.get(self.logging_level, 3)
        msg_level_map = {"essential": 1, "extended": 2, "everything": 3}
        msg_level = msg_level_map.get(level, 3)

        if current_level == 0 or msg_level > current_level:
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] [WhisperR] {message}"

        print(log_line)

        if self.log_to_file:
            # ... (rest of file logging logic remains the same)
            if not self.log_file:
                log_filename = f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                try:
                    log_dir = os.path.dirname(CONFIG_FILE) if os.path.dirname(CONFIG_FILE) else '.'
                    full_log_path = os.path.join(log_dir, log_filename)
                    self.log_file = open(full_log_path, 'a', encoding='utf-8')
                    print(f"[INFO] Logging to file: {full_log_path}") # Use print here as log_message might recurse
                except Exception as e:
                    print(f"[ERROR] Error opening log file: {e}")
                    self.log_to_file = False
            if self.log_file:
                try:
                    self.log_file.write(log_line + "\n")
                    self.log_file.flush()
                except Exception as e:
                    # Avoid logging error during logging to prevent loops
                    print(f"[ERROR] Error writing to log file: {e}")
                    try:
                         self.log_file.close()
                    except:
                         pass
                    self.log_file = None
                    self.log_to_file = False


    def load_settings(self):
        # Default values first
        self.hotkey_toggle_record = DEFAULT_HOTKEY_TOGGLE
        self.hotkey_show_window = DEFAULT_HOTKEY_SHOW

        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # ... (load other settings as before)
                    self.versioning_enabled = settings.get('versioning_enabled', True)
                    self.whisper_executable = settings.get('whisper_executable', 'whisper')
                    self.language = settings.get('language', 'en')
                    self.model = settings.get('model', 'large')
                    self.translation_enabled = settings.get('translation_enabled', False)
                    self.command_mode = settings.get('command_mode', False)
                    self.timestamps_disabled = settings.get('timestamps_disabled', False)
                    self.clear_text_output = settings.get('clear_text_output', False)
                    self.selected_audio_device_index = settings.get('selected_audio_device_index', None)
                    self.export_folder = settings.get('export_folder', '.')
                    self.silence_threshold_seconds = float(settings.get('silence_threshold_seconds', 5.0))
                    self.vad_energy_threshold = int(settings.get('vad_energy_threshold', 300))
                    self.clear_audio_on_exit = settings.get('clear_audio_on_exit', False)
                    self.clear_text_on_exit = settings.get('clear_text_on_exit', False)
                    self.logging_level = settings.get('logging_level', 'Everything')
                    self.log_to_file = settings.get('log_to_file', False)
                    self.beep_on_transcription = settings.get('beep_on_transcription', False)
                    self.beep_on_save_audio_segment = settings.get('beep_on_save_audio_segment', False) # <<< NEW
                    self.auto_paste = settings.get('auto_paste', False)
                    self.auto_paste_delay = float(settings.get('auto_paste_delay', 1.0))
                    self.backup_folder = settings.get('backup_folder', 'OldVersions')
                    self.max_backups = int(settings.get('max_backups', 10))
                    self.close_behavior = settings.get('close_behavior', 'Minimize to tray')
                    self.hotkey_toggle_record = settings.get('hotkey_toggle_record', DEFAULT_HOTKEY_TOGGLE) # Load hotkey
                    self.hotkey_show_window = settings.get('hotkey_show_window', DEFAULT_HOTKEY_SHOW)     # Load hotkey
                self.log_message("essential", f"Core settings loaded from {CONFIG_FILE}")
            else:
                self.log_message("essential", f"{CONFIG_FILE} not found, using core defaults.")
        except Exception as e:
            self.log_message("essential", f"Error loading {CONFIG_FILE}: {e}. Using core defaults.")
            # Ensure defaults are set even if file load fails partially
            self.hotkey_toggle_record = getattr(self, 'hotkey_toggle_record', DEFAULT_HOTKEY_TOGGLE)
            self.hotkey_show_window = getattr(self, 'hotkey_show_window', DEFAULT_HOTKEY_SHOW)

        # ... (load prompt and commands remain the same)
        try:
            if os.path.exists(PROMPT_FILE):
                with open(PROMPT_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.prompt = data.get('prompt', '')
                self.log_message("essential", f"Prompt loaded from {PROMPT_FILE}")
            else:
                self.prompt = ""
                self.log_message("essential", f"{PROMPT_FILE} not found, using empty prompt.")
        except Exception as e:
            self.prompt = ""
            self.log_message("essential", f"Error loading {PROMPT_FILE}: {e}. Using empty prompt.")

        try:
            if os.path.exists(COMMANDS_FILE):
                with open(COMMANDS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    loaded = data.get('commands', [])
                    if isinstance(loaded, list):
                        self.loaded_commands = [cmd for cmd in loaded if isinstance(cmd, dict) and "voice" in cmd and "action" in cmd]
                    else:
                        self.loaded_commands = []
                self.log_message("essential", f"Commands loaded from {COMMANDS_FILE}")
            else:
                self.loaded_commands = []
                self.log_message("essential", f"{COMMANDS_FILE} not found, using empty commands.")
        except Exception as e:
            self.loaded_commands = []
            self.log_message("essential", f"Error loading {COMMANDS_FILE}: {e}. Using empty commands.")

        self.commands = list(self.loaded_commands)
        self.vad_enabled = self.command_mode

    def save_settings(self):
        settings = {
            # ... (other settings)
            'versioning_enabled': getattr(self, 'versioning_enabled', True),
            'whisper_executable': getattr(self, 'whisper_executable', 'whisper'),
            'language': getattr(self, 'language', 'en'),
            'model': getattr(self, 'model', 'large'),
            'translation_enabled': getattr(self, 'translation_enabled', False),
            'command_mode': getattr(self, 'command_mode', False),
            'timestamps_disabled': getattr(self, 'timestamps_disabled', False),
            'clear_text_output': getattr(self, 'clear_text_output', False),
            'selected_audio_device_index': getattr(self, 'selected_audio_device_index', None),
            'export_folder': getattr(self, 'export_folder', '.'),
            'silence_threshold_seconds': getattr(self, 'silence_threshold_seconds', 5.0),
            'vad_energy_threshold': getattr(self, 'vad_energy_threshold', 300),
            'clear_audio_on_exit': getattr(self, 'clear_audio_on_exit', False),
            'clear_text_on_exit': getattr(self, 'clear_text_on_exit', False),
            'logging_level': getattr(self, 'logging_level', 'Everything'),
            'log_to_file': getattr(self, 'log_to_file', False),
            'beep_on_transcription': getattr(self, 'beep_on_transcription', False),
            'beep_on_save_audio_segment': getattr(self, 'beep_on_save_audio_segment', False), # <<< NEW
            'auto_paste': getattr(self, 'auto_paste', False),
            'auto_paste_delay': getattr(self, 'auto_paste_delay', 1.0),
            'backup_folder': getattr(self, 'backup_folder', 'OldVersions'),
            'max_backups': getattr(self, 'max_backups', 10),
            'close_behavior': getattr(self, 'close_behavior', 'Minimize to tray'),
            'hotkey_toggle_record': getattr(self, 'hotkey_toggle_record', DEFAULT_HOTKEY_TOGGLE), # Save hotkey
            'hotkey_show_window': getattr(self, 'hotkey_show_window', DEFAULT_HOTKEY_SHOW)     # Save hotkey
        }
        try:
            self.create_backup(CONFIG_FILE)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            self.log_message("essential", f"Core settings saved to {CONFIG_FILE}")
        except Exception as e:
            self.log_message("essential", f"Error saving {CONFIG_FILE}: {e}")

    # ... (save_prompt, save_commands remain the same) ...
    def save_prompt(self):
        prompt_data = {
            'prompt': getattr(self, 'prompt', '')
        }
        try:
            self.create_backup(PROMPT_FILE)
            with open(PROMPT_FILE, 'w', encoding='utf-8') as f:
                json.dump(prompt_data, f, indent=4)
            self.log_message("essential", f"Prompt saved to {PROMPT_FILE}")
        except Exception as e:
            self.log_message("essential", f"Error saving {PROMPT_FILE}: {e}")

    def save_commands(self):
        commands_data = {
            'commands': getattr(self, 'loaded_commands', [])
        }
        try:
            self.create_backup(COMMANDS_FILE)
            with open(COMMANDS_FILE, 'w', encoding='utf-8') as f:
                json.dump(commands_data, f, indent=4)
            self.log_message("essential", f"Commands saved to {COMMANDS_FILE}")
        except Exception as e:
            self.log_message("essential", f"Error saving {COMMANDS_FILE}: {e}")

    def create_menu(self):
        # ... (Menu creation remains the same) ...
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        # --- File Menu ---
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Prompt", command=self.import_prompt)
        file_menu.add_command(label="Export Prompt", command=self.export_prompt)
        file_menu.add_separator()
        file_menu.add_command(label="Open Scratchpad", command=self.open_scratchpad_window)
        file_menu.add_separator()
        file_menu.add_command(label="Quit WhisperR", command=self.quit_app_action)

        # --- Settings Menu ---
        settings_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Configuration...", command=self.open_configuration_window)
        settings_menu.add_command(label="Configure Commands...", command=self.open_command_configuration_window)

        # --- Queue Menu ---
        queue_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Queue", menu=queue_menu)
        self.pause_queue_menu_var = tk.BooleanVar(value=self.queue_processing_paused)
        queue_menu.add_checkbutton(label="Pause Queue Processing",
                                  command=self.toggle_queue_processing,
                                  variable=self.pause_queue_menu_var)
        queue_menu.add_command(label="Clear Queue", command=self.clear_transcription_queue)

    def create_widgets(self):
        self.create_menu()

        # --- Top Settings Frame ---
        # Use ttk.Frame for better style compatibility
        top_frame = ttk.Frame(self.root, style='TFrame') # Apply style with default bg
        top_frame.pack(pady=10, padx=10, fill=tk.X)

        lang_frame = ttk.Frame(top_frame, style='TFrame') # Use styled frame
        lang_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(lang_frame, text="Language:", style='TLabel').pack(side=tk.LEFT, padx=(0, 5)) # Apply style
        self.language_options = ["en", "es", "fr", "de", "it", "ja", "zh", "el"]
        self.language_combobox = ttk.Combobox(lang_frame, values=self.language_options, state="readonly", width=10)
        self.language_combobox.set(self.language)
        self.language_combobox.pack(side=tk.LEFT)
        self.language_combobox.bind("<<ComboboxSelected>>", lambda event: self.update_language())

        model_frame = ttk.Frame(top_frame, style='TFrame') # Use styled frame
        model_frame.pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Label(model_frame, text="Model:", style='TLabel').pack(side=tk.LEFT, padx=(0, 5)) # Apply style
        self.model_options = ["tiny", "base", "small", "medium", "large"]
        self.model_combobox = ttk.Combobox(model_frame, values=self.model_options, state="readonly", width=10)
        self.model_combobox.set(self.model)
        self.model_combobox.pack(side=tk.LEFT)
        self.model_combobox.bind("<<ComboboxSelected>>", lambda event: self.update_model())

        # --- Toggles Frame 1 ---
        toggle_frame1 = ttk.Frame(self.root, style='TFrame') # Use styled frame
        toggle_frame1.pack(pady=5, padx=10, fill=tk.X)

        trans_frame = ttk.Frame(toggle_frame1, style='TFrame') # Use styled frame
        trans_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.translation_var = tk.BooleanVar(value=self.translation_enabled)
        self.translation_var.trace_add("write", self.update_translation)
        ttk.Checkbutton(trans_frame, text="Enable Translation", variable=self.translation_var, style='TCheckbutton').pack(side=tk.LEFT) # Apply style

        cmd_frame = ttk.Frame(toggle_frame1, style='TFrame') # Use styled frame
        cmd_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.command_mode_var = tk.BooleanVar(value=self.command_mode)
        self.command_mode_var.trace_add("write", self.update_command_mode)
        ttk.Checkbutton(cmd_frame, text="Enable Auto-Pause / Commands", variable=self.command_mode_var, style='TCheckbutton').pack(side=tk.LEFT) # Apply style

        # --- Toggles Frame 2 ---
        toggle_frame2 = ttk.Frame(self.root, style='TFrame') # Use styled frame
        toggle_frame2.pack(pady=2, padx=10, fill=tk.X)

        ts_frame = ttk.Frame(toggle_frame2, style='TFrame') # Use styled frame
        ts_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.timestamps_disabled_var = tk.BooleanVar(value=self.timestamps_disabled)
        self.timestamps_disabled_var.trace_add("write", self.update_timestamps_disabled)
        ttk.Checkbutton(ts_frame, text="Disable Timestamps", variable=self.timestamps_disabled_var, style='TCheckbutton').pack(side=tk.LEFT) # Apply style

        clear_text_frame = ttk.Frame(toggle_frame2, style='TFrame') # Use styled frame
        clear_text_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.clear_text_output_var = tk.BooleanVar(value=self.clear_text_output)
        self.clear_text_output_var.trace_add("write", self.update_clear_text_output)
        ttk.Checkbutton(clear_text_frame, text="Clear Text Output", variable=self.clear_text_output_var, style='TCheckbutton').pack(side=tk.LEFT) # Apply style

        # --- Prompt Section ---
        prompt_label_frame = ttk.Frame(self.root, style='TFrame') # Frame for label
        prompt_label_frame.pack(pady=(10, 0), padx=10, fill=tk.X)
        ttk.Label(prompt_label_frame, text="Prompt:", style='TLabel').pack(anchor=tk.W) # Apply style
        # Text widget doesn't use ttk styles directly, bg/fg can be set manually if needed
        self.prompt_text = tk.Text(self.root, height=8, wrap=tk.WORD, undo=True)
        self.prompt_text.insert("1.0", self.prompt)
        self.prompt_text.pack(pady=5, padx=10, fill=tk.X)
        self.prompt_text.bind("<KeyRelease>", self.update_prompt)
        self.prompt_text.bind("<FocusOut>", self.update_prompt)

        # --- Import/Export/Scratchpad Buttons Frame ---
        self.import_export_frame = ttk.Frame(self.root, style='TFrame') # Use styled frame
        self.import_export_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.import_export_frame, text="Import Prompt", command=self.import_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(self.import_export_frame, text="Export Prompt", command=self.export_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        # --- Scratchpad Button ---
        self.scratchpad_button = ttk.Button(self.root, text="Scratchpad", command=self.open_scratchpad_window)
        self.scratchpad_button.pack(pady=(5, 10), padx=10, fill=tk.X, ipady=10)

        # --- OK Button ---
        self.ok_button = ttk.Button(self.root, text="OK (Hide Window)", command=self.store_settings_and_hide)
        self.ok_button.pack(pady=(0, 10), padx=10, fill=tk.X, ipady=10)

        # --- Start/Stop Recording Button and Indicator ---
        # Use a normal tk.Frame here if ttk.Frame causes issues with indicator label bg
        self.start_stop_frame = tk.Frame(self.root, bg=self.default_bg)
        self.start_stop_frame.pack(pady=0, padx=10, fill=tk.X) # Reduce bottom padding here
        self.start_stop_button = ttk.Button(self.start_stop_frame, text="Start Recording", command=self.toggle_recording)
        self.start_stop_button.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=10)
        self.recording_indicator_label = ttk.Label(self.start_stop_frame, text="●", font=("Arial", 12), style='TLabel') # Apply style
        self.recording_indicator_label.pack(side=tk.LEFT, padx=5)
        self.update_recording_indicator()

        # --- Bottom Frame (Hotkeys, Queue Controls, Queue Indicator) ---
        bottom_frame = ttk.Frame(self.root, style='TFrame') # Use styled frame
        # Reduced top padding here, bottom padding remains 10
        bottom_frame.pack(pady=(5, 10), padx=10, fill=tk.X, side=tk.BOTTOM, anchor=tk.S)

        # Hotkey Label (Left) - Use StringVar for dynamic updates
        ttk.Label(bottom_frame, textvariable=self.shortcut_display_var, justify=tk.LEFT, style='TLabel').pack(side=tk.LEFT, anchor=tk.W) # Apply style
        self.update_shortcut_display() # Set initial text

        # Queue Indicator (Right)
        self.queue_indicator_label = ttk.Label(bottom_frame, textvariable=self.queue_indicator_var, justify=tk.RIGHT, style='TLabel') # Apply style
        self.queue_indicator_label.pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))

        # Clear Queue Button (Right, before indicator)
        self.clear_queue_button = ttk.Button(bottom_frame, text="Clear Queue", command=self.clear_transcription_queue, width=12)
        self.clear_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))

        # Pause Queue Button (Right, before clear button)
        self.pause_queue_button = ttk.Button(bottom_frame, text="Pause Queue", command=self.toggle_queue_processing, width=15)
        self.pause_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))

        # Hotkey registration moved to update_hotkeys() method


    # ... (Update methods for language, model, toggles, prompt remain the same) ...
    def update_language(self):
        self.language = self.language_combobox.get()
        self.log_message("essential", f"Language updated to: {self.language}")
        self.save_settings()

    def update_model(self):
        self.model = self.model_combobox.get()
        self.log_message("essential", f"Model updated to: {self.model}")
        self.save_settings()

    def update_translation(self, *args):
        self.translation_enabled = self.translation_var.get()
        self.log_message("essential", f"Translation enabled: {self.translation_enabled}")
        self.save_settings()

    def update_command_mode(self, *args):
        self.command_mode = self.command_mode_var.get()
        self.vad_enabled = self.command_mode # vad_enabled mirrors command_mode
        self.log_message("essential", f"Command mode / VAD enabled: {self.command_mode}")
        self.save_settings()

    def update_timestamps_disabled(self, *args):
        self.timestamps_disabled = self.timestamps_disabled_var.get()
        self.log_message("essential", f"Timestamps disabled: {self.timestamps_disabled}")
        self.save_settings()

    def update_clear_text_output(self, *args):
        self.clear_text_output = self.clear_text_output_var.get()
        self.log_message("essential", f"Clear text output: {self.clear_text_output}")
        self.save_settings()

    def update_prompt(self, event=None):
        if hasattr(self, 'prompt_text') and self.prompt_text.winfo_exists():
            new_prompt = self.prompt_text.get("1.0", tk.END).strip()
            if new_prompt != self.prompt:
                self.prompt = new_prompt
                self.log_message("essential", f"Prompt updated (length: {len(self.prompt)} chars)")
                # Debounce saving? For now, save on change/focus out.
                self.save_prompt()
        else:
             # Fallback if widget destroyed? Unlikely for prompt.
             pass


    # --- Hotkey Management ---
    def update_hotkeys(self):
        """Unregisters old hotkeys and registers new ones based on settings."""
        self.log_message("extended", "Updating hotkeys...")

        # 1. Unregister previously registered hotkeys
        keys_to_remove = list(self._registered_hotkeys.keys())
        for key_string in keys_to_remove:
            callback = self._registered_hotkeys.pop(key_string, None)
            if callback:
                try:
                    keyboard.remove_hotkey(key_string)
                    self.log_message("extended", f"Unregistered hotkey: {key_string}")
                except Exception as e:
                    # remove_hotkey might fail if key wasn't actually registered by us or library changed
                    self.log_message("essential", f"Error unregistering hotkey '{key_string}': {e}")

        # Clear just in case
        self._registered_hotkeys = {}

        # 2. Register new hotkeys
        new_hotkeys = {
            self.hotkey_toggle_record: self.toggle_recording,
            self.hotkey_show_window: self.return_to_window
        }

        registration_errors = []
        for key_string, callback in new_hotkeys.items():
            if not key_string or not isinstance(key_string, str):
                self.log_message("essential", f"Skipping invalid hotkey string: {key_string}")
                continue
            try:
                keyboard.add_hotkey(key_string, callback, suppress=False) # suppress=False allows key combo to pass through
                self._registered_hotkeys[key_string] = callback
                self.log_message("essential", f"Registered hotkey: {key_string}")
            except ValueError as ve:
                 # keyboard library raises ValueError for invalid syntax
                 error_msg = f"Invalid hotkey syntax: '{key_string}'. Error: {ve}"
                 self.log_message("essential", error_msg)
                 registration_errors.append(error_msg)
            except Exception as e:
                 error_msg = f"Failed to register hotkey '{key_string}': {e}"
                 self.log_message("essential", error_msg)
                 registration_errors.append(error_msg)

        # 3. Update the display label
        self.update_shortcut_display()

        # 4. Show errors if any occurred during registration
        if registration_errors and (not self.config_window or not self.config_window.winfo_exists()):
             # Show error only if config window is not open (it handles errors on save)
             messagebox.showwarning("Hotkey Error", "Could not register one or more hotkeys. Please check syntax in Configuration.\n\nDetails:\n- " + "\n- ".join(registration_errors))
        elif registration_errors and self.config_window and self.config_window.winfo_exists():
             # Let the config window save handle the error message if it's open
             pass

        return not bool(registration_errors) # Return True if successful, False otherwise


    def update_shortcut_display(self):
        """Updates the text variable for the shortcut display label."""
        if hasattr(self, 'shortcut_display_var'):
            toggle_key = self.hotkey_toggle_record or "[Not Set]"
            show_key = self.hotkey_show_window or "[Not Set]"
            shortcut_text = f"{toggle_key}: Toggle Record\n{show_key}: Show Window"
            self.shortcut_display_var.set(shortcut_text)


    # --- Tray Icon Logic ---
    def setup_tray_icon_thread(self):
        # ... (Finding icon remains the same) ...
        try:
            icon_path = None
            possible_paths = []
            if hasattr(sys, '_MEIPASS'):
                possible_paths.append(os.path.join(sys._MEIPASS, 'icon.png'))
            possible_paths.append('icon.png')
            script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else '.'
            possible_paths.append(os.path.join(script_dir, 'icon.png'))

            for path in possible_paths:
                norm_path = os.path.normpath(path)
                if os.path.exists(norm_path):
                    icon_path = norm_path
                    self.log_message("extended", f"Found icon at: {icon_path}")
                    break

            if not icon_path:
                self.log_message("essential", "Warning: icon.png not found. Using fallback gray icon.")
                image = Image.new('RGB', (64, 64), color='gray')
            else:
                try:
                    image = Image.open(icon_path)
                except Exception as img_e:
                    self.log_message("essential", f"Error loading icon image '{icon_path}': {img_e}. Using fallback.")
                    image = Image.new('RGB', (64, 64), color='gray')

            menu = pystray.Menu(
                pystray.MenuItem("Show WhisperR", self.show_window_action, default=True),
                pystray.MenuItem("Toggle Recording", self.toggle_recording_action), # Keep this action
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit WhisperR", self.quit_app_action)
            )
            self.tray_icon = pystray.Icon("WhisperR", image, "WhisperR", menu)
            self.log_message("essential", "Running tray icon...")
            self.tray_icon.run()
            self.log_message("essential", "Tray icon thread finished.")
        except Exception as e:
            self.log_message("essential", f"Error setting up tray icon: {e}")
            # Don't necessarily quit if tray fails
            self.tray_icon = None # Ensure tray icon is None if setup failed

    def show_window_action(self, icon=None, item=None):
        self.root.after(0, self._show_window)

    def toggle_recording_action(self, icon=None, item=None):
         # Call the main toggle method, ensuring it runs on the main thread if needed
         self.root.after(0, self.toggle_recording)

    # --- Window Management ---
    def handle_close_button(self):
        self.log_message("extended", f"Close button clicked. Behavior: {self.close_behavior}")
        if self.close_behavior == "Minimize to tray":
            self.root.withdraw()
            if self.tray_icon and hasattr(self.tray_icon, 'visible') and self.tray_icon.visible:
                 # Show a notification if possible
                 try:
                      self.tray_icon.notify("WhisperR is running in the background.", "WhisperR")
                 except Exception as e_notify:
                      self.log_message("extended", f"Could not show tray notification: {e_notify}")
            else:
                 self.log_message("extended", "Minimized, but no tray icon available for notification.")
        else: # "Exit app"
            self.quit_app_action()

    def quit_app_action(self, icon=None, item=None):
        self.log_message("essential", "Quit action initiated...")

        # Unregister all hotkeys cleanly before exiting
        keys_to_remove = list(self._registered_hotkeys.keys())
        self.log_message("extended", f"Unregistering {len(keys_to_remove)} hotkeys on quit...")
        for key_string in keys_to_remove:
             callback = self._registered_hotkeys.pop(key_string, None)
             if callback:
                 try:
                     keyboard.remove_hotkey(key_string)
                 except Exception: # Ignore errors here as we are quitting anyway
                     pass
        # Alternative, if sure no other app uses 'keyboard':
        # try:
        #     keyboard.unhook_all() # More forceful
        #     self.log_message("extended", "Unhooked all keyboard listeners.")
        # except Exception as e:
        #     self.log_message("essential", f"Error unhooking keyboard: {e}")

        # --- Rest of shutdown sequence ---
        self.recording = False
        if self.audio_stream:
            try:
                if not self.audio_stream.closed:
                     self.audio_stream.stop()
                     self.audio_stream.close()
                     self.log_message("extended", "Audio stream stopped and closed.")
            except Exception as e:
                self.log_message("essential", f"Error stopping/closing audio stream: {e}")

        if self.recording_thread and self.recording_thread.is_alive():
            self.log_message("extended", "Joining recording thread...")
            self.recording_thread.join(timeout=1.0)

        self.log_message("essential", f"Signalling transcription worker to stop. Queue size: {self.transcription_queue.qsize()}")
        self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.log_message("extended", "Joining transcription worker thread...")
            self.transcription_worker_thread.join(timeout=2.0)

        if getattr(self, 'clear_audio_on_exit', False) or getattr(self, 'clear_text_on_exit', False):
            self.log_message("essential", "Performing file cleanup on exit...")
            self.delete_session_files(ask_confirm=False)

        if self.tray_icon and hasattr(self.tray_icon, 'stop'):
            self.log_message("essential", "Stopping tray icon...")
            try:
                self.tray_icon.stop()
            except Exception as e:
                self.log_message("essential", f"Error stopping tray icon: {e}")
            self.tray_icon = None # Ensure it's cleared

        self.log_message("extended", "Saving final settings...")
        self.save_settings()
        self.save_prompt()
        self.save_commands()

        if self.log_file:
            try:
                self.log_message("extended", "Closing log file...")
                self.log_file.close()
                self.log_file = None
            except Exception as e:
                self.log_message("essential", f"Error closing log file: {e}")

        self.log_message("essential", "Scheduling root quit...")
        # Ensure root destroy happens cleanly
        if self.root.winfo_exists():
             self.root.after(50, self.root.destroy) # Use destroy instead of quit
        else:
             # If root somehow already destroyed, just exit process
             sys.exit(0)


    # --- Green Line Indicator ---
    def create_green_line(self):
        # ... (remains the same) ...
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            hwnd = win32gui.CreateWindowEx(win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED | win32con.WS_EX_TOOLWINDOW,
                                          "Static", None, win32con.WS_VISIBLE | win32con.WS_POPUP, 0, 0, screen_width, 5, 0, 0, 0, None)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
            win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0,0,0), 100, win32con.LWA_ALPHA) # Example: 100 alpha
            hdc = win32gui.GetDC(hwnd)
            brush = win32gui.CreateSolidBrush(win32api.RGB(0, 150, 0)) # Green color
            rect = win32gui.GetClientRect(hwnd)
            win32gui.FillRect(hdc, rect, brush)
            win32gui.ReleaseDC(hwnd, hdc)
            win32gui.DeleteObject(brush)
            self.green_line = hwnd
            self.log_message("extended", "Green line created.")
        except Exception as e:
            self.log_message("essential", f"Error creating green line: {e}")

    def destroy_green_line(self):
        # ... (remains the same) ...
        if self.green_line:
            try:
                win32gui.DestroyWindow(self.green_line)
                self.green_line = None
                self.log_message("extended", "Green line destroyed.")
            except Exception as e:
                if hasattr(e, 'winerror') and e.winerror == 1400:
                    self.log_message("extended", "Green line window already destroyed.")
                else:
                    self.log_message("essential", f"Error destroying green line: {e}")
                self.green_line = None


    # --- Recording Logic ---
    def toggle_recording(self):
        # ... (remains the same) ...
        self.log_message("essential", "Toggle recording requested.")
        if self.recording:
            self.root.after(0, self.stop_recording)
        else:
            self.root.after(0, self.start_recording)

    def start_recording(self):
        if self.recording:
            return
        self.log_message("essential", "Starting recording...")
        self.recording = True
        # self.audio_buffer = deque() # ***FIX: Removed unused audio_buffer***
        self.current_segment = []
        self.is_speaking = False
        self.silence_start_time = None
        self.create_green_line()
        if self.recording_thread and self.recording_thread.is_alive():
            self.log_message("extended", "Waiting for previous recording thread to finish...")
            self.recording_thread.join()
        self.log_message("extended", "Starting new recording thread.")
        self.recording_thread = threading.Thread(target=self.record_audio_continuously, daemon=True)
        self.recording_thread.start()
        self.root.after(0, self.update_recording_indicator)

    def stop_recording(self):
        if not self.recording:
            return
        self.log_message("essential", "Manual stop recording requested...")
        self.recording = False # This will signal the recording_thread to stop its loop
        self.root.after(0, self.destroy_green_line)

        # ***FIX: Changed condition from `self.audio_buffer or self.current_segment` to just `self.current_segment` ***
        if self.current_segment:
            self.log_message("extended", "Processing final audio buffer on manual stop...")
            # self.current_segment already holds the data, no need to copy to final_segment_data here.
            # save_segment_and_reset_vad will handle self.current_segment
            self.root.after(0, self.save_segment_and_reset_vad)
        else:
            self.log_message("extended", "No audio in current_segment to save on manual stop.")
            # Ensure VAD state is clean even if no audio was captured.
            # start_recording also resets these, but good for explicit cleanup.
            self.current_segment = [] # Should be empty already
            self.is_speaking = False
            self.silence_start_time = None

        # self.audio_buffer = deque() # ***FIX: Removed unused audio_buffer***
        self.log_message("essential", f"Recording stopped. Transcription queue size: {self.transcription_queue.qsize()}")
        self.root.after(0, self.update_recording_indicator)


    def update_recording_indicator(self):
        # ... (remains the same) ...
        if hasattr(self, 'recording_indicator_label'):
            color = "red" if self.recording else "gray"
            text = "Stop Recording" if self.recording else "Start Recording"
            # Use winfo_exists checks
            if hasattr(self, 'start_stop_button') and self.start_stop_button.winfo_exists():
                 self.start_stop_button.config(text=text)
            if hasattr(self, 'recording_indicator_label') and self.recording_indicator_label.winfo_exists():
                 self.recording_indicator_label.config(foreground=color)
            self.log_message("extended", f"Recording indicator updated (Recording: {self.recording}).")


    def record_audio_continuously(self):
        self.log_message("essential", "Continuous audio recording thread started.")
        samplerate = 44100
        channels = 1
        dtype = 'int16'
        blocksize = 1024 # Number of frames per callback
        device_index = getattr(self, 'selected_audio_device_index', None)

        # VAD specific parameters (only used if self.command_mode is True)
        silence_duration = getattr(self, 'silence_threshold_seconds', 5.0)
        energy_threshold = getattr(self, 'vad_energy_threshold', 300)

        # ***FIX: Renamed vad_buffer_limit to vad_mode_max_segment_chunks for clarity***
        # Max chunks for a single segment *only when VAD mode is active*.
        # Prevents overly long segments if silence detection is problematic or speech is continuous.
        # Approx 60 seconds: (samplerate frames/sec * 60 seconds) / (blocksize frames/chunk) = chunks
        vad_mode_max_segment_chunks = int(samplerate * 60 / blocksize)

        def audio_callback(indata, frames, time_info, status):
            if status:
                # Avoid logging directly from callback if causing issues
                # Consider queueing log messages for main thread if problems arise
                print(f"Audio CB Status: {status}") # Use print as fallback log

            # This check is crucial. If recording is stopped, do nothing further.
            if not self.recording:
                return

            current_time_monotonic = time.monotonic()
            data_copy = indata.copy()
            self.current_segment.append(data_copy) # Always append incoming audio data

            # --- VAD Logic (Auto-Pause / Commands mode) ---
            if self.command_mode: # self.command_mode is True when "Enable Auto-Pause / Commands" is checked
                rms = np.sqrt(np.mean(data_copy.astype(np.float32)**2))
                is_currently_loud = rms >= energy_threshold

                if self.logging_level == "Everything":
                    silence_status = "Silence Detected" if self.silence_start_time else "No Silence"
                    speaking_status = "Speaking" if self.is_speaking else "Not Speaking"
                    log_msg = f"RMS: {rms:.2f} (Thresh: {energy_threshold}) | Loud: {is_currently_loud} | State: {speaking_status} | {silence_status}"
                    print(f"[VAD DEBUG] {log_msg}")

                if is_currently_loud:
                    if not self.is_speaking:
                        if self.logging_level == "Everything": print("[VAD DEBUG] Speech started.")
                        self.is_speaking = True
                    self.silence_start_time = None # Reset silence timer if speech is detected
                elif self.is_speaking: # Was speaking, but current chunk is not loud
                    if self.silence_start_time is None:
                        if self.logging_level == "Everything": print("[VAD DEBUG] Silence started...")
                        self.silence_start_time = current_time_monotonic

                    # Check if silence duration has been met
                    if current_time_monotonic - self.silence_start_time >= silence_duration:
                        if self.logging_level == "Everything": print(f"[VAD DEBUG] Silence threshold ({silence_duration}s) reached.")
                        if self.recording: # Ensure still in recording state before scheduling save
                            self.root.after(0, self.save_segment_and_reset_vad)
                            # VAD state (is_speaking, silence_start_time) is reset by save_segment_and_reset_vad

                # ***FIX: Moved max segment length check INSIDE the `if self.command_mode:` block***
                # This check is now specific to VAD mode.
                if len(self.current_segment) > vad_mode_max_segment_chunks:
                    actual_duration_s = len(self.current_segment) * blocksize / samplerate
                    self.log_message("essential", f"VAD mode segment limit triggered (current segment approx {actual_duration_s:.1f}s), saving.")
                    if self.recording: # Ensure still in recording state
                        self.root.after(0, self.save_segment_and_reset_vad)
                        # VAD state also reset by save_segment_and_reset_vad

            # else: (self.command_mode is False)
            #   In "record forever" mode, self.current_segment just keeps accumulating.
            #   No automatic saving based on silence or segment length occurs here.
            #   The entire recording will be saved only when self.stop_recording() is manually called.

        # --- Stream setup and loop ---
        self.audio_stream = None
        try:
            self.log_message("extended", f"Opening InputStream (Continuous) device: {device_index}, blocksize: {blocksize}")
            self.audio_stream = sd.InputStream(
                device=device_index,
                samplerate=samplerate,
                channels=channels,
                dtype=dtype,
                blocksize=blocksize,
                callback=audio_callback
            )
            self.log_message("extended", "Starting audio stream...")
            self.audio_stream.start()
            self.log_message("essential", "Audio stream started.")

            # Main loop for the recording thread. Keeps running as long as self.recording is True.
            # The audio_callback handles data processing.
            while self.recording:
                time.sleep(0.1) # Keep thread alive, check self.recording periodically

        except sd.PortAudioError as pae:
            self.log_message("essential", f"PortAudioError in recording thread: {pae}")
            self.root.after(0, lambda: messagebox.showerror("Audio Error", f"Failed to open audio device: {pae}\n\nPlease check the selected device in Configuration."))
            self.root.after(0, self.handle_recording_error) # This will set self.recording = False
        except Exception as e:
            self.log_message("essential", f"Unexpected error in continuous recording thread: {e}")
            import traceback
            self.log_message("essential", traceback.format_exc())
            self.root.after(0, self.handle_recording_error) # This will set self.recording = False
        finally:
            if self.audio_stream:
                try:
                    if not self.audio_stream.closed:
                        self.log_message("extended", "Stopping and closing audio stream in finally block.")
                        self.audio_stream.stop()
                        self.audio_stream.close()
                except Exception as e_close:
                    self.log_message("essential", f"Error closing audio stream: {e_close}")

            self.log_message("essential", f"Continuous audio recording thread finished (self.recording is now {self.recording}).")

            # If the thread exited for some reason while self.recording was still True
            # (e.g., an unhandled exception in the while loop or stream closed unexpectedly),
            # ensure the application's recording state is correctly updated.
            if self.recording:
                 self.log_message("extended","Recording thread finished, but self.recording was still True. Forcing stop_recording.")
                 self.root.after(0, self.stop_recording) # This will set self.recording = False and process any final segment


    def save_segment_and_reset_vad(self):
        if not self.current_segment:
            self.log_message("extended", "Save segment called but no segment data.")
            return

        segment_to_save = list(self.current_segment) # Make a copy of the list of audio chunks
        self.current_segment = [] # Reset for the next segment or for a clean state

        # Reset VAD state variables. This is crucial for VAD mode.
        # For "record forever" mode, these are reset by start_recording if a new recording begins.
        # This ensures they are clean after any segment save.
        self.is_speaking = False
        self.silence_start_time = None

        self.save_segment(segment_to_save) # segment_to_save is a list of numpy arrays


    def save_segment(self, segment_data):
        # ... (remains the same, maybe add exist_ok=True to makedirs) ...
        if not segment_data:
            self.log_message("extended", "No segment data provided to save.")
            return

        self.log_message("extended", f"Saving segment with {len(segment_data)} chunks at {datetime.datetime.now()}")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"recording_{timestamp}.wav"

        export_dir = getattr(self, 'export_folder', '.')
        norm_export_dir = os.path.normpath(export_dir)
        # Simplified directory check/creation
        try:
            os.makedirs(norm_export_dir, exist_ok=True)
        except OSError as e:
             self.log_message("essential", f"Error creating export directory '{norm_export_dir}': {e}. Saving may fail.")
             # Decide whether to fallback or proceed
             # norm_export_dir = '.' # Fallback example

        filepath = os.path.normpath(os.path.join(norm_export_dir, filename))

        try:
            if not segment_data:
                raise ValueError("Segment data empty before concat.")
            audio_array = np.concatenate(segment_data, axis=0)
            if audio_array.size == 0:
                raise ValueError("Concatenated segment empty.")

            with wave.open(filepath, 'wb') as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2) # Assumes int16
                wf.setframerate(44100)
                wf.writeframes(audio_array.tobytes())

            self.log_message("essential", f"Segment saved to {filepath}")

            # <<< PLAY BEEP IF ENABLED FOR SEGMENT SAVE >>>
            if self.beep_on_save_audio_segment:
                # Ensure play_beep is called from the main thread if save_segment can be called from other threads.
                # Currently, save_segment (via save_segment_and_reset_vad) is scheduled with self.root.after(0, ...),
                # so it runs on the main thread, making this direct call safe.
                self.play_beep()
            # <<< END BEEP LOGIC >>>

            self.transcription_queue.put(filepath)
            self.root.after(0, self.update_queue_indicator)

        except ValueError as e:
            self.log_message("essential", f"Error processing segment data: {e}")
        except wave.Error as e:
             self.log_message("essential", f"Error writing wave file '{filepath}': {e}")
        except Exception as e:
            self.log_message("essential", f"Unexpected error saving segment '{filepath}': {e}")


    # --- Transcription Worker ---
    def transcription_worker(self):
        # ... (Pause/Clear logic remains the same) ...
        self.log_message("essential", "Transcription worker thread started.")
        while True:
            try:
                # --- Handle Pause ---
                while self.queue_processing_paused:
                    if not hasattr(self, 'pause_checked_time') or time.monotonic() - self.pause_checked_time > 5:
                         self.log_message("extended", "Queue processing paused. Waiting...")
                         self.pause_checked_time = time.monotonic()
                    time.sleep(0.5)

                # --- Get Item ---
                try:
                    filepath = self.transcription_queue.get(timeout=1.0)
                except queue.Empty:
                    continue

                if filepath is AUDIO_QUEUE_SENTINEL:
                    self.log_message("essential", "Worker received sentinel. Stopping.")
                    self.transcription_queue.task_done()
                    break

                # --- Handle Clear Flag ---
                if self.clear_queue_flag:
                    self.log_message("extended", f"Worker discarding item due to clear flag: {filepath}")
                    self.transcription_queue.task_done()
                    self.root.after(0, self.update_queue_indicator)
                    continue

                # --- Process Item ---
                norm_filepath = os.path.normpath(filepath)
                self.log_message("essential", f"Worker processing: {norm_filepath}. Queue size: {self.transcription_queue.qsize()}")
                self.root.after(0, self.update_queue_indicator) # Update before processing

                if os.path.exists(norm_filepath) and not norm_filepath.endswith(".transcribed"):
                    self.transcribe_audio(norm_filepath)
                else:
                    self.log_message("extended", f"Worker skipping: {norm_filepath} (Not found or already transcribed)")

                self.transcription_queue.task_done()
                self.root.after(0, self.update_queue_indicator) # Update after processing
                self.log_message("extended", f"Worker finished {norm_filepath}. Remaining: {self.transcription_queue.qsize()}")

            except Exception as e:
                # ... (Error handling remains the same) ...
                log_path = 'unknown file'
                if 'filepath' in locals() and filepath is not AUDIO_QUEUE_SENTINEL:
                    log_path = os.path.normpath(filepath)
                    # Ensure task_done is called even on error if item was retrieved
                    try:
                         self.transcription_queue.task_done()
                    except ValueError: # May happen if already marked done
                         pass
                    self.root.after(0, self.update_queue_indicator)

                self.log_message("essential", f"Error in transcription worker loop for {log_path}: {e}")
                import traceback
                self.log_message("essential", traceback.format_exc())
                time.sleep(1)


    def start_transcription_worker(self):
        # ... (remains the same) ...
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.log_message("extended", "Transcription worker thread already running.")
            return
        self.transcription_worker_thread = threading.Thread(target=self.transcription_worker, daemon=True)
        self.transcription_worker_thread.start()
        self.log_message("essential", "Transcription worker thread started/restarted.")

    def handle_recording_error(self):
        # ... (remains the same) ...
        self.log_message("essential", "Handling recording error state.")
        if self.recording: # If it was recording, stop it
            self.recording = False # This signals the recording thread to stop
            self.root.after(0, self.destroy_green_line)
            self.root.after(0, self.update_recording_indicator)
        if self.audio_stream and not self.audio_stream.closed:
             try:
                  self.log_message("extended", "Attempting to close stream after recording error.")
                  self.audio_stream.stop()
                  self.audio_stream.close()
             except Exception as e:
                  self.log_message("essential", f"Error closing stream during error handling: {e}")

    # --- Transcription Execution ---
    def transcribe_audio(self, audio_path):
        # ... (remains largely the same, ensure prompt retrieval is robust) ...
        self.log_message("essential", f"Transcription task started for: {audio_path}")
        # Get current settings safely
        language = self.language_combobox.get() if hasattr(self, 'language_combobox') and self.language_combobox.winfo_exists() else getattr(self, 'language', 'en')
        model = self.model_combobox.get() if hasattr(self, 'model_combobox') and self.model_combobox.winfo_exists() else getattr(self, 'model', 'large')
        translation_enabled = self.translation_var.get() if hasattr(self, 'translation_var') else getattr(self, 'translation_enabled', False)
        command_mode = self.command_mode_var.get() if hasattr(self, 'command_mode_var') else getattr(self, 'command_mode', False)
        timestamps_disabled = self.timestamps_disabled_var.get() if hasattr(self, 'timestamps_disabled_var') else getattr(self, 'timestamps_disabled', False)
        clear_text_output = self.clear_text_output_var.get() if hasattr(self, 'clear_text_output_var') else getattr(self, 'clear_text_output', False)
        export_dir = os.path.normpath(getattr(self, 'export_folder', '.'))
        whisper_exec = os.path.normpath(getattr(self, 'whisper_executable', 'whisper'))

        # Get prompt safely
        prompt_content = ""
        try:
            if hasattr(self, 'prompt_text') and self.prompt_text.winfo_exists():
                 prompt_content = self.prompt_text.get("1.0", tk.END).strip()
            else:
                 prompt_content = getattr(self, 'prompt', '')
        except Exception as e_prompt:
            self.log_message("essential", f"Error getting prompt content: {e_prompt}. Using stored value.")
            prompt_content = getattr(self, 'prompt', '')


        norm_audio_path = os.path.normpath(audio_path)

        try:
            os.makedirs(export_dir, exist_ok=True)
        except Exception as e:
            self.log_message("essential", f"Failed create export directory '{export_dir}': {e}. Whisper might fail.")

        output_filename_base = f"{os.path.splitext(os.path.basename(norm_audio_path))[0]}"
        norm_output_filepath_txt = os.path.normpath(os.path.join(export_dir, f"{output_filename_base}.txt"))

        # Build command
        command = [whisper_exec, norm_audio_path, "--model", model, "--language", language, "--output_format", "txt", "--output_dir", export_dir]
        if prompt_content:
            # Basic quoting for prompt - may need platform-specific handling
            prompt_cmd = f'"{prompt_content}"' if '"' not in prompt_content else prompt_content
            command.extend(["--initial_prompt", prompt_cmd])

        command.extend(["--task", "translate" if translation_enabled else "transcribe"])

        transcription_successful = False
        parsed_text = ""
        try:
            command_str_list = [str(c) for c in command]
            cmd_display = subprocess.list2cmdline(command_str_list) # Safer way to display command for logging
            self.log_message("extended", f"Running Whisper command: {cmd_display}")

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(command_str_list, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')

            if os.path.exists(norm_output_filepath_txt):
                self.log_message("essential", f"Transcription text file created: {norm_output_filepath_txt}")
                transcription_successful = True
                parsed_text = self.parse_transcription_text(norm_output_filepath_txt)

                # --- Schedule UI updates/actions on main thread ---
                self.root.after(0, self.update_scratchpad_text, parsed_text)
                self.root.after(0, self._update_clipboard, parsed_text)

                if self.beep_on_transcription:
                    self.root.after(0, self.play_beep)

                if self.auto_paste:
                    delay_ms = int(max(0, self.auto_paste_delay) * 1000) # Ensure non-negative delay
                    self.root.after(delay_ms, self.perform_auto_paste)

                if command_mode: # Use the current command_mode setting for this transcribed segment
                    self.log_message("extended", f"Processing commands with transcription: '{parsed_text}'")
                    current_commands_snapshot = list(self.loaded_commands) # Take snapshot
                    self.root.after(0, self.execute_command_from_text, parsed_text, current_commands_snapshot)
                # --- End scheduled actions ---
            else:
                self.log_message("essential", f"Whisper command finished, but output file '{norm_output_filepath_txt}' not found.")
                self.log_message("extended", f"Whisper stdout: {result.stdout}")
                self.log_message("extended", f"Whisper stderr: {result.stderr}")

        except subprocess.CalledProcessError as e:
            # ... (error handling as before) ...
            self.log_message("essential", f"Error running whisper: {e}")
            self.log_message("essential", f"Stdout: {e.stdout}")
            self.log_message("essential", f"Stderr: {e.stderr}")
            if "ModuleNotFoundError" in e.stderr or "'whisper' is not recognized" in e.stderr:
                 self.log_message("essential", "Hint: Whisper might not be installed correctly or not in the system's PATH.")
                 self.root.after(0, lambda: messagebox.showerror("Whisper Error", "Whisper command failed. It might not be installed or configured correctly. Check the Configuration and logs."))
        except FileNotFoundError:
            self.log_message("essential", f"Error: Whisper executable not found at '{whisper_exec}'. Please check the path in Configuration.")
            self.root.after(0, lambda: messagebox.showerror("Whisper Error", f"Whisper executable not found at:\n{whisper_exec}\n\nPlease correct the path in Configuration."))
        except Exception as e:
            self.log_message("essential", f"Unexpected transcription error: {e}")
            import traceback
            self.log_message("essential", traceback.format_exc())

        finally:
            # ... (renaming logic as before) ...
            if transcription_successful and os.path.exists(norm_audio_path):
                try:
                    transcribed_path = norm_audio_path + ".transcribed"
                    norm_transcribed_path = os.path.normpath(transcribed_path)
                    shutil.move(norm_audio_path, norm_transcribed_path)
                    self.log_message("extended", f"Renamed audio file to: {norm_transcribed_path}")
                except Exception as e_mv:
                    self.log_message("essential", f"Error renaming audio file {norm_audio_path} to {norm_transcribed_path}: {e_mv}")
            elif os.path.exists(norm_audio_path):
                 self.log_message("extended", f"Audio file {norm_audio_path} kept (transcription successful: {transcription_successful}).")

        self.log_message("essential", f"Transcription task finished for: {norm_audio_path}")


    def play_beep(self):
        # ... (remains the same) ...
        try:
            winsound.Beep(500, 200)
            self.log_message("extended", "Played beep.")
        except Exception as e:
            self.log_message("essential", f"Error playing beep sound: {e}")


    def perform_auto_paste(self):
        # ... (remains the same) ...
        try:
            # Add short delay before paste?
            # time.sleep(0.05)
            keyboard.press_and_release('ctrl+v')
            self.log_message("extended", "Performed auto-paste (Ctrl+V).")
        except Exception as e:
            self.log_message("essential", f"Error performing auto-paste: {e}")

    # --- Text Parsing & Command Execution ---
    def parse_transcription_text(self, filepath):
        # ... (remains the same) ...
        self.log_message("extended", f"Parsing transcription file: {filepath}")
        full_text = ""
        try:
            with open(filepath, "r", encoding='utf-8') as f:
                full_text = f.read()
            # Log raw only if verbose?
            # if self.logging_level == "Everything":
            #      self.log_message("everything", f"Raw transcription text:\n---\n{full_text}\n---")

            if not full_text.strip():
                self.log_message("extended", "Transcription file is empty.")
                return ""

            if not self.clear_text_output and not self.timestamps_disabled:
                self.log_message("everything", "Returning raw text (no cleaning options enabled).")
                return full_text.strip()

            lines = full_text.splitlines()
            cleaned_lines = []
            timestamp_pattern = re.compile(r'^\[\s*\d{2}:\d{2}\.\d{3}\s*-->\s*\d{2}:\d{2}\.\d{3}\s*\]\s*')

            for line in lines:
                line_stripped = line.strip()
                if not line_stripped:
                    continue

                if self.clear_text_output and (line_stripped.startswith("---") or line_stripped.startswith("===") or re.match(r'^\d{4}-\d{2}-\d{2}', line_stripped)):
                    # self.log_message("everything", f"Skipping potential header/meta line: '{line_stripped}'")
                    continue

                match = timestamp_pattern.match(line_stripped)
                if match:
                     if self.timestamps_disabled or self.clear_text_output:
                         text_part = line_stripped[match.end():].strip()
                         if text_part:
                             cleaned_lines.append(text_part)
                             # self.log_message("everything", f"Extracted text after timestamp: '{text_part}'")
                     else:
                         cleaned_lines.append(line_stripped)
                         # self.log_message("everything", f"Keeping line with timestamp: '{line_stripped}'")
                else:
                    cleaned_lines.append(line_stripped)
                    # self.log_message("everything", f"Keeping non-timestamp line: '{line_stripped}'")


            cleaned_text = "\n".join(cleaned_lines).strip()
            self.log_message("extended", f"Parsed transcription result (length: {len(cleaned_text)}).")
            # Log full parsed only if verbose?
            # if self.logging_level == "Everything":
            #      self.log_message("everything", f"Parsed text:\n---\n{cleaned_text}\n---")
            return cleaned_text

        except FileNotFoundError:
            self.log_message("essential", f"File not found during parsing: {filepath}")
            return ""
        except Exception as e:
            self.log_message("essential", f"Error parsing transcription file '{filepath}': {e}")
            return full_text.strip() # Fallback


    def execute_command_from_text(self, transcription_text, commands_list):
        # ... (remains the same) ...
        self.log_message("extended", f"Attempting command execution based on text: '{transcription_text}'")
        if not transcription_text.strip():
            self.log_message("extended", "Transcription text is empty, skipping command check.")
            return
        if not commands_list:
            self.log_message("extended", "No commands provided for checking.")
            return

        try:
            cleaned_transcription = transcription_text.lower().strip()
            # self.log_message("everything", f"Normalized text for command matching: '{cleaned_transcription}'")

            for command_data in commands_list:
                voice_cmd = command_data.get("voice", "").strip().lower()
                action_template = command_data.get("action", "").strip()

                if not voice_cmd or not action_template:
                    continue

                escaped_voice_cmd = re.escape(voice_cmd)
                pattern_str_core = escaped_voice_cmd.replace(re.escape(" ff "), r'(.*?)')

                # Use word boundaries more carefully - allow non-word chars at ends
                prefix = r'\b' if pattern_str_core and pattern_str_core[0].isalnum() else ''
                suffix = r'\b' if pattern_str_core and pattern_str_core[-1].isalnum() else ''
                # Match anywhere, allow optional trailing punctuation
                pattern_str = prefix + pattern_str_core + suffix + r'[.,]?$'

                # self.log_message("everything", f"Trying pattern: r'{pattern_str}' on '{cleaned_transcription}'")
                match = re.search(pattern_str, cleaned_transcription, re.IGNORECASE)

                if match:
                    self.log_message("essential", f"Command '{voice_cmd}' matched!")
                    action = action_template
                    if r'(.*?)' in pattern_str_core and match.groups():
                        wildcard_value = match.group(1).strip()
                        self.log_message("extended", f"Wildcard ' FF ' matched value: '{wildcard_value}'")
                        action = action.replace(" FF ", wildcard_value)

                    self.log_message("essential", f"Final action to execute: '{action}'")
                    threading.Thread(target=self.run_subprocess, args=(action,), daemon=True).start()
                    return

        except Exception as e:
            self.log_message("essential", f"Error during command execution logic: {e}")


    def run_subprocess(self, action):
        # ... (remains the same) ...
        self.log_message("essential", f"Running subprocess action: '{action}'")
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(action, shell=True, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')
            self.log_message("extended", f"Command '{action}' executed. Output: {result.stdout}")
        except subprocess.CalledProcessError as e:
            self.log_message("essential", f"Command '{action}' failed. Return code: {e.returncode}")
            self.log_message("essential", f"Stderr: {e.stderr}")
            if "is not recognized" in e.stderr.lower() or "cannot find the path" in e.stderr.lower():
                self.log_message("essential", f"Hint: The command or path '{action}' might be incorrect or not in the system's PATH.")
        except FileNotFoundError:
            self.log_message("essential", f"Error: Command or program in '{action}' not found.")
        except Exception as e:
            self.log_message("essential", f"Unexpected error running subprocess '{action}': {e}")


    # --- Clipboard ---
    def _update_clipboard(self, text):
        # ... (remains the same) ...
        if not isinstance(text, str):
             text = str(text)
        try:
            # Check if root window still exists
            if not self.root.winfo_exists():
                 self.log_message("essential", "Cannot update clipboard, root window destroyed.")
                 return
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.log_message("extended", f"Clipboard updated with {len(text)} characters.")
        except tk.TclError as e:
             self.log_message("essential", f"Error updating clipboard (TclError): {e}")
        except Exception as e:
            self.log_message("essential", f"Unexpected error updating clipboard: {e}")


    # --- Versioning / Backup ---
    def create_backup(self, filepath):
        # ... (remains the same) ...
        if not getattr(self, 'versioning_enabled', False) or not os.path.exists(filepath):
            return

        backup_folder = getattr(self, 'backup_folder', 'OldVersions')
        max_backups = getattr(self, 'max_backups', 10)

        if max_backups <= 0:
            return

        norm_backup_folder = os.path.normpath(backup_folder)

        try:
            os.makedirs(norm_backup_folder, exist_ok=True)
        except OSError as e:
            self.log_message("essential", f"Error creating backup folder '{norm_backup_folder}': {e}. Backup skipped.")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)
        backup_filename = f"{name}_{timestamp}{ext}"
        backup_filepath = os.path.join(norm_backup_folder, backup_filename)
        norm_backup_filepath = os.path.normpath(backup_filepath)

        try:
            shutil.copy2(filepath, norm_backup_filepath)
            self.log_message("extended", f"Created backup of '{filename}' to: {norm_backup_filepath}")
            self.manage_backups(norm_backup_folder, name, ext, max_backups)
        except Exception as e:
            self.log_message("essential", f"Error creating backup for '{filepath}': {e}")


    def manage_backups(self, backup_folder, filename_base, file_extension, max_backups):
        # ... (remains the same) ...
        if max_backups <= 0:
            return

        try:
            pattern_str = rf"^{re.escape(filename_base)}_\d{{8}}_\d{{6}}{re.escape(file_extension)}$"
            pattern = re.compile(pattern_str)

            backup_files = []
            for entry in os.listdir(backup_folder):
                full_path = os.path.join(backup_folder, entry)
                if pattern.match(entry) and os.path.isfile(full_path):
                    try:
                        mtime = os.path.getmtime(full_path)
                        backup_files.append((full_path, mtime))
                    except OSError as e:
                        self.log_message("essential", f"Error getting mtime for backup file {full_path}: {e}")

            backup_files.sort(key=lambda x: x[1])
            num_to_delete = len(backup_files) - max_backups

            if num_to_delete > 0:
                files_to_delete = backup_files[:num_to_delete]
                self.log_message("extended", f"Found {len(backup_files)} backups for {filename_base}{file_extension}, limit {max_backups}. Deleting {num_to_delete} oldest.")
                for file_path, _ in files_to_delete:
                    try:
                        os.remove(file_path)
                        self.log_message("extended", f"Deleted old backup: {file_path}")
                    except Exception as e:
                        self.log_message("essential", f"Error deleting old backup {file_path}: {e}")

        except FileNotFoundError:
             self.log_message("essential", f"Backup folder '{backup_folder}' not found during management.")
        except Exception as e:
            self.log_message("essential", f"Error managing backups in '{backup_folder}' for base '{filename_base}': {e}")


    # --- UI Actions ---
    def return_to_window(self):
        self.log_message("essential", "Return to window requested via hotkey.")
        self.root.after(0, self._show_window)

    def _show_window(self):
        # ... (remains the same) ...
        try:
            if self.root.winfo_exists():
                self.log_message("essential", "Showing main window.")
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            else:
                self.log_message("essential", "Cannot show window, root widget destroyed.")
        except tk.TclError as e:
            self.log_message("essential", f"Error showing window (TclError): {e}")

    def store_settings_and_hide(self):
        # ... (remains the same) ...
        self.log_message("essential", "OK button pressed. Hiding window.")
        # Prompt update should happen on keyrelease/focusout
        # self.update_prompt() # Ensure prompt saved if needed
        # Settings are saved from config window now
        # self.save_settings()
        self.root.withdraw()

    def update_audio_device(self):
        # ... (remains the same) ...
        device_index = getattr(self, 'selected_audio_device_index', None)
        if device_index is not None:
            try:
                devices = sd.query_devices()
                if isinstance(device_index, int) and 0 <= device_index < len(devices) and devices[device_index]['max_input_channels'] > 0:
                    sd.default.device = device_index
                    self.log_message("essential", f"Set default audio input device to index: {device_index} ({devices[device_index]['name']})")
                elif isinstance(device_index, int):
                     self.log_message("essential", f"Audio device index {device_index} is invalid or not an input device. Clearing selection.")
                     self.selected_audio_device_index = None
                     sd.default.device = None # Reset default
            except Exception as e:
                self.log_message("essential", f"Failed to query or set audio device (index '{device_index}'): {e}")


    # --- Configuration Window ---
    def open_configuration_window(self):
        if self.config_window and self.config_window.winfo_exists():
            self.config_window.lift()
            self.config_window.focus_force()
            return

        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("WhisperR Configuration")
        # Increased height slightly more for hotkeys
        self.config_window.geometry("450x960") # Adjusted height for new checkbox
        self.config_window.transient(self.root)
        self.config_window.grab_set()

        # Consistent padding options
        pad_options = {'padx': 10, 'pady': 2} # Reduced pady slightly
        frame_options = {'fill': tk.X, 'padx': 10, 'pady': 5} # External padding for LabelFrames
        label_options = {'anchor': tk.W, **pad_options}
        entry_options = {'fill': tk.X, **pad_options}

        # --- Whisper Executable ---
        # Use a standard tk.Frame if ttk.LabelFrame causes bg issues inside
        whisper_outer_frame = tk.Frame(self.config_window)
        whisper_outer_frame.pack(**frame_options)
        tk.Label(whisper_outer_frame, text="Whisper Executable Path:").pack(**label_options)
        whisper_browse_frame = tk.Frame(whisper_outer_frame)
        whisper_browse_frame.pack(fill=tk.X, padx=pad_options['padx'])
        self.whisper_executable_entry = ttk.Entry(whisper_browse_frame)
        self.whisper_executable_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.whisper_executable_entry.insert(0, getattr(self, 'whisper_executable', 'whisper'))
        ttk.Button(whisper_browse_frame, text="Browse...", command=self.browse_whisper_executable).pack(side=tk.LEFT, padx=(5,0))

        # --- Audio Input Device ---
        audio_outer_frame = tk.Frame(self.config_window)
        audio_outer_frame.pack(**frame_options)
        tk.Label(audio_outer_frame, text="Audio Input Device:").pack(**label_options)
        self.device_list_strings = []
        self.device_index_map = {}
        try:
            # ... (device query logic as before) ...
            devices = sd.query_devices()
            default_input_idx = sd.default.device[0] if isinstance(sd.default.device, (list, tuple)) else sd.default.device
            found_current = False
            for i, d in enumerate(devices):
                if d['max_input_channels'] > 0:
                    host_api_name = sd.query_hostapis(d['hostapi'])['name']
                    default_marker = " (System Default)" if i == default_input_idx else ""
                    display_string = f"[{i}] {d['name']} ({host_api_name}){default_marker}"
                    self.device_list_strings.append(display_string)
                    self.device_index_map[display_string] = i
                    if i == self.selected_audio_device_index:
                         found_current = True

            if not self.device_list_strings:
                self.device_list_strings = ["No input devices found"]
        except Exception as e:
            self.log_message("essential", f"Error querying audio devices: {e}")
            self.device_list_strings = ["Error querying devices"]

        self.audio_device_combobox = ttk.Combobox(audio_outer_frame, values=self.device_list_strings, state="readonly")
        # ... (setting combobox value as before) ...
        current_display_string = None
        if self.selected_audio_device_index is not None and found_current:
             for display_str, index in self.device_index_map.items():
                 if index == self.selected_audio_device_index:
                     current_display_string = display_str
                     break
        elif not found_current and self.selected_audio_device_index is not None:
             self.selected_audio_device_index = None

        if current_display_string:
            self.audio_device_combobox.set(current_display_string)
        elif self.device_list_strings and self.device_list_strings[0] not in ["Error querying devices", "No input devices found"]:
            self.audio_device_combobox.set(self.device_list_strings[0])

        self.audio_device_combobox.pack(**entry_options)

        # --- Export Folder ---
        export_outer_frame = tk.Frame(self.config_window)
        export_outer_frame.pack(**frame_options)
        tk.Label(export_outer_frame, text="Export Folder (Audio/Text Files):").pack(**label_options)
        export_browse_frame = tk.Frame(export_outer_frame)
        export_browse_frame.pack(fill=tk.X, padx=pad_options['padx'])
        self.export_folder_entry = ttk.Entry(export_browse_frame)
        self.export_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.export_folder_entry.insert(0, getattr(self, 'export_folder', '.'))
        ttk.Button(export_browse_frame, text="Browse...", command=self.browse_export_folder).pack(side=tk.LEFT, padx=(5,0))

        # --- VAD Settings ---
        vad_frame = ttk.LabelFrame(self.config_window, text="Auto-Pause (VAD) Settings", style="Config.TLabelframe")
        vad_frame.pack(**frame_options)
        vad_inner_frame = tk.Frame(vad_frame) # Use tk.Frame inside if needed
        vad_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(vad_inner_frame, text="Silence Duration (s):").pack(side=tk.LEFT, padx=(0,5))
        self.silence_duration_entry = ttk.Entry(vad_inner_frame, width=5)
        self.silence_duration_entry.insert(0, str(getattr(self, 'silence_threshold_seconds', 5.0)))
        self.silence_duration_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(vad_inner_frame, text="Energy Threshold:").pack(side=tk.LEFT, padx=(0,5))
        self.vad_energy_entry = ttk.Entry(vad_inner_frame, width=7)
        self.vad_energy_entry.insert(0, str(getattr(self, 'vad_energy_threshold', 300)))
        self.vad_energy_entry.pack(side=tk.LEFT)

        # --- Hotkey Settings ---
        hotkey_frame = ttk.LabelFrame(self.config_window, text="Global Hotkeys", style="Config.TLabelframe")
        hotkey_frame.pack(**frame_options)
        hotkey_inner_frame = tk.Frame(hotkey_frame)
        hotkey_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        # Toggle Recording Hotkey
        tk.Label(hotkey_inner_frame, text="Toggle Recording:", width=15, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
        self.hotkey_toggle_entry = ttk.Entry(hotkey_inner_frame, width=30)
        self.hotkey_toggle_entry.grid(row=0, column=1, sticky=tk.EW, padx=2, pady=2)
        self.hotkey_toggle_entry.insert(0, self.hotkey_toggle_record)
        # Show Window Hotkey
        tk.Label(hotkey_inner_frame, text="Show Window:", width=15, anchor=tk.W).grid(row=1, column=0, sticky=tk.W, padx=2, pady=2)
        self.hotkey_show_entry = ttk.Entry(hotkey_inner_frame, width=30)
        self.hotkey_show_entry.grid(row=1, column=1, sticky=tk.EW, padx=2, pady=2)
        self.hotkey_show_entry.insert(0, self.hotkey_show_window)
        # Help Text
        hotkey_help = "Use format like 'ctrl+alt+space', 'win+shift+x'. See 'keyboard' library docs for keys."
        tk.Label(hotkey_inner_frame, text=hotkey_help, justify=tk.LEFT, wraplength=380).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=2, pady=(5,2))
        hotkey_inner_frame.columnconfigure(1, weight=1) # Allow entry to expand

        # --- Logging Settings ---
        logging_frame = ttk.LabelFrame(self.config_window, text="Logging", style="Config.TLabelframe")
        logging_frame.pack(**frame_options)
        log_inner_frame = tk.Frame(logging_frame)
        log_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        ttk.Label(log_inner_frame, text="Level:").pack(side=tk.LEFT, padx=(0,5))
        self.logging_level_combobox = ttk.Combobox(log_inner_frame, values=["None", "Essential", "Extended", "Everything"], state="readonly", width=10)
        self.logging_level_combobox.set(self.logging_level)
        self.logging_level_combobox.pack(side=tk.LEFT, padx=(0, 10))
        self.log_to_file_var = tk.BooleanVar(value=self.log_to_file)
        ttk.Checkbutton(log_inner_frame, text="Log to File", variable=self.log_to_file_var).pack(side=tk.LEFT, padx=(10, 0))

        # --- Transcription Settings ---
        transcription_frame = ttk.LabelFrame(self.config_window, text="Transcription Output", style="Config.TLabelframe")
        transcription_frame.pack(**frame_options)
        trans_inner_frame = tk.Frame(transcription_frame)
        trans_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)

        # <<< START NEW CHECKBOX >>>
        self.beep_on_save_audio_segment_var = tk.BooleanVar(value=self.beep_on_save_audio_segment)
        ttk.Checkbutton(trans_inner_frame, text="Beep on Audio Segment Save (when .wav is created)", variable=self.beep_on_save_audio_segment_var).pack(anchor=tk.W)
        # <<< END NEW CHECKBOX >>>

        self.beep_on_transcription_var = tk.BooleanVar(value=self.beep_on_transcription)
        ttk.Checkbutton(trans_inner_frame, text="Beep on Transcription Completion", variable=self.beep_on_transcription_var).pack(anchor=tk.W)
        # Auto-Paste frame
        auto_paste_frame = tk.Frame(trans_inner_frame)
        auto_paste_frame.pack(anchor=tk.W, fill=tk.X, pady=(5,0))
        self.auto_paste_var = tk.BooleanVar(value=self.auto_paste)
        ttk.Checkbutton(auto_paste_frame, text="Auto-Paste After Transcription", variable=self.auto_paste_var).pack(side=tk.LEFT)
        delay_frame = tk.Frame(auto_paste_frame)
        delay_frame.pack(side=tk.LEFT, padx=(10, 0))
        ttk.Label(delay_frame, text="Delay (s):").pack(side=tk.LEFT)
        self.auto_paste_delay_entry = ttk.Entry(delay_frame, width=5)
        self.auto_paste_delay_entry.insert(0, str(self.auto_paste_delay))
        self.auto_paste_delay_entry.pack(side=tk.LEFT)

        # --- File Versioning ---
        versioning_outer_frame = ttk.LabelFrame(self.config_window, text="Settings/Prompt File Versioning", style="Config.TLabelframe")
        versioning_outer_frame.pack(**frame_options)
        ver_inner_frame = tk.Frame(versioning_outer_frame)
        ver_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        self.versioning_var_config = tk.BooleanVar(value=getattr(self, 'versioning_enabled', True))
        ttk.Checkbutton(ver_inner_frame, text="Enable Backups for Config/Prompt/Commands", variable=self.versioning_var_config).pack(anchor=tk.W)
        versioning_details_frame = tk.Frame(ver_inner_frame)
        versioning_details_frame.pack(fill=tk.X, pady=(5,0))
        tk.Label(versioning_details_frame, text="Backup Folder:").pack(side=tk.LEFT, padx=(0, 5))
        self.backup_folder_entry = ttk.Entry(versioning_details_frame, width=15)
        self.backup_folder_entry.insert(0, getattr(self, 'backup_folder', 'OldVersions'))
        self.backup_folder_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        tk.Label(versioning_details_frame, text="Max Backups:").pack(side=tk.LEFT, padx=(0, 5))
        self.max_backups_entry = ttk.Entry(versioning_details_frame, width=4)
        self.max_backups_entry.insert(0, str(getattr(self, 'max_backups', 10)))
        self.max_backups_entry.pack(side=tk.LEFT)

        # --- Cleanup Settings ---
        cleanup_frame = ttk.LabelFrame(self.config_window, text="Session File Cleanup", style="Config.TLabelframe")
        cleanup_frame.pack(**frame_options)
        clean_inner_frame = tk.Frame(cleanup_frame)
        clean_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        self.clear_audio_var = tk.BooleanVar(value=getattr(self, 'clear_audio_on_exit', False))
        self.clear_text_var = tk.BooleanVar(value=getattr(self, 'clear_text_on_exit', False))
        ttk.Checkbutton(clean_inner_frame, text="Clear Audio (.wav/.transcribed) on Exit", variable=self.clear_audio_var).pack(anchor=tk.W)
        ttk.Checkbutton(clean_inner_frame, text="Clear Text (.txt) on Exit", variable=self.clear_text_var).pack(anchor=tk.W)
        ttk.Button(clean_inner_frame, text="Delete Session Files Now", command=self.delete_session_files).pack(pady=(5,0), anchor=tk.W)

        # --- Close Button Behavior ---
        close_behavior_frame = ttk.LabelFrame(self.config_window, text="Window Close (X) Button", style="Config.TLabelframe")
        close_behavior_frame.pack(**frame_options)
        close_inner_frame = tk.Frame(close_behavior_frame)
        close_inner_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)
        tk.Label(close_inner_frame, text="Action:").pack(side=tk.LEFT, padx=(0,5))
        self.close_behavior_combobox = ttk.Combobox(close_inner_frame, values=["Minimize to tray", "Exit app"], state="readonly", width=20)
        self.close_behavior_combobox.set(self.close_behavior)
        self.close_behavior_combobox.pack(anchor=tk.W)

        # --- Buttons Frame ---
        button_frame = tk.Frame(self.config_window)
        # Increased pady for more space before buttons
        button_frame.pack(pady=(15, 10), padx=pad_options['padx'], fill=tk.X)
        ttk.Button(button_frame, text="Save Configuration", command=self.save_configuration).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Configure Commands...", command=self.open_command_configuration_window).pack(side=tk.LEFT)


    def browse_whisper_executable(self):
        # ... (remains the same) ...
        filepath = filedialog.askopenfilename(title="Select Whisper Executable", filetypes=[("Executable files", "*.exe"), ("All files", "*.*")])
        if filepath:
            norm_path = os.path.normpath(filepath)
            self.whisper_executable_entry.delete(0, tk.END)
            self.whisper_executable_entry.insert(0, norm_path)
            self.log_message("extended", f"Whisper executable selected: {norm_path}")

    def browse_export_folder(self):
        # ... (remains the same) ...
        directory = filedialog.askdirectory(title="Select Export Folder", initialdir=getattr(self, 'export_folder', '.'))
        if directory:
            norm_dir = os.path.normpath(directory)
            self.export_folder_entry.delete(0, tk.END)
            self.export_folder_entry.insert(0, norm_dir)
            self.log_message("extended", f"Export folder selected: {norm_dir}")


    def save_configuration(self):
        self.log_message("essential", "Saving configuration from dialog...")
        config_changed = False
        hotkeys_valid = True

        # --- Read values and check for changes ---
        new_whisper_exec = self.whisper_executable_entry.get().strip()
        if new_whisper_exec != self.whisper_executable:
            self.whisper_executable = new_whisper_exec
            config_changed = True

        selected_device_string = self.audio_device_combobox.get()
        new_device_index = self.device_index_map.get(selected_device_string, None)
        if new_device_index != self.selected_audio_device_index:
             if new_device_index is None and selected_device_string not in ["Error querying devices", "No input devices found"]:
                 self.log_message("essential", f"Warning: Selected audio device string '{selected_device_string}' not found in map. Resetting.")
             self.selected_audio_device_index = new_device_index
             config_changed = True
             self.update_audio_device() # Apply immediately

        new_export_folder = self.export_folder_entry.get().strip() or "."
        if new_export_folder != self.export_folder:
            self.export_folder = new_export_folder
            config_changed = True

        # VAD Settings (with validation)
        try:
            new_silence = float(self.silence_duration_entry.get())
            if new_silence >= 0 and new_silence != self.silence_threshold_seconds:
                self.silence_threshold_seconds = new_silence
                config_changed = True
            elif new_silence < 0: raise ValueError("Negative value")
        except ValueError:
            self.log_message("essential", "Invalid silence duration input.")
            self.silence_duration_entry.delete(0, tk.END); self.silence_duration_entry.insert(0, str(self.silence_threshold_seconds))
        try:
            new_vad_energy = int(self.vad_energy_entry.get())
            if new_vad_energy >= 0 and new_vad_energy != self.vad_energy_threshold:
                self.vad_energy_threshold = new_vad_energy
                config_changed = True
            elif new_vad_energy < 0: raise ValueError("Negative value")
        except ValueError:
            self.log_message("essential", "Invalid VAD energy threshold input.")
            self.vad_energy_entry.delete(0, tk.END); self.vad_energy_entry.insert(0, str(self.vad_energy_threshold))

        # Hotkeys
        new_hotkey_toggle = self.hotkey_toggle_entry.get().strip().lower()
        new_hotkey_show = self.hotkey_show_entry.get().strip().lower()
        hotkeys_changed = (new_hotkey_toggle != self.hotkey_toggle_record or
                           new_hotkey_show != self.hotkey_show_window)
        if hotkeys_changed:
            self.hotkey_toggle_record = new_hotkey_toggle
            self.hotkey_show_window = new_hotkey_show
            # Attempt to update hotkeys immediately and check validity
            hotkeys_valid = self.update_hotkeys()
            config_changed = True

        # Logging
        new_log_level = self.logging_level_combobox.get()
        if new_log_level != self.logging_level:
            self.logging_level = new_log_level
            config_changed = True
        new_log_to_file = self.log_to_file_var.get()
        if new_log_to_file != self.log_to_file:
            self.log_to_file = new_log_to_file
            config_changed = True # May need to open/close file

        # Transcription
        new_beep_on_save = self.beep_on_save_audio_segment_var.get() # <<< NEW
        if new_beep_on_save != self.beep_on_save_audio_segment:      # <<< NEW
            self.beep_on_save_audio_segment = new_beep_on_save       # <<< NEW
            config_changed = True                                     # <<< NEW

        new_beep = self.beep_on_transcription_var.get()
        if new_beep != self.beep_on_transcription: self.beep_on_transcription = new_beep; config_changed = True
        new_auto_paste = self.auto_paste_var.get()
        if new_auto_paste != self.auto_paste: self.auto_paste = new_auto_paste; config_changed = True
        try:
            new_paste_delay = float(self.auto_paste_delay_entry.get())
            if new_paste_delay >= 0 and new_paste_delay != self.auto_paste_delay:
                 self.auto_paste_delay = new_paste_delay; config_changed = True
            elif new_paste_delay < 0: raise ValueError("Negative value")
        except ValueError:
            self.log_message("essential", "Invalid auto-paste delay.")
            self.auto_paste_delay_entry.delete(0, tk.END); self.auto_paste_delay_entry.insert(0, str(self.auto_paste_delay))

        # Versioning
        new_versioning = self.versioning_var_config.get()
        if new_versioning != self.versioning_enabled: self.versioning_enabled = new_versioning; config_changed = True
        new_backup_folder = self.backup_folder_entry.get().strip() or "OldVersions"
        if new_backup_folder != self.backup_folder: self.backup_folder = new_backup_folder; config_changed = True
        try:
             new_max_backups = int(self.max_backups_entry.get())
             if new_max_backups >= 0 and new_max_backups != self.max_backups:
                  self.max_backups = new_max_backups; config_changed = True
             elif new_max_backups < 0: raise ValueError("Negative value")
        except ValueError:
             self.log_message("essential", "Invalid max backups.")
             self.max_backups_entry.delete(0, tk.END); self.max_backups_entry.insert(0, str(self.max_backups))

        # Cleanup
        new_clear_audio = self.clear_audio_var.get()
        if new_clear_audio != self.clear_audio_on_exit: self.clear_audio_on_exit = new_clear_audio; config_changed = True
        new_clear_text = self.clear_text_var.get()
        if new_clear_text != self.clear_text_on_exit: self.clear_text_on_exit = new_clear_text; config_changed = True

        # Close Behavior
        new_close_behavior = self.close_behavior_combobox.get()
        if new_close_behavior != self.close_behavior:
             self.close_behavior = new_close_behavior
             config_changed = True
             self.root.protocol("WM_DELETE_WINDOW", self.handle_close_button) # Re-bind protocol


        # --- Save and Close ---
        if not hotkeys_valid:
            messagebox.showerror("Invalid Hotkey", "One or more hotkeys have invalid syntax. Please correct them and save again.", parent=self.config_window)
            return # Do not save or close

        if config_changed:
            self.save_settings()
            self.log_message("essential", "Configuration saved successfully.")
        else:
            self.log_message("extended", "No configuration changes detected.")

        if self.config_window and self.config_window.winfo_exists():
            self.config_window.destroy()
            self.config_window = None


    # --- Command Configuration Window ---
    def open_command_configuration_window(self):
        # ... (remains the same) ...
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.lift()
            self.command_config_window.focus_force()
            return
        self.command_config_window = tk.Toplevel(self.root)
        self.command_config_window.title("Configure Commands")
        self.command_config_window.geometry("700x450")
        self.command_config_window.transient(self.root)
        self.command_config_window.grab_set()
        self.create_command_config_widgets()

    def create_command_config_widgets(self):
        # ... (remains the same) ...
        instructions = ("Define voice commands and actions.\n"
                        "Use ' FF ' (with spaces) as a wildcard for captured text.\n"
                        "Example Voice: 'open file FF ' | Action: 'explorer FF '")
        ttk.Label(self.command_config_window, text=instructions, justify=tk.LEFT).pack(pady=(5, 10), padx=10, anchor=tk.W)

        list_container = tk.Frame(self.command_config_window)
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        canvas = tk.Canvas(list_container)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.scrollable_command_frame = tk.Frame(canvas)
        self.scrollable_command_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_command_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Mouse wheel binding (simplified)
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))


        self.command_widgets = []
        if not self.loaded_commands:
            self.add_command_row()
        else:
            for cmd in self.loaded_commands:
                self.add_command_row(cmd.get("voice", ""), cmd.get("action", ""))

        bottom_button_frame = tk.Frame(self.command_config_window)
        bottom_button_frame.pack(pady=10)
        ttk.Button(bottom_button_frame, text="Add Command Row", command=self.add_command_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_button_frame, text="Save and Close", command=self.save_commands_from_ui).pack(side=tk.LEFT, padx=5)


    def add_command_row(self, voice_cmd="", action_cmd=""):
        # ... (remains the same) ...
        command_frame = tk.Frame(self.scrollable_command_frame)
        command_frame.pack(pady=2, padx=5, fill=tk.X, expand=True)
        ttk.Label(command_frame, text="Voice Cmd:", width=10).pack(side=tk.LEFT, padx=(0, 2))
        voice_entry = ttk.Entry(command_frame, width=30)
        voice_entry.insert(0, voice_cmd)
        voice_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Label(command_frame, text="Action:", width=8).pack(side=tk.LEFT, padx=(5, 2))
        action_entry = ttk.Entry(command_frame, width=30)
        action_entry.insert(0, action_cmd)
        action_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        remove_button = ttk.Button(command_frame, text="X", width=3, style="Toolbutton",
                                  command=lambda f=command_frame: self.remove_command_row(f))
        remove_button.pack(side=tk.LEFT, padx=(5, 0))
        self.command_widgets.append({"frame": command_frame, "voice": voice_entry, "action": action_entry})
        self.scrollable_command_frame.update_idletasks()
        canvas = self.scrollable_command_frame.master
        canvas.configure(scrollregion=canvas.bbox("all"))


    def remove_command_row(self, command_frame):
        # ... (remains the same) ...
        widget_ref_to_remove = None
        for i, ref in enumerate(self.command_widgets):
            if ref["frame"] == command_frame:
                widget_ref_to_remove = ref
                break
        if widget_ref_to_remove:
            command_frame.destroy()
            self.command_widgets.remove(widget_ref_to_remove)
            self.scrollable_command_frame.update_idletasks()
            canvas = self.scrollable_command_frame.master
            canvas.configure(scrollregion=canvas.bbox("all"))

    def save_commands_from_ui(self):
        # ... (remains the same) ...
        current_commands = []
        for widget_ref in self.command_widgets:
            voice = widget_ref["voice"].get().strip()
            action = widget_ref["action"].get().strip()
            if voice and action:
                current_commands.append({"voice": voice, "action": action})
        self.loaded_commands = current_commands
        self.commands = list(self.loaded_commands)
        self.save_commands()
        self.log_message("essential", f"Commands saved from UI: {len(self.loaded_commands)} commands.")
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.destroy()
            self.command_config_window = None

    # --- Prompt Import/Export ---
    def import_prompt(self):
        # ... (remains the same) ...
        filename = filedialog.askopenfilename(initialdir=".", title="Select Prompt File", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")))
        if filename:
            try:
                with open(filename, "r", encoding='utf-8') as f: content = f.read()
                if hasattr(self, 'prompt_text') and self.prompt_text.winfo_exists():
                    self.prompt_text.delete("1.0", tk.END); self.prompt_text.insert("1.0", content)
                self.prompt = content; self.save_prompt()
                self.log_message("essential", f"Prompt imported from {filename}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Error importing prompt:\n{e}")
                self.log_message("essential", f"Error importing prompt: {e}")

    def export_prompt(self):
        # ... (remains the same) ...
        filename = filedialog.asksaveasfilename(initialdir=".", title="Save Prompt As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                content_to_export = self.prompt_text.get("1.0", tk.END) if hasattr(self, 'prompt_text') and self.prompt_text.winfo_exists() else self.prompt
                with open(filename, "w", encoding='utf-8') as f: f.write(content_to_export)
                self.log_message("essential", f"Prompt exported to {filename}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Error exporting prompt:\n{e}")
                self.log_message("essential", f"Error exporting prompt: {e}")

    # --- Scratchpad ---
    def open_scratchpad_window(self):
        # ... (remains the same) ...
        if self.scratchpad_window and self.scratchpad_window.winfo_exists():
            self.scratchpad_window.lift(); self.scratchpad_window.focus_force(); return

        self.scratchpad_window = tk.Toplevel(self.root)
        self.scratchpad_window.title("WhisperR Scratchpad")
        self.scratchpad_window.geometry("500x700")

        text_frame = tk.Frame(self.scratchpad_window); text_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.scratchpad_text_widget = tk.Text(text_frame, wrap=tk.WORD, undo=True)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.scratchpad_text_widget.yview)
        self.scratchpad_text_widget.configure(yscrollcommand=scrollbar.set)
        self.scratchpad_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        controls_frame = tk.Frame(self.scratchpad_window); controls_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(controls_frame, text="Import...", command=self.import_to_scratchpad).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(controls_frame, text="Export As...", command=self.export_from_scratchpad).pack(side=tk.LEFT, padx=(2, 2))
        ttk.Button(controls_frame, text="Clear", command=self.clear_scratchpad).pack(side=tk.LEFT, padx=(2, 0))
        append_frame = tk.Frame(controls_frame); append_frame.pack(side=tk.RIGHT)
        self.scratchpad_append_var = tk.BooleanVar(value=self.scratchpad_append_mode)
        ttk.Checkbutton(append_frame, text="Append Mode", variable=self.scratchpad_append_var, command=self.toggle_scratchpad_append).pack()


    def clear_scratchpad(self):
        # ... (remains the same) ...
        if self.scratchpad_text_widget and self.scratchpad_text_widget.winfo_exists():
             if messagebox.askyesno("Clear Scratchpad", "Clear scratchpad content?", parent=self.scratchpad_window):
                  self.scratchpad_text_widget.delete("1.0", tk.END)
                  self.log_message("essential", "Scratchpad cleared.")


    def toggle_scratchpad_append(self):
        # ... (remains the same) ...
        if hasattr(self, 'scratchpad_append_var'):
             self.scratchpad_append_mode = self.scratchpad_append_var.get()
             self.log_message("essential", f"Scratchpad Append Mode set to: {self.scratchpad_append_mode}")


    def update_scratchpad_text(self, text_to_add):
        # ... (remains the same) ...
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget):
            return
        try:
            if self.scratchpad_append_mode:
                current_content = self.scratchpad_text_widget.get("1.0", tk.END).strip()
                separator = "\n\n" if current_content else ""
                self.scratchpad_text_widget.insert(tk.END, separator + text_to_add)
                self.log_message("extended", f"Appended text to scratchpad.")
                self.scratchpad_text_widget.see(tk.END)
            else:
                self.scratchpad_text_widget.delete("1.0", tk.END)
                self.scratchpad_text_widget.insert("1.0", text_to_add)
                self.log_message("extended", f"Replaced text in scratchpad.")
                self.scratchpad_text_widget.see("1.0")
        except tk.TclError as e: self.log_message("essential", f"Error updating scratchpad (TclError): {e}")
        except Exception as e: self.log_message("essential", f"Unexpected error updating scratchpad: {e}")


    def import_to_scratchpad(self):
        # ... (remains the same) ...
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget): return
        filename = filedialog.askopenfilename(initialdir=".", title="Import to Scratchpad", filetypes=(("Text files", "*.txt*"),("All files", "*.*")))
        if filename:
            try:
                with open(filename, "r", encoding='utf-8') as f: content = f.read()
                if self.scratchpad_append_mode:
                     separator = "\n\n" if self.scratchpad_text_widget.get("1.0", tk.END).strip() else ""
                     self.scratchpad_text_widget.insert(tk.END, separator + content); self.scratchpad_text_widget.see(tk.END)
                else:
                     self.scratchpad_text_widget.delete("1.0", tk.END); self.scratchpad_text_widget.insert("1.0", content); self.scratchpad_text_widget.see("1.0")
                self.log_message("essential", f"Imported '{filename}' to scratchpad.")
            except Exception as e:
                messagebox.showerror("Import Error", f"Error importing to scratchpad:\n{e}", parent=self.scratchpad_window)
                self.log_message("essential", f"Error importing to scratchpad: {e}")

    def export_from_scratchpad(self):
        # ... (remains the same) ...
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget): return
        filename = filedialog.asksaveasfilename(initialdir=".", title="Export Scratchpad As", filetypes=(("Text files", "*.txt"),("All files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                content = self.scratchpad_text_widget.get("1.0", tk.END)
                if content.endswith('\n'): content = content[:-1]
                with open(filename, "w", encoding='utf-8') as f: f.write(content)
                self.log_message("essential", f"Exported scratchpad to '{filename}'.")
            except Exception as e:
                messagebox.showerror("Export Error", f"Error exporting from scratchpad:\n{e}", parent=self.scratchpad_window)
                self.log_message("essential", f"Error exporting from scratchpad: {e}")


    # --- Queue Management ---
    def update_queue_indicator(self):
        # ... (remains the same) ...
        if hasattr(self, 'queue_indicator_var'):
            try: qsize = self.transcription_queue.qsize(); self.queue_indicator_var.set(f"Queue: {qsize}")
            except Exception as e: self.log_message("essential", f"Error updating queue indicator: {e}")


    def toggle_queue_processing(self):
        # ... (remains the same) ...
        self.queue_processing_paused = not self.queue_processing_paused
        new_state = "Paused" if self.queue_processing_paused else "Running"
        button_text = "Resume Queue" if self.queue_processing_paused else "Pause Queue"
        self.log_message("essential", f"Queue processing {new_state}.")
        if hasattr(self, 'pause_queue_button') and self.pause_queue_button.winfo_exists(): self.pause_queue_button.config(text=button_text)
        if hasattr(self, 'pause_queue_menu_var'): self.pause_queue_menu_var.set(self.queue_processing_paused)


    def clear_transcription_queue(self):
        # ... (remains the same) ...
        if messagebox.askyesno("Clear Queue", "Remove all pending items from the transcription queue?"):
            self.log_message("essential", "Clearing transcription queue...")
            self.clear_queue_flag = True
            cleared_count = 0
            try:
                 while not self.transcription_queue.empty():
                     try:
                         item = self.transcription_queue.get_nowait()
                         if item is not AUDIO_QUEUE_SENTINEL:
                             cleared_count += 1; self.log_message("extended", f"Removing item from queue: {item}")
                             # TODO: Option to delete files? Risky.
                         else: self.transcription_queue.put(AUDIO_QUEUE_SENTINEL) # Put back sentinel
                         self.transcription_queue.task_done()
                     except queue.Empty: break
                     except Exception as e_get: self.log_message("essential", f"Error getting item during queue clear: {e_get}"); break
            finally:
                 self.clear_queue_flag = False
                 self.log_message("essential", f"Queue clear signal finished. Approx {cleared_count} items removed/discarded.")
                 self.root.after(0, self.update_queue_indicator)


    # --- File Cleanup ---
    def delete_session_files(self, ask_confirm=True):
        # ... (remains the same) ...
        export_dir = getattr(self, 'export_folder', '.')
        norm_export_dir = os.path.normpath(export_dir)

        if not os.path.isdir(norm_export_dir):
            self.log_message("essential", f"Cleanup skipped: Export directory not found: {norm_export_dir}")
            if ask_confirm: messagebox.showwarning("Cleanup Warning", f"Export directory not found:\n{norm_export_dir}")
            return

        delete_audio = False; delete_text = False
        if ask_confirm:
             # Check config window state for source of truth
             if self.config_window and self.config_window.winfo_exists() and hasattr(self, 'clear_audio_var'):
                 delete_audio = self.clear_audio_var.get()
                 delete_text = self.clear_text_var.get()
             else: # Read from instance vars if config not open
                 delete_audio = getattr(self, 'clear_audio_on_exit', False)
                 delete_text = getattr(self, 'clear_text_on_exit', False)
        else: # Called on exit
            delete_audio = getattr(self, 'clear_audio_on_exit', False)
            delete_text = getattr(self, 'clear_text_on_exit', False)

        if not (delete_audio or delete_text):
            if ask_confirm: messagebox.showinfo("Cleanup Info", "Please enable 'Clear Audio' or 'Clear Text' in Configuration to select files for deletion.")
            return

        self.log_message("extended", f"Scanning '{norm_export_dir}' for session files (Audio: {delete_audio}, Text: {delete_text})...")
        files_to_delete = []
        try:
             for filename in os.listdir(norm_export_dir):
                 filepath = os.path.join(norm_export_dir, filename)
                 if not os.path.isfile(filepath): continue
                 if delete_audio and filename.startswith("recording_") and (filename.endswith(".wav") or filename.endswith(".wav.transcribed")): files_to_delete.append(filepath)
                 elif delete_text and filename.startswith("recording_") and filename.endswith(".txt"): files_to_delete.append(filepath)
        except Exception as e_list:
              self.log_message("essential", f"Error listing files in export directory '{norm_export_dir}': {e_list}")
              if ask_confirm: messagebox.showerror("Cleanup Error", f"Error reading export directory:\n{e_list}")
              return

        if not files_to_delete:
            if ask_confirm: messagebox.showinfo("Cleanup Info", f"No matching session files found in '{norm_export_dir}'.")
            return

        confirm = True
        if ask_confirm:
            types = [t for t, flag in [("Audio", delete_audio), ("Text", delete_text)] if flag]
            confirm_msg = (f"Permanently delete {len(files_to_delete)} session file(s)\n"
                           f"({', '.join(types)} types) from:\n'{norm_export_dir}'?\n\nThis cannot be undone.")
            confirm = messagebox.askyesno("Confirm Deletion", confirm_msg, icon=messagebox.WARNING)

        if confirm:
            deleted_count, errors = 0, 0
            self.log_message("essential", f"Starting deletion of {len(files_to_delete)} files...")
            for filepath in files_to_delete:
                try: os.remove(filepath); deleted_count += 1; self.log_message("extended", f"Deleted: {filepath}")
                except Exception as e: errors += 1; self.log_message("essential", f"Error deleting {filepath}: {e}")
            result_msg = f"Deleted {deleted_count} file(s)." + (f" Failed to delete {errors}." if errors else "")
            self.log_message("essential", f"Cleanup finished. {result_msg}")
            if ask_confirm: messagebox.showinfo("Cleanup Result", result_msg)
        else: self.log_message("essential", "File deletion cancelled.")


# --- Main Execution ---
if __name__ == "__main__":
    # --- DPI Awareness (Windows) ---
    try: from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e: print(f"Note: Could not set DPI awareness ({e}).")

    root = tk.Tk()

    # --- Apply Theme ---
    try:
        style = ttk.Style(root)
        available_themes = style.theme_names()
        print(f"Available ttk themes: {available_themes}")
        # Prefer modern themes if available
        preferred_themes = ['clam', 'vista', 'xpnative', 'winnative', 'default']
        for theme in preferred_themes:
            if theme in available_themes:
                try: style.theme_use(theme); print(f"Using theme: {theme}"); break
                except tk.TclError: print(f"Failed to apply theme: {theme}")
    except Exception as e_style: print(f"Could not apply ttk theme: {e_style}")

    # --- Instantiate and Run ---
    app = WhisperRApp(root)

    # Start tray icon thread
    tray_thread = threading.Thread(target=app.setup_tray_icon_thread, daemon=True)
    tray_thread.start()

    app.log_message("essential", f"WhisperR application started.")
    app.log_message("essential", "Starting Tkinter mainloop...")
    try:
        root.mainloop()
    except KeyboardInterrupt:
        app.log_message("essential", "KeyboardInterrupt received, initiating shutdown...")
        app.quit_app_action()
    except Exception as main_loop_e:
         # Log fatal errors during mainloop
         app.log_message("essential", f"FATAL ERROR in mainloop: {main_loop_e}")
         import traceback
         app.log_message("essential", traceback.format_exc())
         try:
             # Attempt graceful shutdown
             app.quit_app_action()
         except Exception as quit_e:
             app.log_message("essential", f"Error during emergency shutdown: {quit_e}")
             os._exit(1) # Force exit if shutdown fails

    app.log_message("essential", "Tkinter mainloop finished.")
    app.log_message("essential", "WhisperR application exiting.")
    sys.exit(0)