import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
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
from collections import deque
import queue
import winsound

CONFIG_FILE = "config.json"
PROMPT_FILE = "prompt.json"
COMMANDS_FILE = "commands.json"
AUDIO_QUEUE_SENTINEL = None

class VoiceControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voice Control App")
        self.root.minsize(600, 620)

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
        self.auto_paste = False
        self.auto_paste_delay = 1.0
        self.backup_folder = "OldVersions" # New setting for backup folder
        self.max_backups = 10 # New setting for max backups to keep

        # --- Runtime State ---
        self.recording = False
        self.audio_buffer = deque()
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

        self.load_settings()
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app_action)
        self.update_audio_device()
        self.start_transcription_worker()

    def log_message(self, level, message):
        level_map = {"None": 0, "Essential": 1, "Extended": 2, "Everything": 3}
        current_level = level_map.get(self.logging_level, 3)
        msg_level_map = {"essential": 1, "extended": 2, "everything": 3}
        msg_level = msg_level_map.get(level, 3)

        if current_level == 0 or msg_level > current_level:
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] {message}"
        
        # Console output
        print(log_line)

        # File output
        if self.log_to_file:
            if not self.log_file:
                log_filename = f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                try:
                    self.log_file = open(log_filename, 'a', encoding='utf-8')
                    self.log_message("essential", f"Logging to file: {log_filename}")
                except Exception as e:
                    self.log_message("essential", f"Error opening log file: {e}")
            if self.log_file:
                try:
                    self.log_file.write(log_line + "\n")
                    self.log_file.flush()
                except Exception as e:
                    self.log_message("essential", f"Error writing to log file: {e}")

    def load_settings(self):
        # Load config.json
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
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
                    self.auto_paste = settings.get('auto_paste', False)
                    self.auto_paste_delay = float(settings.get('auto_paste_delay', 1.0))
                    self.backup_folder = settings.get('backup_folder', 'OldVersions') # Load new setting
                    self.max_backups = int(settings.get('max_backups', 10)) # Load new setting
                self.log_message("essential", f"Core settings loaded from {CONFIG_FILE}")
            else:
                self.log_message("essential", f"{CONFIG_FILE} not found, using core defaults.")
        except Exception as e:
            self.log_message("essential", f"Error loading {CONFIG_FILE}: {e}. Using core defaults.")

        # Load prompt.json
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

        # Load commands.json
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
            'auto_paste': getattr(self, 'auto_paste', False),
            'auto_paste_delay': getattr(self, 'auto_paste_delay', 1.0),
            'backup_folder': getattr(self, 'backup_folder', 'OldVersions'), # Save new setting
            'max_backups': getattr(self, 'max_backups', 10) # Save new setting
        }
        try:
            self.create_backup(CONFIG_FILE)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            self.log_message("essential", f"Core settings saved to {CONFIG_FILE}")
        except Exception as e:
            self.log_message("essential", f"Error saving {CONFIG_FILE}: {e}")

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

    def create_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, padx=10, fill=tk.X)
        lang_frame = tk.Frame(top_frame)
        lang_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(lang_frame, text="Language:").pack(side=tk.LEFT, padx=(0, 5))
        self.language_options = ["en", "es", "fr", "de", "it", "ja", "zh", "el"]
        self.language_combobox = ttk.Combobox(lang_frame, values=self.language_options, state="readonly", width=10)
        self.language_combobox.set(self.language)
        self.language_combobox.pack(side=tk.LEFT)
        self.language_combobox.bind("<<ComboboxSelected>>", lambda event: self.update_language())

        model_frame = tk.Frame(top_frame)
        model_frame.pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Label(model_frame, text="Model:").pack(side=tk.LEFT, padx=(0, 5))
        self.model_options = ["tiny", "base", "small", "medium", "large"]
        self.model_combobox = ttk.Combobox(model_frame, values=self.model_options, state="readonly", width=10)
        self.model_combobox.set(self.model)
        self.model_combobox.pack(side=tk.LEFT)
        self.model_combobox.bind("<<ComboboxSelected>>", lambda event: self.update_model())

        toggle_frame1 = tk.Frame(self.root)
        toggle_frame1.pack(pady=5, padx=10, fill=tk.X)
        trans_frame = tk.Frame(toggle_frame1)
        trans_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.translation_var = tk.BooleanVar(value=self.translation_enabled)
        self.translation_var.trace_add("write", self.update_translation)
        ttk.Checkbutton(trans_frame, text="Enable Translation", variable=self.translation_var).pack(side=tk.LEFT)
        cmd_frame = tk.Frame(toggle_frame1)
        cmd_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.command_mode_var = tk.BooleanVar(value=self.command_mode)
        self.command_mode_var.trace_add("write", self.update_command_mode)
        ttk.Checkbutton(cmd_frame, text="Enable Auto-Pause / Commands", variable=self.command_mode_var).pack(side=tk.LEFT)

        toggle_frame2 = tk.Frame(self.root)
        toggle_frame2.pack(pady=2, padx=10, fill=tk.X)
        ts_frame = tk.Frame(toggle_frame2)
        ts_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.timestamps_disabled_var = tk.BooleanVar(value=self.timestamps_disabled)
        self.timestamps_disabled_var.trace_add("write", self.update_timestamps_disabled)
        ttk.Checkbutton(ts_frame, text="Disable Timestamps", variable=self.timestamps_disabled_var).pack(side=tk.LEFT)
        clear_text_frame = tk.Frame(toggle_frame2)
        clear_text_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.clear_text_output_var = tk.BooleanVar(value=self.clear_text_output)
        self.clear_text_output_var.trace_add("write", self.update_clear_text_output)
        ttk.Checkbutton(clear_text_frame, text="Clear Text Output", variable=self.clear_text_output_var).pack(side=tk.LEFT)

        ttk.Label(self.root, text="Prompt:").pack(pady=(10, 0), padx=10, anchor=tk.W)
        self.prompt_text = tk.Text(self.root, height=8)
        self.prompt_text.insert("1.0", self.prompt)
        self.prompt_text.pack(pady=5, padx=10, fill=tk.X)
        self.prompt_text.bind("<KeyRelease>", self.update_prompt)
        self.prompt_text.bind("<FocusOut>", self.update_prompt)

        self.import_export_frame = tk.Frame(self.root)
        self.import_export_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.import_export_frame, text="Import Prompt", command=self.import_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(self.import_export_frame, text="Export Prompt", command=self.export_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        ok_scratch_frame = tk.Frame(self.root)
        ok_scratch_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(ok_scratch_frame, text="OK", command=self.store_settings_and_hide).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(ok_scratch_frame, text="Scratchpad", command=self.open_scratchpad_window).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        # Start/Stop Recording button and indicator
        start_stop_frame = tk.Frame(self.root)
        start_stop_frame.pack(pady=(20, 0), padx=10, fill=tk.X)
        button_inner_frame = tk.Frame(start_stop_frame)
        button_inner_frame.pack(fill=tk.X)
        self.start_stop_button = ttk.Button(button_inner_frame, text="Start Recording", command=self.toggle_recording)
        self.start_stop_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.recording_indicator_label = ttk.Label(button_inner_frame, text="●", font=("Arial", 12))
        self.recording_indicator_label.pack(side=tk.LEFT, padx=5)
        self.update_recording_indicator()

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, padx=10, fill=tk.X, side=tk.BOTTOM, anchor=tk.S)
        shortcut_text = "CTRL+Alt+Space: Toggle Record\nCTRL+Alt+Shift+Space: Show Window" # Removed Transcode Last shortcut
        ttk.Label(bottom_frame, text=shortcut_text, justify=tk.LEFT).pack(side=tk.LEFT, anchor=tk.W)
        ttk.Button(bottom_frame, text="Configuration", command=self.open_configuration_window).pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))

        try:
            keyboard.add_hotkey('ctrl+alt+space', self.toggle_recording, suppress=False) # Changed suppress to False
            # Removed hotkey for transcode_last_recording
            keyboard.add_hotkey('ctrl+alt+shift+space', self.return_to_window, suppress=False) # Changed suppress to False
            self.log_message("essential", "Hotkeys registered successfully.")
        except Exception as e:
            self.log_message("essential", f"Failed to register hotkeys: {e}")

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
        self.vad_enabled = self.command_mode
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
        new_prompt = self.prompt_text.get("1.0", tk.END).strip()
        if new_prompt != self.prompt:
            self.prompt = new_prompt
            self.log_message("essential", f"Prompt updated (length: {len(self.prompt)} chars)")
            self.save_prompt()

    def setup_tray_icon_thread(self):
        try:
            # Determine the correct path to icon.png
            icon_path = None
            possible_paths = []
            
            # Check in MEIPASS first if bundled
            if hasattr(sys, '_MEIPASS'):
                possible_paths.append(os.path.join(sys._MEIPASS, 'icon.png'))
            
            # Check in current directory
            possible_paths.append('icon.png')
            
            # Check in script directory
            possible_paths.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.png'))
            
            # Find first existing path
            icon_path = None
            try:
                for path in possible_paths:
                    if os.path.exists(path):
                        icon_path = path
                        break
            except Exception as e:
                self.log_message("essential", f"Error checking icon paths: {e}")
            
            
            if not icon_path:
                self.log_message("essential", "Warning: icon.png not found in any standard locations")
                # Create a blank image as fallback
                image = Image.new('RGB', (64, 64), color='gray')
            else:
                image = Image.open(icon_path)
            menu = pystray.Menu(pystray.MenuItem("Show", self.show_window_action), pystray.MenuItem("Quit", self.quit_app_action))
            self.tray_icon = pystray.Icon("Voice Control App", image, "Voice Control App", menu)
            self.log_message("essential", "Running tray icon...")
            self.tray_icon.run()
            self.log_message("essential", "Tray icon thread finished.")
        except FileNotFoundError:
            self.log_message("essential", f"Error: icon.png not found at {icon_path}")
            self.root.after(100, self.quit_app_action)
        except Exception as e:
            self.log_message("essential", f"Error setting up tray icon: {e}")
            self.root.after(100, self.quit_app_action)

    def show_window_action(self):
        self.root.after(0, self._show_window)

    def quit_app_action(self):
        self.log_message("essential", "Quit action initiated...")
        self.recording = False
        if self.audio_stream:
            self.audio_stream.stop()
            self.audio_stream.close()
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout=0.5)
        self.log_message("essential", f"Queue size before shutdown: {self.transcription_queue.qsize()}")
        self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.transcription_worker_thread.join(timeout=2.0)
        if getattr(self, 'clear_audio_on_exit', False) or getattr(self, 'clear_text_on_exit', False):
            self.delete_session_files(ask_confirm=False)
        if self.tray_icon:
            self.log_message("essential", "Stopping tray icon...")
            self.tray_icon.stop()
        self.log_message("essential", "Scheduling root quit...")
        self.save_settings()
        self.save_prompt()
        self.save_commands()
        if self.log_file:
            try:
                self.log_file.close()
                self.log_file = None
            except Exception as e:
                self.log_message("essential", f"Error closing log file: {e}")
        self.root.after(0, self.root.quit)

    def create_green_line(self):
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            hwnd = win32gui.CreateWindowEx(win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED | win32con.WS_EX_TOOLWINDOW,
                                          "Static", None, win32con.WS_VISIBLE | win32con.WS_POPUP, 0, 0, screen_width, 5, 0, 0, 0, None)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
            win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0,0,0), int(0.25 * 255), win32con.LWA_ALPHA)
            hdc = win32gui.GetDC(hwnd)
            brush = win32gui.CreateSolidBrush(win32api.RGB(0, 150, 0))
            rect = win32gui.GetClientRect(hwnd)
            win32gui.FillRect(hdc, rect, brush)
            win32gui.ReleaseDC(hwnd, hdc)
            win32gui.DeleteObject(brush)
            self.green_line = hwnd
            self.log_message("extended", "Green line created.")
        except Exception as e:
            self.log_message("essential", f"Error creating green line: {e}")

    def destroy_green_line(self):
        if self.green_line:
            try:
                win32gui.DestroyWindow(self.green_line)
                self.green_line = None
                self.log_message("extended", "Green line destroyed.")
            except Exception as e:
                self.log_message("essential", f"Error destroying green line: {e}")
                self.green_line = None

    def toggle_recording(self):
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
        self.audio_buffer = deque()
        self.current_segment = []
        self.is_speaking = False
        self.silence_start_time = None
        self.create_green_line()
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join()
        self.recording_thread = threading.Thread(target=self.record_audio_continuously, daemon=True)
        self.recording_thread.start()
        self.update_recording_indicator()

    def stop_recording(self):
        if not self.recording:
            return
        self.log_message("essential", "Manual stop recording requested...")
        self.recording = False
        self.destroy_green_line()
        if self.audio_buffer or self.current_segment:
            self.log_message("extended", "Processing final audio buffer on manual stop...")
            if self.audio_buffer:
                self.current_segment.extend(list(self.audio_buffer))
            self.save_segment_and_reset_vad()
        else:
            self.log_message("extended", "No audio in buffer to save on manual stop.")
            self.current_segment = []
            self.is_speaking = False
            self.silence_start_time = None
        self.audio_buffer = deque()
        self.log_message("essential", f"Recording stopped. Transcription queue size: {self.transcription_queue.qsize()}")
        self.update_recording_indicator()

    def update_recording_indicator(self):
        if hasattr(self, 'recording_indicator_label'):
            color = "red" if self.recording else "gray"
            self.start_stop_button.config(text="Stop Recording" if self.recording else "Start Recording")
            self.recording_indicator_label.config(foreground=color)
            self.log_message("extended", f"Recording indicator updated to {color}.")

    def record_audio_continuously(self):
        self.log_message("essential", "Continuous audio recording thread started.")
        samplerate = 44100
        channels = 1
        dtype = 'int16'
        blocksize = 1024
        device_index = getattr(self, 'selected_audio_device_index', None)
        silence_duration = getattr(self, 'silence_threshold_seconds', 5.0)
        energy_threshold = getattr(self, 'vad_energy_threshold', 300)
        buffer_duration_factor = 1.2
        buffer_max_size = int(samplerate * silence_duration * buffer_duration_factor / blocksize)

        def audio_callback(indata, frames, time_info, status):
            nonlocal self
            if status:
                self.log_message("extended", f"Audio CB Status: {status}")
            if not self.recording:
                return

            current_time_monotonic = time.monotonic()

            self.audio_buffer.append(indata.copy())
            if len(self.audio_buffer) > buffer_max_size:
                self.audio_buffer.popleft()

            rms = np.sqrt(np.mean(indata.astype(np.float32)**2))
            self.log_message("everything", f"RMS: {rms:.2f} (Threshold: {energy_threshold})")

            self.current_segment.append(indata.copy())

            if not self.command_mode:
                self.log_message("extended", "Command mode disabled, continuing to record without VAD.")
                return

            if rms >= energy_threshold:
                if not self.is_speaking:
                    self.log_message("extended", "Speech started.")
                    self.is_speaking = True
                self.silence_start_time = None
            elif self.is_speaking:
                if self.silence_start_time is None:
                    self.silence_start_time = current_time_monotonic
                    self.log_message("extended", "Silence started...")
                if current_time_monotonic - self.silence_start_time >= silence_duration:
                    self.log_message("extended", f"Silence threshold ({silence_duration}s) reached.")
                    if self.recording:
                        self.root.after(0, self.save_segment_and_reset_vad)

        self.audio_stream = None
        try:
            self.log_message("extended", f"Opening InputStream (Continuous) with device index: {device_index}")
            self.audio_stream = sd.InputStream(device=device_index, samplerate=samplerate, channels=channels, dtype=dtype, blocksize=blocksize, callback=audio_callback)
            with self.audio_stream:
                self.log_message("extended", "InputStream opened (Continuous).")
                while self.recording:
                    time.sleep(0.1)
        except sd.PortAudioError as pae:
            self.log_message("essential", f"PortAudioError: {pae}")
            self.root.after(0, self.handle_recording_error)
        except Exception as e:
            self.log_message("essential", f"Error in continuous recording thread: {e}")
            self.root.after(0, self.handle_recording_error)
        finally:
            if self.audio_stream and not self.audio_stream.closed:
                self.log_message("extended", "Closing audio stream.")
                self.audio_stream.close()
            self.log_message("essential", "Continuous audio recording thread finished.")

    def save_segment_and_reset_vad(self):
        if self.current_segment:
            self.save_segment()
            self.current_segment = []
            self.is_speaking = False
            self.silence_start_time = None
        else:
            self.log_message("extended", "Save segment called but no segment data.")

    def save_segment(self):
        if not self.current_segment:
            self.log_message("extended", "No segment data to save.")
            return
        self.log_message("extended", f"Saving segment with {len(self.current_segment)} chunks at {datetime.datetime.now()}")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"recording_{timestamp}.wav"
        # Normalize export_dir right away
        export_dir = os.path.normpath(getattr(self, 'export_folder', '.'))
        if not os.path.exists(export_dir):
            try:
                # Ask user if they want to create the directory
                if messagebox.askyesno("Directory Not Found", 
                                     f"Export directory '{export_dir}' not found.\nCreate it now?"):
                    os.makedirs(export_dir)
                    self.log_message("essential", f"Created export directory: {export_dir}")
                else:
                    # Fallback to executable directory if user declines
                    if hasattr(sys, '_MEIPASS'):
                        export_dir = os.path.normpath(sys._MEIPASS)
                    else:
                        export_dir = os.path.normpath(os.path.dirname(os.path.abspath(__file__)))
                    self.log_message("essential", f"Using fallback directory: {export_dir}")
            except OSError as e:
                self.log_message("essential", f"Error creating dir '{export_dir}': {e}")
                # Fallback to executable directory
                if hasattr(sys, '_MEIPASS'):
                    export_dir = os.path.normpath(sys._MEIPASS)
                else:
                    export_dir = os.path.normpath(os.path.dirname(os.path.abspath(__file__)))
                self.log_message("essential", f"Using fallback directory: {export_dir}")
        # Construct and normalize the full filepath before adding to queue
        filepath = os.path.normpath(os.path.join(export_dir, filename))

        segment_to_save = list(self.current_segment)

        try:
            if not segment_to_save:
                raise ValueError("Segment data empty.")
            audio_array = np.concatenate(segment_to_save, axis=0)
            if audio_array.size == 0:
                raise ValueError("Concatenated segment empty.")
            self.log_message("everything", f"Segment audio shape: {audio_array.shape}")
        except ValueError as e:
            self.log_message("essential", f"Error: {e}")
            return
        except Exception as e:
            self.log_message("essential", f"Error concatenating segment: {e}")
            return

        try:
            with wave.open(filepath, 'wb') as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(44100)
                wf.writeframes(audio_array.tobytes())
            self.log_message("essential", f"Segment saved to {filepath} at {datetime.datetime.now()}")
            self.transcription_queue.put(filepath)
            self.log_message("extended", f"Added {filepath} to transcription queue at {datetime.datetime.now()}. Queue size: {self.transcription_queue.qsize()}")
        except Exception as e:
            self.log_message("essential", f"Error saving segment wave file: {e}")

    def transcription_worker(self):
        self.log_message("essential", "Transcription worker thread started.")
        while True:
            try:
                filepath = self.transcription_queue.get(timeout=1.0)
                if filepath is AUDIO_QUEUE_SENTINEL:
                    self.log_message("essential", "Worker received sentinel. Stopping transcription worker.")
                    break
                # Normalize the path immediately after retrieving from queue
                norm_filepath = os.path.normpath(filepath)
                self.log_message("essential", f"Worker started processing: {norm_filepath} at {datetime.datetime.now()}. Queue size: {self.transcription_queue.qsize()}")
                # Use normalized path for checks and transcription call
                if os.path.exists(norm_filepath) and not norm_filepath.endswith(".transcribed"):
                    self.transcribe_audio(norm_filepath)
                else:
                    self.log_message("extended", f"Worker skipping: {norm_filepath} (file does not exist or already transcribed)")
                self.transcription_queue.task_done()
                # Use normalized path in the finished log message
                self.log_message("extended", f"Finished processing {norm_filepath}. Remaining queue size: {self.transcription_queue.qsize()}")
            except queue.Empty:
                # No need to log queue empty message frequently if logging level is low
                if self.logging_level in ["Extended", "Everything"]:
                    self.log_message("extended", f"Transcription queue empty at {datetime.datetime.now()}. Waiting for new files...")
                continue
            except Exception as e:
                # Normalize path for error logging too
                log_path = os.path.normpath(filepath) if 'filepath' in locals() else 'unknown file'
                self.log_message("essential", f"Error in transcription worker for {log_path}: {e}")
                if 'filepath' in locals():
                    self.transcription_queue.task_done()

    def start_transcription_worker(self):
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.log_message("extended", "Transcription worker thread already running.")
            return
        self.transcription_worker_thread = threading.Thread(target=self.transcription_worker, daemon=True)
        self.transcription_worker_thread.start()
        self.log_message("essential", "Transcription worker thread started.")

    def handle_recording_error(self):
        self.log_message("essential", "Handling recording error.")
        self.recording = False
        self.destroy_green_line()

    def transcribe_audio(self, audio_path):
        self.log_message("essential", f"Transcription task started for: {audio_path}")
        language = self.language_combobox.get() if hasattr(self, 'language_combobox') and self.language_combobox.get() else getattr(self, 'language', 'en')
        model = self.model_combobox.get() if hasattr(self, 'model_combobox') and self.model_combobox.get() else getattr(self, 'model', 'large')
        translation_enabled = self.translation_var.get() if hasattr(self, 'translation_var') else getattr(self, 'translation_enabled', False)
        command_mode = self.command_mode_var.get() if hasattr(self, 'command_mode_var') else getattr(self, 'command_mode', False)
        timestamps_disabled = self.timestamps_disabled_var.get() if hasattr(self, 'timestamps_disabled_var') else getattr(self, 'timestamps_disabled', False)
        clear_text_output = self.clear_text_output_var.get() if hasattr(self, 'clear_text_output_var') else getattr(self, 'clear_text_output', False)
        export_dir = os.path.normpath(getattr(self, 'export_folder', '.')) # Normalize export dir
        whisper_exec = os.path.normpath(getattr(self, 'whisper_executable', 'whisper')) # Normalize executable path
        prompt_content = self.prompt_text.get("1.0", tk.END).strip() if hasattr(self, 'prompt_text') else getattr(self, 'prompt', '')
        norm_audio_path = os.path.normpath(audio_path) # Normalize audio input path

        self.log_message("extended", f"Command mode enabled: {command_mode}")

        output_filename_base = f"{os.path.splitext(os.path.basename(norm_audio_path))[0]}"
        output_filepath_txt = os.path.join(export_dir, f"{output_filename_base}.txt")
        norm_output_filepath_txt = os.path.normpath(output_filepath_txt) # Normalize output text path

        command = [whisper_exec, norm_audio_path, "--model", model, "--language", language, "--output_format", "txt", "--output_dir", export_dir]
        if prompt_content:
            command.extend(["--initial_prompt", prompt_content])
        command.extend(["--task", "translate" if translation_enabled else "transcribe"])

        transcription_successful = False
        try:
            self.log_message("extended", f"Running Whisper command: {' '.join(command)}")
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            # Ensure command list elements are strings
            command_str = [str(c) for c in command]
            self.log_message("extended", f"Running Whisper command (normalized): {' '.join(command_str)}")
            result = subprocess.run(command_str, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')
            # Use normalized path for checking existence
            if os.path.exists(norm_output_filepath_txt):
                self.log_message("essential", f"Transcription saved to {norm_output_filepath_txt}")
                transcription_successful = True
                # Use normalized path for parsing
                parsed_text = self.parse_transcription_text(norm_output_filepath_txt)
                self.root.after(0, self.update_scratchpad_text, parsed_text)
                self.root.after(0, self._update_clipboard, parsed_text)
                if self.beep_on_transcription:
                    try:
                        winsound.Beep(500, 200)
                        self.log_message("extended", "Played beep for transcription.")
                    except Exception as e:
                        self.log_message("essential", f"Error playing beep: {e}")
                if self.auto_paste:
                    delay_ms = int(self.auto_paste_delay * 1000)
                    self.root.after(delay_ms, self.perform_auto_paste)
                if command_mode:
                    self.log_message("extended", f"Processing commands with transcription: '{parsed_text}'")
                    current_commands_snapshot = list(self.loaded_commands)
                    self.log_message("extended", f"Loaded commands: {current_commands_snapshot}")
                    self.root.after(0, self.execute_command_from_text, parsed_text, current_commands_snapshot)
            else:
                self.log_message("essential", f"Whisper ran, but output file {norm_output_filepath_txt} not found.")
        except subprocess.CalledProcessError as e:
            self.log_message("essential", f"Error running whisper: {e}\nStdout: {e.stdout}\nStderr: {e.stderr}")
        except FileNotFoundError:
            self.log_message("essential", f"Error: Whisper executable not found at '{whisper_exec}'.")
        except Exception as e:
            self.log_message("essential", f"Transcription error: {e}")
        finally:
            # Use normalized path for renaming check and operation
            if transcription_successful and os.path.exists(norm_audio_path):
                try:
                    transcribed_path = norm_audio_path + ".transcribed"
                    norm_transcribed_path = os.path.normpath(transcribed_path) # Normalize renamed path
                    shutil.move(norm_audio_path, norm_transcribed_path)
                    self.log_message("extended", f"Renamed audio file to: {norm_transcribed_path}")
                except Exception as e_mv:
                    self.log_message("essential", f"Error renaming audio file {norm_audio_path}: {e_mv}")
            elif os.path.exists(norm_audio_path):
                 # Log if original audio still exists but wasn't renamed (e.g., transcription failed)
                 self.log_message("extended", f"Audio file {norm_audio_path} still exists (transcription might have failed).")
                 pass # Keep the original audio if transcription failed
        self.log_message("essential", f"Transcription task finished for: {norm_audio_path}")

    def perform_auto_paste(self):
        try:
            keyboard.press_and_release('ctrl+v')
            self.log_message("extended", "Performed auto-paste (Ctrl+V).")
        except Exception as e:
            self.log_message("essential", f"Error performing auto-paste: {e}")

    def parse_transcription_text(self, filepath):
        self.log_message("extended", f"Parsing transcription file: {filepath}")
        try:
            with open(filepath, "r", encoding='utf-8') as f:
                full_text = f.read()
            self.log_message("everything", f"Raw transcription text: '{full_text}'")
            if not full_text.strip():
                self.log_message("extended", "Transcription file is empty.")
                return ""

            if not self.clear_text_output and not self.timestamps_disabled:
                self.log_message("everything", "Returning raw text as clear_text_output and timestamps_disabled are False.")
                return full_text.strip()

            lines = full_text.split('\n')
            cleaned_lines = []
            timestamp_pattern = re.compile(r'^\[\s*\d{2}:\d{2}\.\d{3}\s*-->\s*\d{2}:\d{2}\.\d{3}\s*\]\s*')

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if self.clear_text_output and (line.startswith("===") or re.match(r'^\d{4}-\d{2}-\d{2}', line)):
                    self.log_message("everything", f"Skipping header line: '{line}'")
                    continue
                match = timestamp_pattern.match(line)
                if match and (self.timestamps_disabled or self.clear_text_output):
                    text_part = line[match.end():].strip()
                    if text_part:
                        cleaned_lines.append(text_part)
                        self.log_message("everything", f"Extracted text after timestamp: '{text_part}'")
                else:
                    cleaned_lines.append(line)
                    self.log_message("everything", f"Keeping line: '{line}'")

            cleaned_text = "\n".join(cleaned_lines).strip()
            if cleaned_text.endswith('.'):
                cleaned_text = cleaned_text[:-1]
                self.log_message("everything", f"Removed trailing period: '{cleaned_text}'")
            self.log_message("extended", f"Parsed transcription text: '{cleaned_text}'")
            return cleaned_text
        except FileNotFoundError:
            self.log_message("essential", f"File not found for parsing: {filepath}")
            return ""
        except Exception as e:
            self.log_message("essential", f"Error parsing transcription {filepath}: {e}")
            return full_text.strip() if 'full_text' in locals() else ""

    def execute_command_from_text(self, transcription_text, commands_list):
        self.log_message("extended", f"Attempting to execute command based on text: '{transcription_text}'")
        if not transcription_text.strip():
            self.log_message("extended", "Transcription text is empty, skipping command execution.")
            return
        if not commands_list:
            self.log_message("extended", "No commands defined in loaded_commands.")
            return
        try:
            cleaned_transcription = transcription_text.lower().strip()
            for command_data in commands_list:
                voice_cmd = command_data.get("voice", "").strip().lower()
                action_template = command_data.get("action", "").strip()
                if not voice_cmd or not action_template:
                    self.log_message("extended", f"Skipping invalid command: voice='{voice_cmd}', action='{action_template}'")
                    continue
                escaped_voice_cmd = re.escape(voice_cmd).replace(r'\,', r'\s*,?\s*')
                pattern_str = r'\b' + escaped_voice_cmd.replace(r'\ ff\ ', r'(.*)') + r'\b\.?'
                self.log_message("everything", f"Matching against pattern: '{pattern_str}'")
                match = re.search(pattern_str, cleaned_transcription, re.IGNORECASE)
                if match:
                    action = action_template
                    if match.groups():
                        wildcard_value = match.group(1).strip()
                        self.log_message("extended", f"Wildcard ' FF ' matched: '{wildcard_value}'")
                        action = action.replace(" FF ", wildcard_value)
                    self.log_message("essential", f"Executing command: '{action}'")
                    threading.Thread(target=self.run_subprocess, args=(action,), daemon=True).start()
                    return
                else:
                    self.log_message("everything", f"No match for pattern: '{pattern_str}' against '{cleaned_transcription}'")
        except Exception as e:
            self.log_message("essential", f"Error executing command: {e}")

    def run_subprocess(self, action):
        self.log_message("essential", f"Running subprocess: '{action}'")
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(action, shell=True, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')
            self.log_message("extended", f"Command '{action}' executed successfully. Output: {result.stdout}")
        except subprocess.CalledProcessError as e:
            self.log_message("essential", f"Command '{action}' failed with exit code {e.returncode}. Stderr: {e.stderr}")
            if "is not recognized" in e.stderr.lower():
                self.log_message("essential", f"Hint: '{action}' not found in PATH. Try using the full path.")
        except Exception as e:
            self.log_message("essential", f"Error running subprocess '{action}': {e}")

    def copy_to_clipboard(self, filepath):
        self.log_message("extended", f"Copying {filepath} to clipboard.")
        try:
            parsed_text = self.parse_transcription_text(filepath)
            self._update_clipboard(parsed_text)
        except FileNotFoundError:
            self.log_message("essential", f"File not found for clipboard: {filepath}")
        except Exception as e:
            self.log_message("essential", f"Error reading/parsing file for clipboard: {e}")

    def _update_clipboard(self, text):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.log_message("extended", "Clipboard updated.")
        except Exception as e:
            self.log_message("essential", f"Error updating clipboard: {e}")

    # Removed transcode_last_recording method

    def create_backup(self, filepath):
        if not getattr(self, 'versioning_enabled', False):
            return
        if not getattr(self, 'versioning_enabled', False) or not os.path.exists(filepath):
            return

        backup_folder = getattr(self, 'backup_folder', 'OldVersions')
        max_backups = getattr(self, 'max_backups', 10)

        if not os.path.exists(backup_folder):
            try:
                os.makedirs(backup_folder)
                self.log_message("extended", f"Created backup folder: {backup_folder}")
            except OSError as e:
                self.log_message("essential", f"Error creating backup folder '{backup_folder}': {e}")
                return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)
        backup_filename = f"{name}_{timestamp}{ext}"
        backup_filepath = os.path.join(backup_folder, backup_filename)

        try:
            shutil.copy2(filepath, backup_filepath)
            self.log_message("extended", f"Created backup: {backup_filepath}")
            self.manage_backups(backup_folder, name, ext, max_backups)
        except Exception as e:
            self.log_message("essential", f"Error creating backup: {e}")

    def manage_backups(self, backup_folder, filename_base, file_extension, max_backups):
        if max_backups <= 0:
            self.log_message("extended", "Max backups set to 0 or less, skipping backup management.")
            return

        try:
            backup_files = []
            pattern = re.compile(rf"^{re.escape(filename_base)}_\d{{8}}_\d{{6}}{re.escape(file_extension)}$")
            for entry in os.listdir(backup_folder):
                if pattern.match(entry):
                    full_path = os.path.join(backup_folder, entry)
                    if os.path.isfile(full_path):
                        backup_files.append((full_path, os.path.getmtime(full_path)))

            if len(backup_files) > max_backups:
                backup_files.sort(key=lambda x: x[1]) # Sort by modification time (oldest first)
                files_to_delete = backup_files[:len(backup_files) - max_backups]
                self.log_message("extended", f"Found {len(backup_files)} backups for {filename_base}{file_extension}, exceeding limit of {max_backups}. Deleting {len(files_to_delete)} oldest.")
                for file_path, _ in files_to_delete:
                    try:
                        os.remove(file_path)
                        self.log_message("extended", f"Deleted old backup: {file_path}")
                    except Exception as e:
                        self.log_message("essential", f"Error deleting old backup {file_path}: {e}")
            else:
                self.log_message("extended", f"Number of backups for {filename_base}{file_extension} ({len(backup_files)}) is within the limit of {max_backups}.")
        except Exception as e:
            self.log_message("essential", f"Error managing backups in {backup_folder}: {e}")

    def return_to_window(self):
        self.log_message("essential", "Return to window requested.")
        self.root.after(0, self._show_window)

    def _show_window(self):
        try:
            self.log_message("essential", "Showing window.")
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except tk.TclError as e:
            self.log_message("essential", f"Error showing window: {e}")

    def store_settings_and_hide(self):
        self.log_message("essential", "Storing settings and hiding window...")
        self.root.withdraw()

    def update_audio_device(self):
        device_index = getattr(self, 'selected_audio_device_index', None)
        if device_index is not None:
            try:
                devices = sd.query_devices()
                if 0 <= device_index < len(devices):
                    sd.default.device = device_index
                    self.log_message("essential", f"Set default audio input device to index: {device_index} ({devices[device_index]['name']})")
                else:
                    self.log_message("essential", f"Invalid audio device index: {device_index}")
                    self.selected_audio_device_index = None
            except Exception as e:
                self.log_message("essential", f"Failed to set audio device index '{device_index}': {e}")

    def open_configuration_window(self):
        if self.config_window and self.config_window.winfo_exists():
            self.config_window.lift()
            return

        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("Configuration")
        self.config_window.geometry("450x650")
        self.config_window.transient(self.root)

        pad_options = {'padx': 10, 'pady': 5}

        tk.Label(self.config_window, text="Whisper Executable Path:").pack(anchor=tk.W, **pad_options)
        whisper_frame = tk.Frame(self.config_window)
        whisper_frame.pack(fill=tk.X, **pad_options)
        self.whisper_executable_entry = ttk.Entry(whisper_frame)
        self.whisper_executable_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.whisper_executable_entry.insert(0, getattr(self, 'whisper_executable', 'whisper'))
        ttk.Button(whisper_frame, text="Browse...", command=self.browse_whisper_executable).pack(side=tk.LEFT, padx=(5,0))

        tk.Label(self.config_window, text="Audio Input Device:").pack(anchor=tk.W, **pad_options)
        self.device_list_strings = []
        self.device_index_map = {}
        try:
            devices = sd.query_devices()
            for i, d in enumerate(devices):
                if d['max_input_channels'] > 0:
                    host_api_name = sd.query_hostapis(d['hostapi'])['name']
                    display_string = f"[{i}] {d['name']} ({host_api_name})"
                    self.device_list_strings.append(display_string)
                    self.device_index_map[display_string] = i
            if not self.device_list_strings:
                self.device_list_strings = ["No input devices found"]
        except Exception as e:
            self.log_message("essential", f"Error querying devices: {e}")
            self.device_list_strings = ["Error querying devices"]
        self.audio_device_combobox = ttk.Combobox(self.config_window, values=self.device_list_strings, state="readonly")
        current_index = getattr(self, 'selected_audio_device_index', None)
        current_display_string = None
        for display_str, index in self.device_index_map.items():
            if index == current_index:
                current_display_string = display_str
                break
        if current_display_string:
            self.audio_device_combobox.set(current_display_string)
        elif self.device_list_strings[0] not in ["Error querying devices", "No input devices found"]:
            self.audio_device_combobox.set(self.device_list_strings[0])
        self.audio_device_combobox.pack(fill=tk.X, **pad_options)

        tk.Label(self.config_window, text="Export Folder:").pack(anchor=tk.W, **pad_options)
        export_frame = tk.Frame(self.config_window)
        export_frame.pack(fill=tk.X, **pad_options)
        self.export_folder_entry = ttk.Entry(export_frame)
        self.export_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.export_folder_entry.insert(0, getattr(self, 'export_folder', '.'))
        ttk.Button(export_frame, text="Browse...", command=self.browse_export_folder).pack(side=tk.LEFT, padx=(5,0))

        vad_frame = ttk.LabelFrame(self.config_window, text="Auto-Pause (VAD) Settings", padding=(10, 5))
        vad_frame.pack(fill=tk.X, **pad_options)
        ttk.Label(vad_frame, text="Silence Duration (s):").pack(side=tk.LEFT, padx=(0,5))
        self.silence_duration_entry = ttk.Entry(vad_frame, width=5)
        self.silence_duration_entry.insert(0, str(getattr(self, 'silence_threshold_seconds', 5.0)))
        self.silence_duration_entry.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(vad_frame, text="Energy Threshold:").pack(side=tk.LEFT, padx=(0,5))

        self.vad_energy_entry = ttk.Entry(vad_frame, width=7)
        self.vad_energy_entry.insert(0, str(getattr(self, 'vad_energy_threshold', 300)))
        self.vad_energy_entry.pack(side=tk.LEFT)

        logging_frame = ttk.LabelFrame(self.config_window, text="Logging Settings", padding=(10, 5))
        logging_frame.pack(fill=tk.X, **pad_options)
        ttk.Label(logging_frame, text="Logging Level:").pack(side=tk.LEFT, padx=(0,5))
        self.logging_level_combobox = ttk.Combobox(logging_frame, values=["None", "Essential", "Extended", "Everything"], state="readonly", width=12)
        self.logging_level_combobox.set(self.logging_level)
        self.logging_level_combobox.pack(side=tk.LEFT, padx=(0, 10))
        self.logging_level_combobox.bind("<<ComboboxSelected>>", lambda event: self.update_logging_level())
        self.log_to_file_var = tk.BooleanVar(value=self.log_to_file)
        ttk.Checkbutton(logging_frame, text="Log to File", variable=self.log_to_file_var).pack(side=tk.LEFT)

        transcription_frame = ttk.LabelFrame(self.config_window, text="Transcription Settings", padding=(10, 5))
        transcription_frame.pack(fill=tk.X, **pad_options)
        self.beep_on_transcription_var = tk.BooleanVar(value=self.beep_on_transcription)
        ttk.Checkbutton(transcription_frame, text="Beep on Transcription", variable=self.beep_on_transcription_var).pack(anchor=tk.W)
        self.auto_paste_var = tk.BooleanVar(value=self.auto_paste)
        ttk.Checkbutton(transcription_frame, text="Auto-Paste After Transcription", variable=self.auto_paste_var).pack(anchor=tk.W)
        delay_frame = tk.Frame(transcription_frame)
        delay_frame.pack(anchor=tk.W, padx=(20, 0))
        ttk.Label(delay_frame, text="Auto-Paste Delay (s):").pack(side=tk.LEFT)
        self.auto_paste_delay_entry = ttk.Entry(delay_frame, width=5)
        self.auto_paste_delay_entry.insert(0, str(self.auto_paste_delay))
        self.auto_paste_delay_entry.pack(side=tk.LEFT)

        cleanup_frame = ttk.LabelFrame(self.config_window, text="Session File Cleanup", padding=(10, 5))
        cleanup_frame.pack(fill=tk.X, **pad_options)
        self.clear_audio_var = tk.BooleanVar(value=getattr(self, 'clear_audio_on_exit', False))
        self.clear_text_var = tk.BooleanVar(value=getattr(self, 'clear_text_on_exit', False))
        ttk.Checkbutton(cleanup_frame, text="Clear Audio (.wav/.transcribed) on Exit", variable=self.clear_audio_var).pack(anchor=tk.W)
        ttk.Checkbutton(cleanup_frame, text="Clear Text (.txt) on Exit", variable=self.clear_text_var).pack(anchor=tk.W)
        ttk.Button(cleanup_frame, text="Delete Existing Session Files Now", command=self.delete_session_files).pack(pady=(5,0))

        self.versioning_var_config = tk.BooleanVar(value=getattr(self, 'versioning_enabled', True))
        ttk.Checkbutton(self.config_window, text="Enable File Versioning (Backups)", variable=self.versioning_var_config).pack(anchor=tk.W, **pad_options)

        versioning_frame = ttk.LabelFrame(self.config_window, text="Versioning Settings", padding=(10, 5))
        versioning_frame.pack(fill=tk.X, **pad_options)

        tk.Label(versioning_frame, text="Backup Folder:").pack(side=tk.LEFT, padx=(0, 5))
        self.backup_folder_entry = ttk.Entry(versioning_frame, width=20)
        self.backup_folder_entry.insert(0, getattr(self, 'backup_folder', 'OldVersions'))
        self.backup_folder_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        tk.Label(versioning_frame, text="Max Backups:").pack(side=tk.LEFT, padx=(0, 5))
        self.max_backups_entry = ttk.Entry(versioning_frame, width=5)
        self.max_backups_entry.insert(0, str(getattr(self, 'max_backups', 10)))
        self.max_backups_entry.pack(side=tk.LEFT)

        ttk.Button(self.config_window, text="Configure Commands", command=self.open_command_configuration_window).pack(pady=5)
        ttk.Button(self.config_window, text="Save Configuration", command=self.save_configuration).pack(pady=10)

    def update_logging_level(self):
        self.logging_level = self.logging_level_combobox.get()
        self.log_message("essential", f"Logging level updated to: {self.logging_level}")
        self.save_settings()

    def browse_whisper_executable(self):
        filepath = filedialog.askopenfilename(title="Select Whisper Executable")
        if filepath:
            self.whisper_executable_entry.delete(0, tk.END)
            self.whisper_executable_entry.insert(0, filepath)

    def browse_export_folder(self):
        directory = filedialog.askdirectory(title="Select Export Folder")
        if directory:
            self.export_folder_entry.delete(0, tk.END)
            self.export_folder_entry.insert(0, directory)

    def save_configuration(self):
        self.log_message("essential", "Saving configuration...")
        self.whisper_executable = self.whisper_executable_entry.get()
        selected_device_string = self.audio_device_combobox.get()
        self.selected_audio_device_index = self.device_index_map.get(selected_device_string, None)
        self.export_folder = self.export_folder_entry.get()
        self.versioning_enabled = self.versioning_var_config.get()
        self.backup_folder = self.backup_folder_entry.get() # Get new setting
        try:
            self.max_backups = int(self.max_backups_entry.get()) # Get new setting
            if self.max_backups < 0:
                raise ValueError("Max backups cannot be negative.")
        except ValueError:
            self.log_message("essential", "Invalid max backups, using default.")
            self.max_backups = 10
        try:
            self.silence_threshold_seconds = float(self.silence_duration_entry.get())
        except ValueError:
            self.log_message("essential", "Invalid silence duration, using default.")
            self.silence_threshold_seconds = 5.0
        try:
            self.vad_energy_threshold = int(self.vad_energy_entry.get())
        except ValueError:
            self.log_message("essential", "Invalid VAD energy threshold, using default.")
            self.vad_energy_threshold = 300
        self.clear_audio_on_exit = self.clear_audio_var.get()
        self.clear_text_on_exit = self.clear_text_var.get()
        self.logging_level = self.logging_level_combobox.get()
        self.log_to_file = self.log_to_file_var.get()
        self.beep_on_transcription = self.beep_on_transcription_var.get()
        self.auto_paste = self.auto_paste_var.get()
        try:
            self.auto_paste_delay = float(self.auto_paste_delay_entry.get())
            if self.auto_paste_delay < 0:
                raise ValueError("Delay cannot be negative.")
        except ValueError:
            self.log_message("essential", "Invalid auto-paste delay, using default.")
            self.auto_paste_delay = 1.0

        self.update_audio_device()
        self.save_settings()
        self.log_message("essential", "Configuration saved.")
        if self.config_window and self.config_window.winfo_exists():
            self.config_window.destroy()

    def open_command_configuration_window(self):
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.lift()
            return
        self.command_config_window = tk.Toplevel(self.root)
        self.command_config_window.title("Configure Commands")
        self.command_config_window.geometry("700x400")
        self.command_config_window.transient(self.root)
        self.create_command_config_widgets()

    def create_command_config_widgets(self):
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
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        self.command_widgets = []
        if not self.loaded_commands:
            self.add_command_row()
        else:
            for cmd in self.loaded_commands:
                self.add_command_row(cmd.get("voice", ""), cmd.get("action", ""))

        ttk.Button(self.command_config_window, text="Add Command Row", command=self.add_command_row).pack(pady=10)
        ttk.Button(self.command_config_window, text="Save Commands", command=self.save_commands_from_ui).pack(pady=5)

    def add_command_row(self, voice_cmd="", action_cmd=""):
        command_frame = tk.Frame(self.scrollable_command_frame)
        command_frame.pack(pady=2, fill=tk.X, expand=True)
        ttk.Label(command_frame, text="Voice Cmd:").pack(side=tk.LEFT, padx=(0, 2))
        voice_entry = ttk.Entry(command_frame, width=30)
        voice_entry.insert(0, voice_cmd)
        voice_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Label(command_frame, text="Action:").pack(side=tk.LEFT, padx=(5, 2))
        action_entry = ttk.Entry(command_frame, width=30)
        action_entry.insert(0, action_cmd)
        action_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        remove_button = ttk.Button(command_frame, text="X", width=3, command=lambda f=command_frame: self.remove_command_row(f))
        remove_button.pack(side=tk.LEFT, padx=(5, 0))
        self.command_widgets.append({"frame": command_frame, "voice": voice_entry, "action": action_entry})

    def remove_command_row(self, command_frame):
        widget_ref_to_remove = None
        for i, ref in enumerate(self.command_widgets):
            if ref["frame"] == command_frame:
                widget_ref_to_remove = ref
                break
        if widget_ref_to_remove:
            command_frame.destroy()
            self.command_widgets.remove(widget_ref_to_remove)

    def save_commands_from_ui(self):
        current_commands = []
        for widget_ref in self.command_widgets:
            voice = widget_ref["voice"].get().strip()
            action = widget_ref["action"].get().strip()
            if voice and action:
                voice = voice.replace("Whisperer ", "Whisperer, ").replace("whisperer ", "whisperer, ")
                current_commands.append({"voice": voice, "action": action})
        self.loaded_commands = current_commands
        self.commands = list(self.loaded_commands)
        self.save_commands()
        self.log_message("essential", f"Commands saved: {self.loaded_commands}")
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.destroy()

    def import_prompt(self):
        filename = filedialog.askopenfilename(initialdir=".", title="Select Prompt File", filetypes=(("Text files", "*.txt*"), ("Markdown files", "*.md*"), ("all files", "*.*")))
        if filename:
            try:
                with open(filename, "r", encoding='utf-8') as f:
                    self.prompt_text.delete("1.0", tk.END)
                    content = f.read()
                    self.prompt_text.insert("1.0", content)
                    self.prompt = content
                    self.save_prompt()
                self.log_message("essential", f"Prompt imported from {filename}")
            except Exception as e:
                self.log_message("essential", f"Error importing prompt: {e}")

    def export_prompt(self):
        filename = filedialog.asksaveasfilename(initialdir=".", title="Save Prompt As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("all files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                with open(filename, "w", encoding='utf-8') as f:
                    f.write(self.prompt_text.get("1.0", tk.END))
                self.log_message("essential", f"Prompt exported to {filename}")
            except Exception as e:
                self.log_message("essential", f"Error exporting prompt: {e}")

    def open_scratchpad_window(self):
        if self.scratchpad_window and self.scratchpad_window.winfo_exists():
            self.scratchpad_window.lift()
            return

        self.scratchpad_window = tk.Toplevel(self.root)
        self.scratchpad_window.title("Scratchpad")
        self.scratchpad_window.geometry("500x700")

        text_frame = tk.Frame(self.scratchpad_window)
        text_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.scratchpad_text_widget = tk.Text(text_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.scratchpad_text_widget.yview)
        self.scratchpad_text_widget.configure(yscrollcommand=scrollbar.set)
        self.scratchpad_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        controls_frame = tk.Frame(self.scratchpad_window)
        controls_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(controls_frame, text="Import", command=self.import_to_scratchpad).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(controls_frame, text="Export", command=self.export_from_scratchpad).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        append_frame = tk.Frame(self.scratchpad_window)
        append_frame.pack(pady=5, padx=10, fill=tk.X)
        self.scratchpad_append_var = tk.BooleanVar(value=self.scratchpad_append_mode)
        ttk.Checkbutton(append_frame, text="Append Mode", variable=self.scratchpad_append_var, command=self.toggle_scratchpad_append).pack(side=tk.LEFT)

    def toggle_scratchpad_append(self):
        self.scratchpad_append_mode = self.scratchpad_append_var.get()
        self.log_message("essential", f"Scratchpad Append Mode: {self.scratchpad_append_mode}")

    def update_scratchpad_text(self, text_to_add):
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget):
            return
        try:
            if self.scratchpad_append_mode:
                if not self.clear_text_output and not self.timestamps_disabled:
                    separator = "\n" + "="*20 + f"\n{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n" + "="*20 + "\n"
                    self.scratchpad_text_widget.insert(tk.END, separator + text_to_add)
                else:
                    self.scratchpad_text_widget.insert(tk.END, "\n" + text_to_add)
                self.scratchpad_text_widget.see(tk.END)
            else:
                self.scratchpad_text_widget.delete("1.0", tk.END)
                self.scratchpad_text_widget.insert("1.0", text_to_add)
        except Exception as e:
            self.log_message("essential", f"Error updating scratchpad: {e}")

    def import_to_scratchpad(self):
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget):
            return
        filename = filedialog.askopenfilename(initialdir=".", title="Import to Scratchpad", filetypes=(("Text files", "*.txt*"), ("Markdown files", "*.md*"), ("all files", "*.*")))
        if filename:
            try:
                with open(filename, "r", encoding='utf-8') as f:
                    content = f.read()
                self.scratchpad_text_widget.delete("1.0", tk.END)
                self.scratchpad_text_widget.insert("1.0", content)
            except Exception as e:
                self.log_message("essential", f"Error importing to scratchpad: {e}")

    def export_from_scratchpad(self):
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget):
            return
        filename = filedialog.asksaveasfilename(initialdir=".", title="Export Scratchpad As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("all files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                with open(filename, "w", encoding='utf-8') as f:
                    f.write(self.scratchpad_text_widget.get("1.0", tk.END))
            except Exception as e:
                self.log_message("essential", f"Error exporting from scratchpad: {e}")

    def delete_session_files(self, ask_confirm=True):
        export_dir = getattr(self, 'export_folder', '.')
        if not os.path.isdir(export_dir):
            self.log_message("essential", f"Export directory not found: {export_dir}")
            return

        files_to_delete = []
        delete_audio = getattr(self, 'clear_audio_on_exit', False) if not ask_confirm else self.clear_audio_var.get()
        delete_text = getattr(self, 'clear_text_on_exit', False) if not ask_confirm else self.clear_text_var.get()

        if not (delete_audio or delete_text):
            if ask_confirm:
                messagebox.showinfo("Cleanup Info", "Cleanup toggles are disabled.")
            return

        self.log_message("extended", f"Scanning {export_dir} for session files...")
        for filename in os.listdir(export_dir):
            filepath = os.path.join(export_dir, filename)
            if delete_audio and filename.startswith("recording_") and (filename.endswith(".wav") or filename.endswith(".wav.transcribed")):
                files_to_delete.append(filepath)
            elif delete_text and filename.startswith("recording_") and filename.endswith(".txt"):
                files_to_delete.append(filepath)

        if not files_to_delete:
            msg = "No session files found to delete."
            self.log_message("essential", msg)
            if ask_confirm:
                messagebox.showinfo("Cleanup Info", msg)
            return

        confirm = True
        if ask_confirm:
            confirm = messagebox.askyesno("Confirm Deletion", f"Delete {len(files_to_delete)} session file(s) from '{export_dir}'?\n(Audio: {delete_audio}, Text: {delete_text})")

        if confirm:
            deleted_count = 0
            errors = 0
            for filepath in files_to_delete:
                try:
                    os.remove(filepath)
                    deleted_count += 1
                except Exception as e:
                    self.log_message("essential", f"Error deleting {filepath}: {e}")
                    errors += 1
            result_msg = f"Deleted {deleted_count} file(s)."
            if errors > 0:
                result_msg += f" Failed to delete {errors}."
            self.log_message("essential", result_msg)
            if ask_confirm:
                messagebox.showinfo("Cleanup Result", result_msg)
        else:
            self.log_message("essential", "File deletion cancelled.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VoiceControlApp(root)

    app.create_backup("voice_control_app.py")

    tray_thread = threading.Thread(target=app.setup_tray_icon_thread, daemon=True)
    tray_thread.start()

    app.log_message("essential", "Starting Tkinter mainloop...")
    try:
        root.mainloop()
    except KeyboardInterrupt:
        app.log_message("essential", "KeyboardInterrupt received, initiating shutdown...")
        app.quit_app_action()

    app.log_message("essential", "Tkinter mainloop finished.")
    app.log_message("essential", "Application exiting.")
