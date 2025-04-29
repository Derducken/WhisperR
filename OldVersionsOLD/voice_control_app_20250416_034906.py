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

CONFIG_FILE = "config.json"
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

        self.load_settings()
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app_action)
        self.update_audio_device()
        self.start_transcription_worker()

    def load_settings(self):
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
                    self.prompt = settings.get('prompt', '')
                    self.selected_audio_device_index = settings.get('selected_audio_device_index', None)
                    self.export_folder = settings.get('export_folder', '.')
                    try:
                        self.silence_threshold_seconds = float(settings.get('silence_threshold_seconds', 5.0))
                    except (ValueError, TypeError):
                        self.silence_threshold_seconds = 5.0
                    try:
                        self.vad_energy_threshold = int(settings.get('vad_energy_threshold', 300))
                    except (ValueError, TypeError):
                        self.vad_energy_threshold = 300
                    self.clear_audio_on_exit = settings.get('clear_audio_on_exit', False)
                    self.clear_text_on_exit = settings.get('clear_text_on_exit', False)
                    loaded = settings.get('commands', [])
                    if isinstance(loaded, list):
                        self.loaded_commands = [cmd for cmd in loaded if isinstance(cmd, dict) and "voice" in cmd and "action" in cmd]
                    else:
                        self.loaded_commands = []
                    print(f"Settings loaded from {CONFIG_FILE}")
            else:
                print(f"{CONFIG_FILE} not found, using defaults.")
        except Exception as e:
            print(f"Error loading settings: {e}. Using defaults.")
        self.commands = list(self.loaded_commands)
        self.vad_enabled = self.command_mode

    def save_settings(self):
        if hasattr(self, 'translation_var'):
            self.translation_enabled = self.translation_var.get()
        if hasattr(self, 'command_mode_var'):
            self.command_mode = self.command_mode_var.get()
            self.vad_enabled = self.command_mode
        if hasattr(self, 'timestamps_disabled_var'):
            self.timestamps_disabled = self.timestamps_disabled_var.get()
        if hasattr(self, 'clear_text_output_var'):
            self.clear_text_output = self.clear_text_output_var.get()
        if hasattr(self, 'prompt_text'):
            self.prompt = self.prompt_text.get("1.0", tk.END).strip()

        settings = {
            'versioning_enabled': getattr(self, 'versioning_enabled', True),
            'whisper_executable': getattr(self, 'whisper_executable', 'whisper'),
            'language': getattr(self, 'language', 'en'),
            'model': getattr(self, 'model', 'large'),
            'translation_enabled': self.translation_enabled,
            'command_mode': self.command_mode,
            'timestamps_disabled': getattr(self, 'timestamps_disabled', False),
            'clear_text_output': getattr(self, 'clear_text_output', False),
            'prompt': self.prompt,
            'selected_audio_device_index': getattr(self, 'selected_audio_device_index', None),
            'export_folder': getattr(self, 'export_folder', '.'),
            'silence_threshold_seconds': getattr(self, 'silence_threshold_seconds', 5.0),
            'vad_energy_threshold': getattr(self, 'vad_energy_threshold', 300),
            'clear_audio_on_exit': getattr(self, 'clear_audio_on_exit', False),
            'clear_text_on_exit': getattr(self, 'clear_text_on_exit', False),
            'commands': getattr(self, 'loaded_commands', [])
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            print(f"Settings saved to {CONFIG_FILE}")
        except Exception as e:
            print(f"Error saving settings: {e}")

    def create_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, padx=10, fill=tk.X)
        lang_frame = tk.Frame(top_frame)
        lang_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(lang_frame, text="Language:").pack(side=tk.LEFT, padx=(0, 5))
        self.language_options = ["en", "es", "fr", "de", "it", "ja", "zh", "el"]  # Added "el"
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
        ttk.Checkbutton(trans_frame, text="Enable Translation", variable=self.translation_var).pack(side=tk.LEFT)
        cmd_frame = tk.Frame(toggle_frame1)
        cmd_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.command_mode_var = tk.BooleanVar(value=self.command_mode)
        self.command_mode_var.trace_add("write", lambda *args: setattr(self, 'vad_enabled', self.command_mode_var.get()))
        ttk.Checkbutton(cmd_frame, text="Enable Auto-Pause / Commands", variable=self.command_mode_var).pack(side=tk.LEFT)

        toggle_frame2 = tk.Frame(self.root)
        toggle_frame2.pack(pady=2, padx=10, fill=tk.X)
        ts_frame = tk.Frame(toggle_frame2)
        ts_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.timestamps_disabled_var = tk.BooleanVar(value=self.timestamps_disabled)
        ttk.Checkbutton(ts_frame, text="Disable Timestamps", variable=self.timestamps_disabled_var).pack(side=tk.LEFT)
        clear_text_frame = tk.Frame(toggle_frame2)
        clear_text_frame.pack(side=tk.RIGHT, padx=(10, 0))
        self.clear_text_output_var = tk.BooleanVar(value=self.clear_text_output)
        ttk.Checkbutton(clear_text_frame, text="Clear Text Output", variable=self.clear_text_output_var).pack(side=tk.LEFT)

        ttk.Label(self.root, text="Prompt:").pack(pady=(10, 0), padx=10, anchor=tk.W)
        self.prompt_text = tk.Text(self.root, height=8)
        self.prompt_text.insert("1.0", self.prompt)
        self.prompt_text.pack(pady=5, padx=10, fill=tk.X)

        self.import_export_frame = tk.Frame(self.root)
        self.import_export_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(self.import_export_frame, text="Import Prompt", command=self.import_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(self.import_export_frame, text="Export Prompt", command=self.export_prompt).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        ok_scratch_frame = tk.Frame(self.root)
        ok_scratch_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Button(ok_scratch_frame, text="OK", command=self.store_settings_and_hide).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(ok_scratch_frame, text="Scratchpad", command=self.open_scratchpad_window).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, padx=10, fill=tk.X, side=tk.BOTTOM, anchor=tk.S)
        shortcut_text = "CTRL+Alt+Space: Toggle Record\nCTRL+Shift+Win+Space: Transcode Last\nCTRL+Alt+Shift+Space: Show Window"
        ttk.Label(bottom_frame, text=shortcut_text, justify=tk.LEFT).pack(side=tk.LEFT, anchor=tk.W)
        ttk.Button(bottom_frame, text="Configuration", command=self.open_configuration_window).pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))

        try:
            keyboard.add_hotkey('ctrl+alt+space', self.toggle_recording, suppress=True)
            keyboard.add_hotkey('ctrl+shift+windows+space', self.transcode_last_recording, suppress=True)
            keyboard.add_hotkey('ctrl+alt+shift+space', self.return_to_window, suppress=True)
            print("Hotkeys registered successfully.")
        except Exception as e:
            print(f"Failed to register hotkeys: {e}")

    def update_language(self):
        self.language = self.language_combobox.get()
        print(f"Language updated to: {self.language}")
        self.save_settings()

    def update_model(self):
        self.model = self.model_combobox.get()
        print(f"Model updated to: {self.model}")
        self.save_settings()

    def setup_tray_icon_thread(self):
        try:
            image = Image.open("icon.png")
            menu = pystray.Menu(pystray.MenuItem("Show", self.show_window_action), pystray.MenuItem("Quit", self.quit_app_action))
            self.tray_icon = pystray.Icon("Voice Control App", image, "Voice Control App", menu)
            print("Running tray icon...")
            self.tray_icon.run()
            print("Tray icon thread finished.")
        except FileNotFoundError:
            print("Error: icon.png not found.")
            self.root.after(100, self.quit_app_action)
        except Exception as e:
            print(f"Error setting up tray icon: {e}")
            self.root.after(100, self.quit_app_action)

    def show_window_action(self):
        self.root.after(0, self._show_window)

    def quit_app_action(self):
        print("Quit action initiated...")
        self.recording = False
        if self.audio_stream:
            self.audio_stream.stop()
            self.audio_stream.close()
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout=0.5)
        # Allow transcription worker to finish processing remaining queue items
        print(f"Queue size before shutdown: {self.transcription_queue.qsize()}")
        self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.transcription_worker_thread.join(timeout=2.0)  # Increased timeout to allow queue processing
        if getattr(self, 'clear_audio_on_exit', False) or getattr(self, 'clear_text_on_exit', False):
            self.delete_session_files(ask_confirm=False)
        if self.tray_icon:
            print("Stopping tray icon...")
            self.tray_icon.stop()
        print("Scheduling root quit...")
        self.save_settings()
        self.root.after(0, self.root.quit)

    def create_green_line(self):
        try:
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            hwnd = win32gui.CreateWindowEx(win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED | win32con.WS_EX_TOOLWINDOW,
                                          "Static", None, win32con.WS_VISIBLE | win32con.WS_POPUP, 0, 0, screen_width, 5, 0, 0, 0, None)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
            win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(0,0,0), int(0.5 * 255), win32con.LWA_ALPHA)
            hdc = win32gui.GetDC(hwnd)
            brush = win32gui.CreateSolidBrush(win32api.RGB(0, 150, 0))
            rect = win32gui.GetClientRect(hwnd)
            win32gui.FillRect(hdc, rect, brush)
            win32gui.ReleaseDC(hwnd, hdc)
            win32gui.DeleteObject(brush)
            self.green_line = hwnd
            print("Green line created.")
        except Exception as e:
            print(f"Error creating green line: {e}")

    def destroy_green_line(self):
        if self.green_line:
            try:
                win32gui.DestroyWindow(self.green_line)
                self.green_line = None
                print("Green line destroyed.")
            except Exception as e:
                print(f"Error destroying green line: {e}")
                self.green_line = None

    def toggle_recording(self):
        print("Toggle recording requested.")
        if self.recording:
            self.root.after(0, self.stop_recording)
        else:
            self.root.after(0, self.start_recording)

    def start_recording(self):
        if self.recording:
            return
        print("Starting recording...")
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

    def stop_recording(self):
        if not self.recording:
            return
        print("Manual stop recording requested...")
        self.recording = False
        self.destroy_green_line()
        if self.audio_buffer:
            print("Processing final audio buffer on manual stop...")
            self.current_segment = list(self.audio_buffer)
            self.save_segment_and_reset_vad()
        else:
            print("No audio in buffer to save on manual stop.")
            self.current_segment = []
            self.is_speaking = False
            self.silence_start_time = None
        self.audio_buffer = deque()
        print(f"Recording stopped. Transcription queue size: {self.transcription_queue.qsize()}")

    def record_audio_continuously(self):
        print("Continuous audio recording thread started.")
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
                print(f"Audio CB Status: {status}", flush=True)
            if not self.recording:
                return

            current_time_monotonic = time.monotonic()

            self.audio_buffer.append(indata.copy())
            if len(self.audio_buffer) > buffer_max_size:
                self.audio_buffer.popleft()

            rms = np.sqrt(np.mean(indata.astype(np.float32)**2))
            print(f"RMS: {rms:.2f} (Threshold: {energy_threshold})", flush=True)

            if rms >= energy_threshold:
                if not self.is_speaking:
                    print("Speech started.")
                    self.is_speaking = True
                    self.current_segment = list(self.audio_buffer)
                else:
                    self.current_segment.append(indata.copy())
                self.silence_start_time = None
            elif self.is_speaking:
                if self.silence_start_time is None:
                    self.silence_start_time = current_time_monotonic
                    print("Silence started...")
                self.current_segment.append(indata.copy())
                if current_time_monotonic - self.silence_start_time >= silence_duration:
                    print(f"Silence threshold ({silence_duration}s) reached.")
                    if self.recording:
                        self.root.after(0, self.save_segment_and_reset_vad)

        self.audio_stream = None
        try:
            print(f"Opening InputStream (Continuous) with device index: {device_index}", flush=True)
            self.audio_stream = sd.InputStream(device=device_index, samplerate=samplerate, channels=channels, dtype=dtype, blocksize=blocksize, callback=audio_callback)
            with self.audio_stream:
                print("InputStream opened (Continuous).", flush=True)
                while self.recording:
                    time.sleep(0.1)
        except sd.PortAudioError as pae:
            print(f"PortAudioError: {pae}", flush=True)
            self.root.after(0, self.handle_recording_error)
        except Exception as e:
            print(f"Error in continuous recording thread: {e}", flush=True)
            self.root.after(0, self.handle_recording_error)
        finally:
            if self.audio_stream and not self.audio_stream.closed:
                print("Closing audio stream.", flush=True)
                self.audio_stream.close()
            print("Continuous audio recording thread finished.", flush=True)

    def save_segment_and_reset_vad(self):
        if self.current_segment:
            self.save_segment()
            self.current_segment = []
            self.is_speaking = False
            self.silence_start_time = None
        else:
            print("Save segment called but no segment data.")

    def save_segment(self):
        if not self.current_segment:
            print("No segment data to save.")
            return
        print(f"Saving segment with {len(self.current_segment)} chunks at {datetime.datetime.now()}")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"recording_{timestamp}.wav"
        export_dir = getattr(self, 'export_folder', '.')
        if not os.path.exists(export_dir):
            try:
                os.makedirs(export_dir)
            except OSError as e:
                print(f"Error creating dir '{export_dir}': {e}")
                export_dir = "."
        filepath = os.path.join(export_dir, filename)

        segment_to_save = list(self.current_segment)

        try:
            if not segment_to_save:
                raise ValueError("Segment data empty.")
            audio_array = np.concatenate(segment_to_save, axis=0)
            if audio_array.size == 0:
                raise ValueError("Concatenated segment empty.")
            print(f"Segment audio shape: {audio_array.shape}")
        except ValueError as e:
            print(f"Error: {e}")
            return
        except Exception as e:
            print(f"Error concatenating segment: {e}")
            return

        try:
            with wave.open(filepath, 'wb') as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(44100)
                wf.writeframes(audio_array.tobytes())
            print(f"Segment saved to {filepath} at {datetime.datetime.now()}")
            self.transcription_queue.put(filepath)
            print(f"Added {filepath} to transcription queue at {datetime.datetime.now()}. Queue size: {self.transcription_queue.qsize()}")
        except Exception as e:
            print(f"Error saving segment wave file: {e}")

    def transcription_worker(self):
        print("Transcription worker thread started.")
        while True:
            try:
                filepath = self.transcription_queue.get(timeout=1.0)  # Add timeout to prevent blocking indefinitely
                if filepath is AUDIO_QUEUE_SENTINEL:
                    print("Worker received sentinel. Stopping transcription worker.")
                    break
                print(f"Worker started processing: {filepath} at {datetime.datetime.now()}. Queue size: {self.transcription_queue.qsize()}")
                if os.path.exists(filepath) and not filepath.endswith(".transcribed"):
                    self.transcribe_audio(filepath)
                else:
                    print(f"Worker skipping: {filepath} (file does not exist or already transcribed)")
                self.transcription_queue.task_done()
                print(f"Finished processing {filepath}. Remaining queue size: {self.transcription_queue.qsize()}")
            except queue.Empty:
                # Queue is empty, continue looping to check for new files
                print(f"Transcription queue empty at {datetime.datetime.now()}. Waiting for new files...")
                continue
            except Exception as e:
                print(f"Error in transcription worker for {filepath if 'filepath' in locals() else 'unknown file'}: {e}")
                if 'filepath' in locals():
                    self.transcription_queue.task_done()

    def start_transcription_worker(self):
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            print("Transcription worker thread already running.")
            return
        self.transcription_worker_thread = threading.Thread(target=self.transcription_worker, daemon=True)
        self.transcription_worker_thread.start()
        print("Transcription worker thread started.")

    def handle_recording_error(self):
        print("Handling recording error.")
        self.recording = False
        self.destroy_green_line()

    def transcribe_audio(self, audio_path):
        print(f"Transcription task started for: {audio_path}")
        language = self.language_combobox.get() if hasattr(self, 'language_combobox') and self.language_combobox.get() else getattr(self, 'language', 'en')
        model = self.model_combobox.get() if hasattr(self, 'model_combobox') and self.model_combobox.get() else getattr(self, 'model', 'large')
        translation_enabled = self.translation_var.get() if hasattr(self, 'translation_var') else getattr(self, 'translation_enabled', False)
        command_mode = self.command_mode_var.get() if hasattr(self, 'command_mode_var') else getattr(self, 'command_mode', False)
        timestamps_disabled = self.timestamps_disabled_var.get() if hasattr(self, 'timestamps_disabled_var') else getattr(self, 'timestamps_disabled', False)
        clear_text_output = self.clear_text_output_var.get() if hasattr(self, 'clear_text_output_var') else getattr(self, 'clear_text_output', False)
        export_dir = getattr(self, 'export_folder', '.')
        whisper_exec = getattr(self, 'whisper_executable', 'whisper')
        prompt_content = self.prompt_text.get("1.0", tk.END).strip() if hasattr(self, 'prompt_text') else getattr(self, 'prompt', '')

        print(f"Command mode enabled: {command_mode}")

        output_filename_base = f"{os.path.splitext(os.path.basename(audio_path))[0]}"
        output_filepath_txt = os.path.join(export_dir, f"{output_filename_base}.txt")

        command = [whisper_exec, audio_path, "--model", model, "--language", language, "--output_format", "txt", "--output_dir", export_dir]
        if prompt_content:
            command.extend(["--initial_prompt", prompt_content])
        command.extend(["--task", "translate" if translation_enabled else "transcribe"])

        transcription_successful = False
        try:
            print(f"Running Whisper command: {' '.join(command)}")
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(command, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')
            expected_output_path_txt = os.path.join(export_dir, f"{os.path.splitext(os.path.basename(audio_path))[0]}.txt")
            if os.path.exists(expected_output_path_txt):
                print(f"Transcription saved to {expected_output_path_txt}")
                transcription_successful = True
                parsed_text = self.parse_transcription_text(expected_output_path_txt)
                self.root.after(0, self.update_scratchpad_text, parsed_text)
                self.root.after(0, self._update_clipboard, parsed_text)
                if command_mode:
                    print(f"Processing commands with transcription: '{parsed_text}'")
                    current_commands_snapshot = list(self.loaded_commands)
                    print(f"Loaded commands: {current_commands_snapshot}")
                    self.root.after(0, self.execute_command_from_text, parsed_text, current_commands_snapshot)
            else:
                print(f"Whisper ran, but output file {expected_output_path_txt} not found.")
        except subprocess.CalledProcessError as e:
            print(f"Error running whisper: {e}\nStdout: {e.stdout}\nStderr: {e.stderr}")
        except FileNotFoundError:
            print(f"Error: Whisper executable not found at '{whisper_exec}'.")
        except Exception as e:
            print(f"Transcription error: {e}")
        finally:
            if transcription_successful and os.path.exists(audio_path):
                try:
                    transcribed_path = audio_path + ".transcribed"
                    shutil.move(audio_path, transcribed_path)
                    print(f"Renamed audio file to: {transcribed_path}")
                except Exception as e_mv:
                    print(f"Error renaming audio file {audio_path}: {e_mv}")
            elif os.path.exists(audio_path):
                pass
        print(f"Transcription task finished for: {audio_path}")

    def parse_transcription_text(self, filepath):
        print(f"Parsing transcription file: {filepath}")
        try:
            with open(filepath, "r", encoding='utf-8') as f:
                full_text = f.read()
            print(f"Raw transcription text: '{full_text}'")
            if not full_text.strip():
                print("Transcription file is empty.")
                return ""

            if not self.clear_text_output and not self.timestamps_disabled:
                print("Returning raw text as clear_text_output and timestamps_disabled are False.")
                return full_text.strip()

            lines = full_text.split('\n')
            cleaned_lines = []
            timestamp_pattern = re.compile(r'^\[\s*\d{2}:\d{2}\.\d{3}\s*-->\s*\d{2}:\d{2}\.\d{3}\s*\]\s*')

            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if self.clear_text_output and (line.startswith("===") or re.match(r'^\d{4}-\d{2}-\d{2}', line)):
                    print(f"Skipping header line: '{line}'")
                    continue
                match = timestamp_pattern.match(line)
                if match and (self.timestamps_disabled or self.clear_text_output):
                    text_part = line[match.end():].strip()
                    if text_part:
                        cleaned_lines.append(text_part)
                        print(f"Extracted text after timestamp: '{text_part}'")
                else:
                    cleaned_lines.append(line)
                    print(f"Keeping line: '{line}'")

            cleaned_text = "\n".join(cleaned_lines).strip()
            if cleaned_text.endswith('.'):
                cleaned_text = cleaned_text[:-1]
                print(f"Removed trailing period: '{cleaned_text}'")
            print(f"Parsed transcription text: '{cleaned_text}'")
            return cleaned_text
        except FileNotFoundError:
            print(f"File not found for parsing: {filepath}")
            return ""
        except Exception as e:
            print(f"Error parsing transcription {filepath}: {e}")
            return full_text.strip() if 'full_text' in locals() else ""

    def execute_command_from_text(self, transcription_text, commands_list):
        print(f"Attempting to execute command based on text: '{transcription_text}'")
        if not transcription_text.strip():
            print("Transcription text is empty, skipping command execution.")
            return
        if not commands_list:
            print("No commands defined in loaded_commands.")
            return
        try:
            cleaned_transcription = transcription_text.lower().strip()
            for command_data in commands_list:
                voice_cmd = command_data.get("voice", "").strip().lower()
                action_template = command_data.get("action", "").strip()
                if not voice_cmd or not action_template:
                    print(f"Skipping invalid command: voice='{voice_cmd}', action='{action_template}'")
                    continue
                # Forgiving regex: handle commas, spaces, periods
                escaped_voice_cmd = re.escape(voice_cmd).replace(r'\,', r'\s*,?\s*')
                pattern_str = r'\b' + escaped_voice_cmd.replace(r'\ ff\ ', r'(.*)') + r'\b\.?'
                print(f"Matching against pattern: '{pattern_str}'")
                match = re.search(pattern_str, cleaned_transcription, re.IGNORECASE)
                if match:
                    action = action_template
                    if match.groups():
                        wildcard_value = match.group(1).strip()
                        print(f"Wildcard ' FF ' matched: '{wildcard_value}'")
                        action = action.replace(" FF ", wildcard_value)
                    print(f"Executing command: '{action}'")
                    threading.Thread(target=self.run_subprocess, args=(action,), daemon=True).start()
                    return
                else:
                    print(f"No match for pattern: '{pattern_str}' against '{cleaned_transcription}'")
        except Exception as e:
            print(f"Error executing command: {e}")

    def run_subprocess(self, action):
        print(f"Running subprocess: '{action}'")
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(action, shell=True, check=True, capture_output=True, text=True, startupinfo=startupinfo, encoding='utf-8', errors='ignore')
            print(f"Command '{action}' executed successfully. Output: {result.stdout}")
        except subprocess.CalledProcessError as e:
            print(f"Command '{action}' failed with exit code {e.returncode}. Stderr: {e.stderr}")
            if "is not recognized" in e.stderr.lower():
                print(f"Hint: '{action}' not found in PATH. Try using the full path (e.g., 'C:\\Program Files\\Mozilla Firefox\\firefox.exe').")
        except Exception as e:
            print(f"Error running subprocess '{action}': {e}")

    def copy_to_clipboard(self, filepath):
        print(f"Copying {filepath} to clipboard.")
        try:
            parsed_text = self.parse_transcription_text(filepath)
            self._update_clipboard(parsed_text)
        except FileNotFoundError:
            print(f"File not found for clipboard: {filepath}")
        except Exception as e:
            print(f"Error reading/parsing file for clipboard: {e}")

    def _update_clipboard(self, text):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            print("Clipboard updated.")
        except Exception as e:
            print(f"Error updating clipboard: {e}")

    def transcode_last_recording(self):
        print("Transcode last recording requested.")
        try:
            export_dir = getattr(self, 'export_folder', '.')
            if not os.path.isdir(export_dir):
                print(f"Export dir not found: {export_dir}")
                return
            wav_files = [os.path.join(export_dir, f) for f in os.listdir(export_dir)
                         if f.startswith("recording_") and f.endswith(".wav") and not f.endswith(".wav.transcribed")]
            if not wav_files:
                print("No untranscribed recordings found.")
                return
            latest_recording = max(wav_files, key=os.path.getmtime)
            print(f"Adding to queue for transcoding: {latest_recording}")
            self.transcription_queue.put(latest_recording)
        except Exception as e:
            print(f"Error finding last recording: {e}")

    def create_backup(self, filepath):
        if not getattr(self, 'versioning_enabled', False):
            return
        if not os.path.exists(filepath):
            print(f"File not found: {filepath}")
            return
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{os.path.splitext(filepath)[0]}_{timestamp}{os.path.splitext(filepath)[1]}"
        try:
            shutil.copy2(filepath, backup_filename)
            print(f"Created backup: {backup_filename}")
        except Exception as e:
            print(f"Error creating backup: {e}")

    def return_to_window(self):
        print("Return to window requested.")
        self.root.after(0, self._show_window)

    def _show_window(self):
        try:
            print("Showing window.")
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except tk.TclError as e:
            print(f"Error showing window: {e}")

    def store_settings_and_hide(self):
        print("Storing settings and hiding window...")
        self.language = self.language_combobox.get()
        self.model = self.model_combobox.get()
        self.translation_enabled = self.translation_var.get()
        self.command_mode = self.command_mode_var.get()
        self.vad_enabled = self.command_mode
        self.timestamps_disabled = self.timestamps_disabled_var.get()
        self.clear_text_output = self.clear_text_output_var.get()
        self.prompt = self.prompt_text.get("1.0", tk.END).strip()
        self.save_settings()
        self.update_audio_device()
        self.root.withdraw()

    def update_audio_device(self):
        device_index = getattr(self, 'selected_audio_device_index', None)
        if device_index is not None:
            try:
                devices = sd.query_devices()
                if 0 <= device_index < len(devices):
                    sd.default.device = device_index
                    print(f"Set default audio input device to index: {device_index} ({devices[device_index]['name']})")
                else:
                    print(f"Invalid audio device index: {device_index}")
                    self.selected_audio_device_index = None
            except Exception as e:
                print(f"Failed to set audio device index '{device_index}': {e}")

    def open_configuration_window(self):
        if self.config_window and self.config_window.winfo_exists():
            self.config_window.lift()
            return

        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("Configuration")
        self.config_window.geometry("450x550")
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
            print(f"Error querying devices: {e}")
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

        cleanup_frame = ttk.LabelFrame(self.config_window, text="Session File Cleanup", padding=(10, 5))
        cleanup_frame.pack(fill=tk.X, **pad_options)
        self.clear_audio_var = tk.BooleanVar(value=getattr(self, 'clear_audio_on_exit', False))
        self.clear_text_var = tk.BooleanVar(value=getattr(self, 'clear_text_on_exit', False))
        ttk.Checkbutton(cleanup_frame, text="Clear Audio (.wav/.transcribed) on Exit", variable=self.clear_audio_var).pack(anchor=tk.W)
        ttk.Checkbutton(cleanup_frame, text="Clear Text (.txt) on Exit", variable=self.clear_text_var).pack(anchor=tk.W)
        ttk.Button(cleanup_frame, text="Delete Existing Session Files Now", command=self.delete_session_files).pack(pady=(5,0))

        self.versioning_var_config = tk.BooleanVar(value=getattr(self, 'versioning_enabled', True))
        ttk.Checkbutton(self.config_window, text="Enable File Versioning (Backups)", variable=self.versioning_var_config).pack(anchor=tk.W, **pad_options)

        ttk.Button(self.config_window, text="Configure Commands", command=self.open_command_configuration_window).pack(pady=5)
        ttk.Button(self.config_window, text="Save Configuration", command=self.save_configuration).pack(pady=10)

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
        print("Saving configuration...")
        self.whisper_executable = self.whisper_executable_entry.get()
        selected_device_string = self.audio_device_combobox.get()
        self.selected_audio_device_index = self.device_index_map.get(selected_device_string, None)
        self.export_folder = self.export_folder_entry.get()
        self.versioning_enabled = self.versioning_var_config.get()
        try:
            self.silence_threshold_seconds = float(self.silence_duration_entry.get())
        except ValueError:
            print("Invalid silence duration, using default.")
            self.silence_threshold_seconds = 5.0
        try:
            self.vad_energy_threshold = int(self.vad_energy_entry.get())
        except ValueError:
            print("Invalid VAD energy threshold, using default.")
            self.vad_energy_threshold = 300
        self.clear_audio_on_exit = self.clear_audio_var.get()
        self.clear_text_on_exit = self.clear_text_var.get()

        self.update_audio_device()
        self.save_settings()
        print("Configuration saved.")
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
        ttk.Button(self.command_config_window, text="Save Commands", command=self.save_commands).pack(pady=5)

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
                del self.command_widgets[i]
                break
        if widget_ref_to_remove:
            command_frame.destroy()

    def save_commands(self):
        current_commands = []
        for widget_ref in self.command_widgets:
            voice = widget_ref["voice"].get().strip()
            action = widget_ref["action"].get().strip()
            if voice and action:
                # Normalize voice command to include comma
                voice = voice.replace("Whisperer ", "Whisperer, ").replace("whisperer ", "whisperer, ")
                current_commands.append({"voice": voice, "action": action})
        self.loaded_commands = current_commands
        self.commands = list(self.loaded_commands)
        self.save_settings()
        print(f"Commands saved: {self.loaded_commands}")
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.destroy()

    def import_prompt(self):
        filename = filedialog.askopenfilename(initialdir=".", title="Select Prompt File", filetypes=(("Text files", "*.txt*"), ("Markdown files", "*.md*"), ("all files", "*.*")))
        if filename:
            try:
                with open(filename, "r", encoding='utf-8') as f:
                    self.prompt_text.delete("1.0", tk.END)
                    self.prompt_text.insert("1.0", f.read())
            except Exception as e:
                print(f"Error importing prompt: {e}")

    def export_prompt(self):
        filename = filedialog.asksaveasfilename(initialdir=".", title="Save Prompt As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("all files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                with open(filename, "w", encoding='utf-8') as f:
                    f.write(self.prompt_text.get("1.0", tk.END))
            except Exception as e:
                print(f"Error exporting prompt: {e}")

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
        print(f"Scratchpad Append Mode: {self.scratchpad_append_mode}")

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
            print(f"Error updating scratchpad: {e}")

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
                print(f"Error importing to scratchpad: {e}")

    def export_from_scratchpad(self):
        if not (self.scratchpad_window and self.scratchpad_window.winfo_exists() and self.scratchpad_text_widget):
            return
        filename = filedialog.asksaveasfilename(initialdir=".", title="Export Scratchpad As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("all files", "*.*")), defaultextension=".txt")
        if filename:
            try:
                with open(filename, "w", encoding='utf-8') as f:
                    f.write(self.scratchpad_text_widget.get("1.0", tk.END))
            except Exception as e:
                print(f"Error exporting from scratchpad: {e}")

    def delete_session_files(self, ask_confirm=True):
        export_dir = getattr(self, 'export_folder', '.')
        if not os.path.isdir(export_dir):
            print(f"Export directory not found: {export_dir}")
            return

        files_to_delete = []
        delete_audio = getattr(self, 'clear_audio_on_exit', False) if not ask_confirm else self.clear_audio_var.get()
        delete_text = getattr(self, 'clear_text_on_exit', False) if not ask_confirm else self.clear_text_var.get()

        if not (delete_audio or delete_text):
            if ask_confirm:
                messagebox.showinfo("Cleanup Info", "Cleanup toggles are disabled.")
            return

        print(f"Scanning {export_dir} for session files...")
        for filename in os.listdir(export_dir):
            filepath = os.path.join(export_dir, filename)
            if delete_audio and filename.startswith("recording_") and (filename.endswith(".wav") or filename.endswith(".wav.transcribed")):
                files_to_delete.append(filepath)
            elif delete_text and filename.startswith("recording_") and filename.endswith(".txt"):
                files_to_delete.append(filepath)

        if not files_to_delete:
            msg = "No session files found to delete."
            print(msg)
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
                    print(f"Error deleting {filepath}: {e}")
                    errors += 1
            result_msg = f"Deleted {deleted_count} file(s)."
            if errors > 0:
                result_msg += f" Failed to delete {errors}."
            print(result_msg)
            if ask_confirm:
                messagebox.showinfo("Cleanup Result", result_msg)
        else:
            print("File deletion cancelled.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VoiceControlApp(root)

    app.create_backup("voice_control_app.py")

    tray_thread = threading.Thread(target=app.setup_tray_icon_thread, daemon=True)
    tray_thread.start()

    print("Starting Tkinter mainloop...")
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("KeyboardInterrupt received, initiating shutdown...")
        app.quit_app_action()

    print("Tkinter mainloop finished.")
    print("Application exiting.")