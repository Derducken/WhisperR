import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import sounddevice as sd
import numpy as np
from collections import deque
import threading
import queue
import time
import os
import subprocess
import shutil
import datetime
import json
import re
import torch
import pystray
from PIL import Image
import wave
import sys

AUDIO_QUEUE_SENTINEL = object()

class VoiceControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Whisperer")
        self.recording = False
        self.audio_buffer = deque()
        self.current_segment = []
        self.is_speaking = False
        self.silence_start_time = None
        self.audio_stream = None
        self.recording_thread = None
        self.transcription_queue = queue.Queue()
        self.transcription_worker_thread = None
        self.vad_model = None
        self.sample_rate = 44100
        self.chunk_duration_ms = 30
        self.chunk_samples = int(self.sample_rate * self.chunk_duration_ms / 1000)
        self.silence_duration_threshold = 0.3
        self.vad_sensitivity = 0.5
        self.command_mode_var = tk.BooleanVar(value=False)
        self.clear_text_output_var = tk.BooleanVar(value=False)
        self.timestamps_disabled_var = tk.BooleanVar(value=False)
        self.translation_var = tk.BooleanVar(value=False)
        self.clear_audio_on_exit = False
        self.clear_text_on_exit = False
        self.loaded_commands = []
        self.commands = []
        self.command_widgets = []
        self.language = "en"
        self.model = "large"
        self.whisper_executable = "whisper"
        self.export_folder = "."
        self.prompt = ""
        self.tray_icon = None
        self.green_line_canvas = None
        self.green_line_id = None
        self.setup_gui()
        self.load_settings()
        self.start_transcription_worker()
        self.bind_hotkeys()
        try:
            self.vad_model = torch.hub.load('snakers4/silero-vad', 'silero_vad', force_reload=False, onnx=False)
        except Exception as e:
            print(f"Error loading VAD model: {e}")

    def setup_gui(self):
        self.scratchpad = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, height=10)
        self.scratchpad.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(button_frame, text="Start Recording", command=self.start_recording).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Stop Recording", command=self.stop_recording).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear Scratchpad", command=self.clear_scratchpad).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Configure", command=self.open_config_window).pack(side=tk.RIGHT, padx=5)

    def open_config_window(self):
        if hasattr(self, 'config_window') and self.config_window.winfo_exists():
            self.config_window.lift()
            return
        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("Configuration")
        self.config_window.geometry("400x550")  # Increased height for button

        tk.Label(self.config_window, text="Whisper Executable:").pack(anchor=tk.W, padx=5, pady=2)
        self.whisper_entry = tk.Entry(self.config_window, width=50)
        self.whisper_entry.pack(fill=tk.X, padx=5, pady=2)
        self.whisper_entry.insert(0, self.whisper_executable)
        tk.Button(self.config_window, text="Browse", command=self.browse_whisper).pack(anchor=tk.W, padx=5, pady=2)

        tk.Label(self.config_window, text="Export Folder:").pack(anchor=tk.W, padx=5, pady=2)
        self.export_entry = tk.Entry(self.config_window, width=50)
        self.export_entry.pack(fill=tk.X, padx=5, pady=2)
        self.export_entry.insert(0, self.export_folder)
        tk.Button(self.config_window, text="Browse", command=self.browse_export).pack(anchor=tk.W, padx=5, pady=2)

        tk.Label(self.config_window, text="Prompt:").pack(anchor=tk.W, padx=5, pady=2)
        self.prompt_text = tk.Text(self.config_window, height=4, width=50)
        self.prompt_text.pack(fill=tk.X, padx=5, pady=2)
        self.prompt_text.insert("1.0", self.prompt)

        tk.Label(self.config_window, text="Language:").pack(anchor=tk.W, padx=5, pady=2)
        self.language_combobox = ttk.Combobox(self.config_window, values=["en", "es", "fr", "de", "it", "ja", "zh", "auto"])
        self.language_combobox.pack(fill=tk.X, padx=5, pady=2)
        self.language_combobox.set(self.language)

        tk.Label(self.config_window, text="Model:").pack(anchor=tk.W, padx=5, pady=2)
        self.model_combobox = ttk.Combobox(self.config_window, values=["tiny", "base", "small", "medium", "large"])
        self.model_combobox.pack(fill=tk.X, padx=5, pady=2)
        self.model_combobox.set(self.model)

        tk.Label(self.config_window, text="Auto-Pause (VAD) Settings:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=5, pady=5)
        tk.Checkbutton(self.config_window, text="Enable Auto-Pause / Commands", variable=self.command_mode_var).pack(anchor=tk.W, padx=5, pady=2)
        tk.Label(self.config_window, text="Silence Duration (seconds):").pack(anchor=tk.W, padx=5, pady=2)
        self.silence_duration_entry = tk.Entry(self.config_window, width=10)
        self.silence_duration_entry.pack(anchor=tk.W, padx=5, pady=2)
        self.silence_duration_entry.insert(0, str(self.silence_duration_threshold))

        tk.Checkbutton(self.config_window, text="Clear Text Output", variable=self.clear_text_output_var).pack(anchor=tk.W, padx=5, pady=2)
        tk.Checkbutton(self.config_window, text="Disable Timestamps", variable=self.timestamps_disabled_var).pack(anchor=tk.W, padx=5, pady=2)
        tk.Checkbutton(self.config_window, text="Enable Translation", variable=self.translation_var).pack(anchor=tk.W, padx=5, pady=2)

        tk.Button(self.config_window, text="Configure Commands", command=self.configure_commands).pack(anchor=tk.W, padx=5, pady=5)

        # Button frame for OK and Start/Stop
        button_frame = tk.Frame(self.config_window)
        button_frame.pack(anchor=tk.W, padx=5, pady=5)
        tk.Button(button_frame, text="OK", command=self.save_config, width=12).pack(side=tk.LEFT, padx=5, pady=(5, 20))
        
        # Start/Stop Recording button and indicator
        start_stop_frame = tk.Frame(button_frame)
        start_stop_frame.pack(side=tk.LEFT, padx=5)
        self.start_stop_button = tk.Button(start_stop_frame, text="Start Recording", command=self.toggle_recording, width=12)
        self.start_stop_button.pack(side=tk.LEFT)
        self.recording_indicator_label = tk.Label(start_stop_frame, text="●", fg="gray", font=("Arial", 12))
        self.recording_indicator_label.pack(side=tk.LEFT, padx=5)
        print("Start/Stop Recording button and indicator created.")

    def toggle_recording(self):
        if self.recording:
            self.stop_recording()
            self.start_stop_button.config(text="Start Recording")
            print("Toggled to stop recording.")
        else:
            self.start_recording()
            self.start_stop_button.config(text="Stop Recording")
            print("Toggled to start recording.")
        self.update_recording_indicator()

    def update_recording_indicator(self):
        color = "red" if self.recording else "gray"
        self.recording_indicator_label.config(fg=color)
        print(f"Recording indicator updated to {color}.")

    def save_config(self):
        self.whisper_executable = self.whisper_entry.get().strip()
        self.export_folder = self.export_entry.get().strip()
        self.prompt = self.prompt_text.get("1.0", tk.END).strip()
        self.language = self.language_combobox.get()
        self.model = self.model_combobox.get()
        try:
            silence_duration = float(self.silence_duration_entry.get().strip())
            if silence_duration >= 0.1:
                self.silence_duration_threshold = silence_duration
                print(f"Set silence duration to {silence_duration} seconds")
            else:
                print("Silence duration must be at least 0.1 seconds. Keeping default.")
        except ValueError:
            print("Invalid silence duration. Keeping default.")
        self.save_settings()
        self.config_window.destroy()

    def browse_whisper(self):
        filepath = filedialog.askopenfilename()
        if filepath:
            self.whisper_entry.delete(0, tk.END)
            self.whisper_entry.insert(0, filepath)

    def browse_export(self):
        folder = filedialog.askdirectory()
        if folder:
            self.export_entry.delete(0, tk.END)
            self.export_entry.insert(0, folder)

    def configure_commands(self):
        if hasattr(self, 'command_config_window') and self.command_config_window.winfo_exists():
            self.command_config_window.lift()
            return
        self.command_config_window = tk.Toplevel(self.root)
        self.command_config_window.title("Configure Commands")
        self.command_config_window.geometry("400x300")

        commands_frame = tk.Frame(self.command_config_window)
        commands_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.command_widgets = []
        for cmd in self.loaded_commands:
            self.add_command_row(commands_frame, cmd.get("voice", ""), cmd.get("action", ""))

        tk.Button(self.command_config_window, text="Add Command", command=lambda: self.add_command_row(commands_frame)).pack(anchor=tk.W, padx=5, pady=2)
        tk.Button(self.command_config_window, text="Save Commands", command=self.save_commands).pack(anchor=tk.W, padx=5, pady=5)

    def add_command_row(self, parent, voice="", action=""):
        row_frame = tk.Frame(parent)
        row_frame.pack(fill=tk.X, padx=5, pady=2)
        voice_entry = tk.Entry(row_frame, width=30)
        voice_entry.pack(side=tk.LEFT, padx=5)
        voice_entry.insert(0, voice)
        action_entry = tk.Entry(row_frame, width=30)
        action_entry.pack(side=tk.LEFT, padx=5)
        action_entry.insert(0, action)
        self.command_widgets.append({"voice": voice_entry, "action": action_entry})

    def save_commands(self):
        current_commands = []
        for widget_ref in self.command_widgets:
            voice = widget_ref["voice"].get().strip()
            action = widget_ref["action"].get().strip()
            if voice and action:
                voice = voice.replace("Whisperer ", "Whisperer, ").replace("whisperer ", "whisperer, ")
                current_commands.append({"voice": voice, "action": action})
        self.loaded_commands = current_commands
        self.commands = list(self.loaded_commands)
        self.save_settings()
        print(f"Commands saved: {self.loaded_commands}")
        if self.command_config_window and self.command_config_window.winfo_exists():
            self.command_config_window.destroy()

    def save_settings(self):
        settings = {
            "whisper_executable": self.whisper_executable,
            "export_folder": self.export_folder,
            "prompt": self.prompt,
            "language": self.language,
            "model": self.model,
            "silence_duration": self.silence_duration_threshold,
            "command_mode": self.command_mode_var.get(),
            "clear_text_output": self.clear_text_output_var.get(),
            "timestamps_disabled": self.timestamps_disabled_var.get(),
            "translation_enabled": self.translation_var.get(),
            "commands": self.loaded_commands
        }
        try:
            with open("config.json", "w", encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            print("Settings saved to config.json")
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        try:
            with open("config.json", "r", encoding='utf-8') as f:
                settings = json.load(f)
            self.whisper_executable = settings.get("whisper_executable", "whisper")
            self.export_folder = settings.get("export_folder", ".")
            self.prompt = settings.get("prompt", "")
            self.language = settings.get("language", "en")
            self.model = settings.get("model", "large")
            self.silence_duration_threshold = settings.get("silence_duration", 0.3)
            self.command_mode_var.set(settings.get("command_mode", False))
            self.clear_text_output_var.set(settings.get("clear_text_output", False))
            self.timestamps_disabled_var.set(settings.get("timestamps_disabled", False))
            self.translation_var.set(settings.get("translation_enabled", False))
            self.loaded_commands = settings.get("commands", [])
            self.commands = list(self.loaded_commands)
            print("Settings loaded from config.json")
        except FileNotFoundError:
            print("No config.json found, using defaults.")
        except Exception as e:
            print(f"Error loading settings: {e}")

    def bind_hotkeys(self):
        try:
            import keyboard
            keyboard.add_hotkey('ctrl+alt+space', self.toggle_recording_hotkey)
            print("Hotkeys bound successfully.")
        except ImportError:
            print("Keyboard module not installed. Hotkeys disabled.")
        except Exception as e:
            print(f"Error binding hotkeys: {e}")

    def toggle_recording_hotkey(self):
        if self.recording:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        if self.recording:
            return
        self.recording = True
        self.recording_thread = threading.Thread(target=self.record_audio, daemon=True)
        self.recording_thread.start()
        self.update_recording_indicator()
        print(f"Started recording with command mode: {self.command_mode_var.get()}")

    def stop_recording(self):
        if not self.recording:
            return
        print("Manual stop recording requested...")
        self.recording = False
        self.destroy_green_line()
        if self.audio_buffer or self.current_segment:
            print("Processing final audio buffer on manual stop...")
            if self.audio_buffer:
                self.current_segment.extend(list(self.audio_buffer))
            self.save_segment()
            self.current_segment = []
            self.audio_buffer = deque()
            self.is_speaking = False
            self.silence_start_time = None
        else:
            print("No audio in buffer to save on manual stop.")
            self.current_segment = []
            self.is_speaking = False
            self.silence_start_time = None
            self.audio_buffer = deque()
        print(f"Recording stopped. Transcription queue size: {self.transcription_queue.qsize()}")
        self.update_recording_indicator()

    def record_audio(self):
        print("Starting recording...")
        self.audio_stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=1,
            dtype='int16',
            blocksize=self.chunk_samples,
            callback=self.audio_callback
        )
        self.audio_stream.start()
        try:
            while self.recording:
                time.sleep(0.1)
                self.process_audio_buffer()
        except KeyboardInterrupt:
            pass
        finally:
            self.audio_stream.stop()
            self.audio_stream.close()
            self.audio_stream = None
            print("Recording stream closed.")

    def audio_callback(self, indata, frames, time_info, status):
        if status:
            print(f"Audio callback status: {status}")
        self.audio_buffer.append(indata.copy())

    def process_audio_buffer(self):
        command_mode = self.command_mode_var.get()
        print(f"Processing audio buffer with command mode: {command_mode}")
        while self.audio_buffer:
            chunk = self.audio_buffer.popleft()
            if chunk is None:
                continue
            chunk_float = chunk.astype(np.float32) / 32768.0
            if self.vad_model is None:
                print("VAD model not initialized.")
                continue
            speech_prob = self.vad_model(torch.from_numpy(chunk_float), self.sample_rate).item()
            current_time = time.time()
            rms = np.sqrt(np.mean(chunk_float ** 2)) * 100
            print(f"RMS: {rms:.2f} (Threshold: {self.vad_sensitivity * 100})")
            
            self.current_segment.append(chunk)  # Always append to current_segment

            if speech_prob > self.vad_sensitivity:
                self.is_speaking = True
                self.silence_start_time = None
                if command_mode:
                    self.update_green_line(True)
            elif self.is_speaking and command_mode:  # Only process silence in command mode
                if self.silence_start_time is None:
                    self.silence_start_time = current_time
                elif current_time - self.silence_start_time >= self.silence_duration_threshold:
                    print("Silence detected, saving segment in command mode...")
                    self.save_segment_and_reset_vad()
            else:
                if command_mode:
                    self.update_green_line(False)
                else:
                    print("Non-command mode: continuing to record without segmenting.")

    def save_segment_and_reset_vad(self):
        print("Saving segment and resetting VAD...")
        self.save_segment()
        self.current_segment = []
        self.is_speaking = False
        self.silence_start_time = None
        self.update_green_line(False)

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
                print(f"Passing parsed text to Scratchpad and clipboard: '{parsed_text}'")
                self.root.after(0, self.update_scratchpad_text, parsed_text)
                self.root.after(0, self._update_clipboard, parsed_text)
                if command_mode:
                    print(f"Processing commands with transcription: '{parsed_text}'")
                    current_commands_snapshot = list(self.loaded_commands)
                    print(f"Loaded commands: {current_commands_snapshot}")
                    self.root.after(0, self.execute_command_from_text, parsed_text, current_commands_snapshot)
                else:
                    print("Command mode disabled, skipping command execution.")
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

    def transcription_worker(self):
        print("Transcription worker thread started.")
        while True:
            try:
                filepath = self.transcription_queue.get(timeout=1.0)
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

    def update_scratchpad_text(self, text_to_add):
        print(f"Updating Scratchpad with text: '{text_to_add}'")
        if not text_to_add.strip():
            return
        current_text = self.scratchpad.get("1.0", tk.END).strip()
        if current_text and not current_text.endswith('\n'):
            self.scratchpad.insert(tk.END, "\n")
        self.scratchpad.insert(tk.END, text_to_add + "\n")
        self.scratchpad.see(tk.END)

    def _update_clipboard(self, text):
        if not text.strip():
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            print("Clipboard updated.")
        except Exception as e:
            print(f"Error updating clipboard: {e}")

    def clear_scratchpad(self):
        self.scratchpad.delete("1.0", tk.END)

    def update_green_line(self, is_speaking):
        if is_speaking:
            if self.green_line_canvas is None:
                self.green_line_canvas = tk.Canvas(self.root, height=5)
                self.green_line_canvas.pack(fill=tk.X, padx=10, pady=2)
                self.green_line_id = self.green_line_canvas.create_rectangle(0, 0, self.root.winfo_width(), 5, fill="green")
        else:
            if self.green_line_canvas:
                self.green_line_canvas.delete(self.green_line_id)
                self.green_line_canvas.destroy()
                self.green_line_canvas = None
                self.green_line_id = None

    def destroy_green_line(self):
        if self.green_line_canvas:
            self.green_line_canvas.delete(self.green_line_id)
            self.green_line_canvas.destroy()
            self.green_line_canvas = None
            self.green_line_id = None

    def quit_app_action(self):
        print("Quit action initiated...")
        self.recording = False
        if self.audio_stream:
            self.audio_stream.stop()
            self.audio_stream.close()
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout=0.5)
        print(f"Queue size before shutdown: {self.transcription_queue.qsize()}")
        self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
        if self.transcription_worker_thread and self.transcription_worker_thread.is_alive():
            self.transcription_worker_thread.join(timeout=2.0)
        if getattr(self, 'clear_audio_on_exit', False) or getattr(self, 'clear_text_on_exit', False):
            self.delete_session_files(ask_confirm=False)
        if self.tray_icon:
            print("Stopping tray icon...")
            self.tray_icon.stop()
        print("Scheduling root quit...")
        self.save_settings()
        self.root.after(0, self.root.quit)

    def delete_session_files(self, ask_confirm=True):
        if ask_confirm:
            if not tk.messagebox.askyesno("Confirm", "Delete all session audio and text files?"):
                return
        export_dir = getattr(self, 'export_folder', '.')
        try:
            for f in os.listdir(export_dir):
                if f.endswith(('.wav', '.txt', '.transcribed')):
                    os.remove(os.path.join(export_dir, f))
                    print(f"Deleted file: {f}")
        except Exception as e:
            print(f"Error deleting session files: {e}")

if __name__ == "__main__":
    try:
        backup_script = "voice_control_app_backup.py"
        script_file = sys.argv[0] if len(sys.argv) > 0 else "voice_control_app.py"
        if os.path.exists(script_file) and script_file.endswith(".py"):
            shutil.copy(script_file, backup_script)
            print(f"Backup created: {backup_script}")
        root = tk.Tk()
        root.geometry("600x400")
        app = VoiceControlApp(root)
        root.protocol("WM_DELETE_WINDOW", app.quit_app_action)
        root.mainloop()
    except Exception as e:
        print(f"Application error: {e}")
        import traceback
        traceback.print_exc()