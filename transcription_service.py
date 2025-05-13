import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import queue
import time
import os
import shutil
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any, Tuple, Type
from abc import ABC, abstractmethod
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning
from persistent_queue_service import PersistentTaskQueue
from constants import AUDIO_QUEUE_SENTINEL, WHISPER_ENGINES
from settings_manager import AppSettings, CommandEntry, SettingsManager


class TranscriptionEngine(ABC):
    def __init__(self, settings_manager_instance: SettingsManager, root_tk_instance: tk.Tk):
        self.settings_manager = settings_manager_instance
        self.root = root_tk_instance

    @abstractmethod
    def transcribe(self, audio_path: Path, current_settings: AppSettings, prompt: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
        pass

    @abstractmethod
    def get_name(self) -> str:
        pass

class CliWhisperEngine(TranscriptionEngine):
    def __init__(self, settings_manager_instance: SettingsManager, root_tk_instance: tk.Tk):
        super().__init__(settings_manager_instance, root_tk_instance)

    def get_name(self) -> str:
        return WHISPER_ENGINES[0]

    def transcribe(self, audio_path: Path, current_settings: AppSettings, prompt: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
        whisper_exec = Path(current_settings.whisper_executable)
        if not whisper_exec.exists() or not whisper_exec.is_file():
            msg = f"Whisper executable not found or is not a file: {whisper_exec}"
            log_error(msg)
            self.root.after(0, lambda: messagebox.showerror("Whisper CLI Error", msg, parent=self.root))
            return None, msg

        export_dir = Path(current_settings.export_folder)
        export_dir.mkdir(parents=True, exist_ok=True)

        out_txt_filename = audio_path.stem + ".txt"
        out_txt_filepath = export_dir / out_txt_filename

        command = [
            str(whisper_exec),
            str(audio_path),
            "--model", current_settings.model,
            "--language", current_settings.language,
            "--output_format", "txt",
            "--output_dir", str(export_dir)
        ]
        
        if prompt:
            if '"' in prompt and "'" not in prompt:
                command.extend(["--initial_prompt", f"'{prompt}'"])
            elif "'" in prompt and '"' not in prompt:
                command.extend(["--initial_prompt", f'"{prompt}"'])
            else:
                command.extend(["--initial_prompt", prompt])

        command.extend(["--task", "translate" if current_settings.translation_enabled else "transcribe"])

        if not current_settings.whisper_cli_beeps_enabled:
            command.append("--beep_off")
        
        log_extended(f"Running Whisper CLI command: {' '.join(command)}")
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                command, check=True, capture_output=True, text=True,
                startupinfo=startupinfo, encoding='utf-8', errors='replace'
            )
            
            if out_txt_filepath.exists():
                with open(out_txt_filepath, 'r', encoding='utf-8') as f:
                    transcribed_text = f.read()
                log_essential(f"Whisper CLI transcription successful for {audio_path.name}.")
                return transcribed_text, None
            else:
                err_msg = f"Whisper CLI output file not found: {out_txt_filepath}.\n" \
                          f"STDOUT: {result.stdout[:500]}\nSTDERR: {result.stderr[:500]}"
                log_error(err_msg)
                return None, err_msg

        except subprocess.CalledProcessError as e:
            err_msg = f"Whisper CLI error. CMD: '{' '.join(e.cmd)}'.\n" \
                      f"Return Code: {e.returncode}\n" \
                      f"Stderr: {e.stderr[:1000]}"
            log_error(err_msg)
            return None, err_msg
        except FileNotFoundError:
            msg = f"Whisper CLI executable not found: {whisper_exec}"
            log_error(msg)
            self.root.after(0, lambda: messagebox.showerror("Whisper CLI Error", msg, parent=self.root))
            return None, msg
        except Exception as e_gen:
            err_msg = f"General error during Whisper CLI transcription: {e_gen}"
            log_error(err_msg, exc_info=True)
            return None, err_msg

class TranscriptionService:
    def __init__(self, settings_manager_instance: SettingsManager, root_tk_instance: tk.Tk, persistent_task_queue_ref: PersistentTaskQueue):
        self.settings_manager = settings_manager_instance
        self.settings = settings_manager_instance.settings
        self.root = root_tk_instance
        self.persistent_task_queue = persistent_task_queue_ref

        self.transcription_queue = queue.Queue()
        self._processed_in_session_cache = set()

        self._worker_thread: Optional[threading.Thread] = None
        self._stop_worker_event = threading.Event()
        self.is_queue_processing_paused: bool = False # <<< Initialized
        self._clear_queue_flag: bool = False
        self._is_transcribing_for_ui: bool = False

        # <<< CRITICAL: Define callback attributes FIRST, initializing to None
        self.on_transcription_complete: Optional[Callable[[str, Path], None]] = None
        self.on_transcription_error: Optional[Callable[[Path, str], None]] = None
        self.on_queue_updated: Optional[Callable[[int, bool], None]] = None
        self.on_transcribing_status_changed: Optional[Callable[[bool], None]] = None

        self.commands_list: List[CommandEntry] = []
        self.selected_engine: Optional[TranscriptionEngine] = None
        
        # <<< THEN, call methods that might use these attributes
        self._initialize_engine() # This is generally okay if it doesn't call callbacks
        self._load_tasks_from_persistent_queue() # Now this can safely call _notify_queue_updated

    def _load_tasks_from_persistent_queue(self):
        log_extended("Loading tasks from persistent queue into in-memory queue...")
        pending_files = self.persistent_task_queue.get_pending_tasks()
        loaded_count = 0
        for filepath_str in pending_files:
            if filepath_str not in self._processed_in_session_cache:
                self.transcription_queue.put(filepath_str)
                self._processed_in_session_cache.add(filepath_str)
                loaded_count += 1
            else:
                log_debug(f"Skipping {filepath_str} from persistent queue, already in session cache.")
        if loaded_count > 0:
            log_essential(f"Loaded {loaded_count} tasks from persistent store into the in-memory queue.")
        self._notify_queue_updated() # This is safe now


    def _initialize_engine(self):
        available_engines: Dict[str, Type[TranscriptionEngine]] = {
            WHISPER_ENGINES[0]: CliWhisperEngine,
        }
        
        chosen_engine_name = self.settings.whisper_engine_type
        engine_class = available_engines.get(chosen_engine_name)

        if engine_class:
            try:
                self.selected_engine = engine_class(self.settings_manager, self.root)
                log_essential(f"Transcription engine selected: {self.selected_engine.get_name()}")
            except Exception as e:
                log_error(f"Failed to initialize engine '{chosen_engine_name}': {e}", exc_info=True)
                if chosen_engine_name != WHISPER_ENGINES[0] and WHISPER_ENGINES[0] in available_engines:
                    try:
                        log_warning(f"Falling back to CLI engine.")
                        self.selected_engine = available_engines[WHISPER_ENGINES[0]](self.settings_manager, self.root)
                    except Exception as e_fallback:
                         log_error(f"Failed to initialize fallback CLI engine: {e_fallback}", exc_info=True)
                         self.selected_engine = None
                else:
                    self.selected_engine = None
        else:
            log_error(f"Unknown transcription engine selected: {chosen_engine_name}. No engine loaded.")
            self.selected_engine = None

        if not self.selected_engine:
            # Ensure root is available for messagebox
            if hasattr(self, 'root') and self.root:
                self.root.after(0, lambda: messagebox.showerror("Engine Error",
                                 f"Could not load transcription engine: {chosen_engine_name}.\n"
                                 "Please check configuration or try CLI engine.", parent=self.root))
            else:
                log_error("Root window not available for showing engine error messagebox.")


    def reinitialize_engine(self):
        log_essential("Re-initializing transcription engine...")
        self.settings = self.settings_manager.settings 
        self._initialize_engine()

    def set_callbacks(self, on_transcription_complete, on_transcription_error,
                      on_queue_updated, on_transcribing_status_changed):
        self.on_transcription_complete = on_transcription_complete
        self.on_transcription_error = on_transcription_error
        self.on_queue_updated = on_queue_updated
        self.on_transcribing_status_changed = on_transcribing_status_changed
        
        # Notify immediately after callbacks are set with initial state
        self._notify_queue_updated() 
        self._notify_transcribing_status(self._is_transcribing_for_ui)


    def update_commands_list(self, commands: List[CommandEntry]):
        self.commands_list = commands

    def _notify_transcribing_status(self, is_transcribing: bool):
        if self._is_transcribing_for_ui != is_transcribing:
            self._is_transcribing_for_ui = is_transcribing
            if self.on_transcribing_status_changed: # Check if callback is set
                self.root.after(0, self.on_transcribing_status_changed, is_transcribing)
    
    def _notify_queue_updated(self):
        if self.on_queue_updated: # Check if callback is set
            current_q_size = self.transcription_queue.qsize()
            is_paused = self.is_queue_processing_paused 
            self.root.after(0, self.on_queue_updated, current_q_size, is_paused)

    def add_to_queue(self, audio_filepath: str, source: str = "unknown"):
        log_extended(f"Adding to queue from source '{source}': {audio_filepath}")
        if self.persistent_task_queue.add_task(audio_filepath):
            self.transcription_queue.put(audio_filepath)
            self._processed_in_session_cache.add(audio_filepath)
            log_extended(f"Task '{audio_filepath}' added to both persistent and in-memory queues.")
        else:
            log_error(f"Failed to add task '{audio_filepath}' to persistent queue. Not adding to in-memory queue.")
        
        self._notify_queue_updated()


    def start_worker(self):
        if self._worker_thread and self._worker_thread.is_alive():
            log_extended("Transcription worker already running.")
            return
        
        self._stop_worker_event.clear()
        self._worker_thread = threading.Thread(target=self._transcription_worker_loop, daemon=True)
        self._worker_thread.start()
        log_essential("Transcription worker thread started.")

    def stop_worker(self):
        if not self._worker_thread or not self._worker_thread.is_alive():
            log_extended("Transcription worker not running or already stopped.")
            return

        log_essential("Stopping transcription worker...")
        self._stop_worker_event.set()
        try:
            self.transcription_queue.put(AUDIO_QUEUE_SENTINEL, block=False) # Try non-blocking first
        except queue.Full:
            log_warning("Transcription queue full when trying to put SENTINEL. Worker might be stuck.")
            # If worker is stuck, join might timeout. Sentinel helps if it's just waiting on queue.get()
        
        self._worker_thread.join(timeout=3.0)
        if self._worker_thread.is_alive():
            log_error("Transcription worker thread did not stop cleanly.")
        self._worker_thread = None
        self._notify_transcribing_status(False)
        log_essential("Transcription worker stopped.")


    def toggle_pause_queue(self):
        self.is_queue_processing_paused = not self.is_queue_processing_paused
        state = "Paused" if self.is_queue_processing_paused else "Resumed"
        log_essential(f"Transcription queue processing {state}.")
        if not self.is_queue_processing_paused:
            self._check_and_load_new_persistent_tasks()
        self._notify_queue_updated()


    def clear_queue(self):
        log_essential("Clearing transcription queue (both in-memory and persistent)...")
        
        self._clear_queue_flag = True 
        
        cleared_in_memory_count = 0
        while not self.transcription_queue.empty():
            try:
                item = self.transcription_queue.get_nowait()
                if item is AUDIO_QUEUE_SENTINEL:
                    self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
                else:
                    cleared_in_memory_count += 1
                self.transcription_queue.task_done()
            except queue.Empty:
                break
        log_extended(f"Cleared {cleared_in_memory_count} items from in-memory queue.")

        if self.persistent_task_queue.clear_all_tasks():
            log_extended("Successfully cleared persistent task queue.")
        else:
            log_error("Failed to clear persistent task queue. Some tasks may remain on disk.")

        self._processed_in_session_cache.clear()
        log_essential(f"Clear queue request processed. In-memory items cleared: {cleared_in_memory_count}.")
        self._notify_queue_updated()
        if self.transcription_queue.qsize() == 0:
            self._notify_transcribing_status(False)


    def _check_and_load_new_persistent_tasks(self):
        log_debug("Checking persistent queue for new tasks...")
        persistent_tasks = self.persistent_task_queue.get_pending_tasks()
        new_tasks_loaded = 0
        for task_path_str in persistent_tasks:
            if task_path_str not in self._processed_in_session_cache:
                self.transcription_queue.put(task_path_str)
                self._processed_in_session_cache.add(task_path_str)
                new_tasks_loaded += 1
                log_extended(f"Loaded new task from persistent store: {task_path_str}")
        
        if new_tasks_loaded > 0:
            log_essential(f"Dynamically loaded {new_tasks_loaded} new tasks from persistent queue.")
            self._notify_queue_updated()


    def _transcription_worker_loop(self):
        engine_name_for_log = self.selected_engine.get_name() if self.selected_engine else 'None'
        log_debug(f"Transcription worker loop started. Engine: {engine_name_for_log}")

        while not self._stop_worker_event.is_set():
            try:
                if not self.selected_engine:
                    log_warning("No transcription engine loaded. Worker pausing.")
                    if self._is_transcribing_for_ui: self._notify_transcribing_status(False)
                    time.sleep(1)
                    continue

                if self.is_queue_processing_paused:
                    if self._is_transcribing_for_ui:
                        self._notify_transcribing_status(False)
                    log_debug("Transcription worker: Queue processing is paused. Sleeping...")
                    time.sleep(0.5)
                    continue

                if self._stop_worker_event.is_set():
                    log_debug("Transcription worker: Stop event detected before getting from queue.")
                    break

                log_debug("Transcription worker: Attempting to get from queue...")
                try:
                    audio_filepath_str = self.transcription_queue.get(timeout=0.5)
                except queue.Empty:
                    if self._is_transcribing_for_ui: self._notify_transcribing_status(False)
                    self._check_and_load_new_persistent_tasks()
                    continue

                if audio_filepath_str is AUDIO_QUEUE_SENTINEL:
                    self.transcription_queue.task_done()
                    break

                if self._clear_queue_flag:
                    log_extended(f"Worker skipping '{audio_filepath_str}' due to clear_queue_flag.")
                    self.transcription_queue.task_done()
                    if self.transcription_queue.empty():
                        log_extended("In-memory queue empty after skipping due to clear_queue_flag, resetting flag.")
                        self._clear_queue_flag = False 
                    self._notify_queue_updated()
                    continue
                
                audio_filepath = Path(audio_filepath_str)
                self._processed_in_session_cache.add(audio_filepath_str)

                if not audio_filepath.exists() or audio_filepath.name.endswith(".transcribed"):
                    log_warning(f"Worker skipping non-existent or already processed file: {audio_filepath}")
                    self.persistent_task_queue.mark_task_complete(audio_filepath_str)
                    self.transcription_queue.task_done()
                    self._notify_queue_updated()
                    continue

                self._notify_transcribing_status(True)
                self._notify_queue_updated()
                log_essential(f"Worker processing: {audio_filepath}")

                transcribed_text = None
                error_msg = None
                
                current_app_settings = self.settings_manager.settings 
                current_prompt = self.settings_manager.prompt
                transcribed_text, error_msg = self.selected_engine.transcribe(
                    audio_filepath, current_app_settings, current_prompt
                )

                if transcribed_text is not None:
                    parsed_text = self._parse_and_clean_transcription_text(transcribed_text)
                    
                    if self.settings_manager.settings.auto_add_space and parsed_text and not parsed_text.endswith(' '):
                        parsed_text += ' '
                        log_debug("Auto-added space to transcription.")

                    if self.on_transcription_complete:
                        self.root.after(0, self.on_transcription_complete, parsed_text, audio_filepath)
                    
                    if not self.persistent_task_queue.mark_task_complete(audio_filepath_str):
                        log_warning(f"Could not mark '{audio_filepath_str}' as complete in persistent queue (it might have been removed already).")
                    
                    try:
                        new_audio_path = audio_filepath.with_suffix(audio_filepath.suffix + ".transcribed")
                        shutil.move(str(audio_filepath), str(new_audio_path))
                        log_extended(f"Renamed processed audio to: {new_audio_path.name}")
                    except Exception as e_mv:
                        log_error(f"Error renaming audio file {audio_filepath.name}: {e_mv}")
                else:
                    if self.on_transcription_error:
                        self.root.after(0, self.on_transcription_error, audio_filepath, error_msg or "Unknown transcription error.")
                
                self.transcription_queue.task_done()
                self._notify_queue_updated()

            except Exception as e:
                log_path_str = 'unknown file'
                # Check if audio_filepath_str is defined and not None or SENTINEL before using
                if 'audio_filepath_str' in locals() and \
                   audio_filepath_str is not None and \
                   audio_filepath_str is not AUDIO_QUEUE_SENTINEL:
                    log_path_str = str(audio_filepath_str)
                
                log_error(f"Critical error in transcription worker for {log_path_str}: {e}", exc_info=True)
                
                try:
                    # Check if audio_filepath_str is defined before trying task_done
                    if 'audio_filepath_str' in locals() and audio_filepath_str is not None:
                         self.transcription_queue.task_done()
                except ValueError:
                    pass
                except queue.Empty:
                    pass
                self._notify_transcribing_status(False)
                time.sleep(1)

        log_essential("Transcription worker loop has exited.")
        self._notify_transcribing_status(False)


    def _parse_and_clean_transcription_text(self, raw_text: str) -> str:
        if not raw_text: return ""
        current_settings = self.settings_manager.settings
        if not current_settings.clear_text_output and not current_settings.timestamps_disabled:
            return raw_text.strip()

        import re
        cleaned_lines = []
        ts_pattern_general = re.compile(r'^\[\s*[\d:.]+\s*-->\s*[\d:.]+\s*\]\s*')

        for line in raw_text.splitlines():
            line_stripped = line.strip()
            if not line_stripped:
                continue

            if current_settings.clear_text_output:
                if line_stripped.startswith("---") or \
                   line_stripped.startswith("===") or \
                   re.match(r'^\d{4}-\d{2}-\d{2}', line_stripped):
                    continue
            
            match = ts_pattern_general.match(line_stripped)
            if match:
                if current_settings.timestamps_disabled:
                    text_part = line_stripped[match.end():].strip()
                    if text_part:
                        cleaned_lines.append(text_part)
                else:
                    cleaned_lines.append(line_stripped)
            else:
                cleaned_lines.append(line_stripped)
        
        processed_lines = [re.sub(r'\s+', ' ', line).strip() for line in cleaned_lines]
        processed_lines = [line for line in processed_lines if line]

        if current_settings.timestamps_disabled:
            final_text = " ".join(processed_lines)
        else:
            final_text = "\n".join(processed_lines)
            
        return final_text.strip()

    def execute_command_from_text(self, transcription_text: str):
        if not transcription_text.strip() or not self.commands_list:
            return
        
        import re
        
        cleaned_transcription = transcription_text.lower().strip()
        log_extended(f"Attempting to match command in: '{cleaned_transcription}'")

        for cmd_entry in self.commands_list:
            voice_trigger = cmd_entry.voice.strip().lower()
            action_to_run = cmd_entry.action.strip()

            if not voice_trigger or not action_to_run:
                continue
            
            pattern_str = voice_trigger
            is_wildcard_command = " ff " in pattern_str
            
            if is_wildcard_command:
                pattern_str = pattern_str.replace(" ff ", r"%%%%WILDCARD%%%%")

            pattern_str = re.escape(pattern_str)

            if is_wildcard_command:
                pattern_str = pattern_str.replace(re.escape("%%%%WILDCARD%%%%"), r"(.*?)")
            
            prefix = r'\b' if pattern_str and pattern_str[0].isalnum() else ''
            suffix = r'\b' if pattern_str and pattern_str[-1].isalnum() else ''
            final_pattern = prefix + pattern_str + suffix

            try:
                match = re.search(final_pattern, cleaned_transcription, re.IGNORECASE)
                if match:
                    matched_action = action_to_run
                    if is_wildcard_command and match.groups():
                        wildcard_content = match.group(1).strip()
                        matched_action = re.sub(r'\bFF\b', wildcard_content, action_to_run, flags=re.IGNORECASE)
                    
                    log_essential(f"Command matched: '{voice_trigger}' -> Action: '{matched_action}'")
                    self._run_subprocess_action(matched_action)
                    return

            except re.error as re_err:
                log_error(f"Regex error for command '{voice_trigger}' (pattern: {final_pattern}): {re_err}")
                continue
        log_extended("No command matched.")


    def _run_subprocess_action(self, action_string: str):
        log_essential(f"Running subprocess action: '{action_string}'")
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            subprocess.run(action_string, shell=True, check=False,
                           capture_output=True, text=True,
                           startupinfo=startupinfo, encoding='utf-8', errors='replace')
            log_extended(f"Subprocess action '{action_string}' completed.")
        except Exception as e:
            log_error(f"Error running subprocess action '{action_string}': {e}", exc_info=True)