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
    # ... (No changes here) ...
    def __init__(self, settings_manager_instance: SettingsManager, root_tk_instance: tk.Tk):
        self.settings_manager = settings_manager_instance
        self.root = root_tk_instance

    @abstractmethod
    def transcribe(self, audio_path: Path, current_settings: AppSettings, prompt: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
        pass

    @abstractmethod
    def get_name(self) -> str:
        pass

    @abstractmethod
    def prime_model(self, language: str, model_name: str, test_audio_path: Path, priming_output_dir: Path) -> Tuple[bool, str]:
        pass

class CliWhisperEngine(TranscriptionEngine):
    # ... (No changes here, it receives paths from TranscriptionService) ...
    def __init__(self, settings_manager_instance: SettingsManager, root_tk_instance: tk.Tk):
        super().__init__(settings_manager_instance, root_tk_instance)

    def get_name(self) -> str:
        return WHISPER_ENGINES[0] 

    def _build_cli_command(self,
                           whisper_exec: Path,
                           audio_path: Path,
                           model_name: str,
                           language: str,
                           output_dir: Path,
                           prompt: Optional[str] = None,
                           task: str = "transcribe", 
                           is_priming: bool = False,
                           current_settings: Optional[AppSettings] = None
                           ) -> List[str]:
        
        command = [
            str(whisper_exec),
            str(audio_path),
            "--model", model_name,
            "--language", language if language else "auto",
            "--output_dir", str(output_dir)
        ]

        if is_priming:
            command.extend(["--output_format", "txt"])
            command.append("--verbose") 
            command.append("False")    
            if current_settings and not current_settings.whisper_cli_beeps_enabled:
                 command.append("--beep_off")
            elif not current_settings: 
                 command.append("--beep_off")
        else: 
            command.extend(["--output_format", "txt"]) 
            if prompt:
                if '"' in prompt and "'" not in prompt:
                    command.extend(["--initial_prompt", f"'{prompt}'"])
                elif "'" in prompt and '"' not in prompt:
                    command.extend(["--initial_prompt", f'"{prompt}"'])
                else: 
                    command.extend(["--initial_prompt", prompt])
            
            command.extend(["--task", task])

            if current_settings: 
                if not current_settings.whisper_cli_beeps_enabled:
                    command.append("--beep_off")
            else:
                log_warning("CliWhisperEngine._build_cli_command called for non-priming without current_settings!")

        log_debug(f"Built Whisper CLI command ({'priming' if is_priming else 'transcribe'}): {' '.join(command)}")
        return command

    def transcribe(self, audio_path: Path, current_settings: AppSettings, prompt: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
        whisper_exec_str = current_settings.whisper_executable
        if not whisper_exec_str:
            msg = "Whisper executable path is not configured."
            log_error(msg)
            if self.root and self.root.winfo_exists(): 
                self.root.after(0, lambda: messagebox.showerror("Whisper CLI Error", msg, parent=self.root))
            return None, msg
            
        whisper_exec = Path(whisper_exec_str)
        if not whisper_exec.is_file(): 
            msg = f"Whisper executable not found or is not a file: {whisper_exec}"
            log_error(msg)
            if self.root and self.root.winfo_exists():
                self.root.after(0, lambda: messagebox.showerror("Whisper CLI Error", msg, parent=self.root))
            return None, msg

        export_dir = Path(current_settings.export_folder)
        transcription_output_dir = export_dir / audio_path.stem 
        try:
            transcription_output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            msg = f"Failed to create output directory {transcription_output_dir}: {e}"
            log_error(msg, exc_info=True)
            return None, msg

        out_txt_filename = audio_path.stem + ".txt" 
        out_txt_filepath = transcription_output_dir / out_txt_filename

        command = self._build_cli_command(
            whisper_exec=whisper_exec,
            audio_path=audio_path,
            model_name=current_settings.model,
            language=current_settings.language,
            output_dir=transcription_output_dir, 
            prompt=prompt,
            task="translate" if current_settings.translation_enabled else "transcribe",
            is_priming=False,
            current_settings=current_settings
        )
        
        log_extended(f"Running Whisper CLI command: {' '.join(command)}")
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                command, check=True, capture_output=True, text=True,
                startupinfo=startupinfo, encoding='utf-8', errors='replace', timeout=600 
            )
            
            if out_txt_filepath.exists():
                with open(out_txt_filepath, 'r', encoding='utf-8') as f:
                    transcribed_text = f.read().strip() 
                log_essential(f"Whisper CLI transcription successful for {audio_path.name}.")
                return transcribed_text, None 
            else:
                other_outputs = list(transcription_output_dir.glob(f"{audio_path.stem}.*"))
                if other_outputs:
                    log_warning(f"Whisper CLI produced output files ({[f.name for f in other_outputs]}), but expected '{out_txt_filepath.name}' was not found.")
                
                err_msg = f"Whisper CLI output file '{out_txt_filepath.name}' not found in '{transcription_output_dir}'.\n" \
                          f"STDOUT: {result.stdout[:500]}\nSTDERR: {result.stderr[:500]}"
                log_error(err_msg)
                return None, err_msg

        except subprocess.CalledProcessError as e:
            err_msg = f"Whisper CLI error. CMD: '{' '.join(e.cmd)}'.\n" \
                      f"Return Code: {e.returncode}\n" \
                      f"Stdout: {e.stdout[:500]}\nStderr: {e.stderr[:500]}" 
            log_error(err_msg)
            return None, err_msg
        except subprocess.TimeoutExpired:
            err_msg = f"Whisper CLI process timed out for {audio_path.name}."
            log_error(err_msg)
            return None, err_msg
        except FileNotFoundError: 
            msg = f"Whisper CLI executable not found during run: {whisper_exec}"
            log_error(msg)
            return None, msg
        except Exception as e_gen:
            err_msg = f"General error during Whisper CLI transcription: {e_gen}"
            log_error(err_msg, exc_info=True)
            return None, err_msg

    def prime_model(self, language: str, model_name: str, test_audio_path: Path, priming_output_dir: Path) -> Tuple[bool, str]:
        log_essential(f"CLI Engine: Priming model '{model_name}' for language '{language}' using '{test_audio_path.name}'.")
        
        current_app_settings = self.settings_manager.settings 
        whisper_exec_str = current_app_settings.whisper_executable
        if not whisper_exec_str:
            msg = "Whisper executable path is not configured for priming."
            log_error(msg)
            return False, msg
            
        whisper_exec = Path(whisper_exec_str)
        if not whisper_exec.is_file():
            msg = f"Whisper executable not found or is not a file for priming: {whisper_exec}"
            log_error(msg)
            return False, msg

        try:
            priming_output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            msg = f"Failed to create priming output directory {priming_output_dir}: {e}"
            log_error(msg, exc_info=True)
            return False, msg

        command = self._build_cli_command(
            whisper_exec=whisper_exec,
            audio_path=test_audio_path,
            model_name=model_name,
            language=language,
            output_dir=priming_output_dir,
            is_priming=True,
            current_settings=current_app_settings
        )
        
        log_extended(f"Running Whisper CLI priming command: {' '.join(command)}")
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                command, check=True, capture_output=True, text=True,
                startupinfo=startupinfo, encoding='utf-8', errors='replace', timeout=600 
            )
            log_essential(f"Whisper CLI priming for model '{model_name}' (lang: {language}) completed successfully.")
            priming_outputs = list(priming_output_dir.glob(f"{test_audio_path.stem}.*"))
            if priming_outputs:
                log_debug(f"Priming created output files: {[f.name for f in priming_outputs]}")
            else:
                log_warning(f"Priming completed but no output files found in {priming_output_dir}. This might be okay if model was already cached.")

            return True, f"Model '{model_name}' (lang: {language}) primed/checked successfully."

        except subprocess.CalledProcessError as e:
            err_msg = f"Whisper CLI priming error for model '{model_name}'. CMD: '{' '.join(e.cmd)}'.\n" \
                      f"Return Code: {e.returncode}\n" \
                      f"Stdout: {e.stdout[:500]}\nStderr: {e.stderr[:500]}"
            log_error(err_msg)
            user_friendly_error = e.stderr.splitlines()[-1] if e.stderr else 'CLI error during priming'
            return False, f"Failed to prime model '{model_name}'. Error: {user_friendly_error}"
        except subprocess.TimeoutExpired:
            err_msg = f"Whisper CLI priming process timed out for model '{model_name}' (likely during download)."
            log_error(err_msg)
            return False, err_msg
        except FileNotFoundError:
            msg = f"Whisper CLI executable not found during priming: {whisper_exec}"
            log_error(msg)
            return False, msg
        except Exception as e_gen:
            err_msg = f"General error during Whisper CLI priming for model '{model_name}': {e_gen}"
            log_error(err_msg, exc_info=True)
            return False, err_msg


class TranscriptionService:
    def __init__(self, 
                 settings_manager_instance: SettingsManager, 
                 root_tk_instance: tk.Tk, 
                 persistent_task_queue_ref: PersistentTaskQueue,
                 app_asset_path_ref: Path): # <<< NEW: Added app_asset_path_ref
        self.settings_manager = settings_manager_instance
        self.settings = settings_manager_instance.settings 
        self.root = root_tk_instance
        self.persistent_task_queue = persistent_task_queue_ref
        self.app_asset_path = app_asset_path_ref # <<< NEW: Store it

        self.transcription_queue = queue.Queue()
        self._processed_in_session_cache = set() 

        self._worker_thread: Optional[threading.Thread] = None
        self._stop_worker_event = threading.Event()
        self.is_queue_processing_paused: bool = False
        self._clear_queue_flag: bool = False 
        self._is_transcribing_for_ui: bool = False 

        self.on_transcription_complete: Optional[Callable[[str, Path], None]] = None
        self.on_transcription_error: Optional[Callable[[Path, str], None]] = None
        self.on_queue_updated: Optional[Callable[[int, bool], None]] = None 
        self.on_transcribing_status_changed: Optional[Callable[[bool], None]] = None

        self.commands_list: List[CommandEntry] = [] 
        self.selected_engine: Optional[TranscriptionEngine] = None
        
        self._initialize_engine()
        self._load_tasks_from_persistent_queue()

        # MODIFIED: Use self.app_asset_path to define test_audio_dir and test_audio_file
        self.test_audio_dir = self.app_asset_path / "TestAudio" 
        self.test_audio_file = self.test_audio_dir / "test.wav"
        self._ensure_test_audio_file() 
        
        self.priming_thread: Optional[threading.Thread] = None
        self.priming_stop_event = threading.Event()


    def _ensure_test_audio_file(self):
        try:
            # No need to create self.test_audio_dir separately if self.app_asset_path points to root of assets
            # PyInstaller will place TestAudio inside the app_asset_path if spec is correct
            if not self.test_audio_dir.exists():
                # This case should ideally not happen if PyInstaller is configured correctly
                # and get_app_asset_path() correctly points to where TestAudio *should* be.
                # If it does happen, creating it at runtime in a bundled app might be tricky
                # regarding write permissions or finding the right relative spot.
                log_error(f"Critical: Bundled 'TestAudio' directory not found at expected location: {self.test_audio_dir}. Priming will fail.")
                # Fallback: Try to create it in user_config_path as a last resort for the dummy file.
                # This is not ideal as the "official" test.wav wouldn't be there.
                user_data_test_audio_dir = Path(self.settings_manager.user_config_dir) / "TestAudio_runtime"
                user_data_test_audio_dir.mkdir(parents=True, exist_ok=True)
                self.test_audio_file = user_data_test_audio_dir / "test_dummy.wav"
                log_warning(f"Will attempt to create dummy test WAV in user data: {self.test_audio_file}")


            if not self.test_audio_file.exists() or self.test_audio_file.stat().st_size < 44: 
                log_warning(f"'{self.test_audio_file}' not found or is too small. Creating a dummy silent WAV for model priming.")
                # Ensure the directory for self.test_audio_file exists before writing
                self.test_audio_file.parent.mkdir(parents=True, exist_ok=True)
                try:
                    import wave
                    nchannels = 1; sampwidth = 2; framerate = 16000
                    nframes = framerate // 20 
                    comptype = "NONE"; compname = "not compressed"
                    with wave.open(str(self.test_audio_file), 'wb') as wf:
                        wf.setnchannels(nchannels)
                        wf.setsampwidth(sampwidth)
                        wf.setframerate(framerate)
                        wf.setnframes(nframes)
                        wf.setcomptype(comptype, compname)
                        wf.writeframes(b'\x00' * (nframes * sampwidth * nchannels))
                    log_debug(f"Dummy silent WAV file created/overwritten at '{self.test_audio_file}'")
                except ImportError:
                    log_error("Could not import 'wave' module. Cannot create dummy 'test.wav'. Model priming will fail if file is missing and cannot be created.")
                except Exception as e_wave:
                    log_error(f"Failed to create dummy silent WAV '{self.test_audio_file}': {e_wave}", exc_info=True)
            else:
                log_debug(f"Found existing valid test audio file: {self.test_audio_file}")

        except Exception as e_dir: # Catch errors related to directory operations primarily
            log_error(f"Error ensuring TestAudio setup using base '{self.app_asset_path}': {e_dir}", exc_info=True)


    # ... (All other methods of TranscriptionService from your last complete version,
    #      including _load_tasks_from_persistent_queue, _initialize_engine, reinitialize_engine,
    #      set_callbacks, update_commands_list, _notify_*, add_to_queue, start_worker,
    #      stop_worker, toggle_pause_queue, clear_queue, _check_and_load_new_persistent_tasks,
    #      _transcription_worker_loop, _parse_and_clean_transcription_text,
    #      execute_command_from_text, _run_subprocess_action,
    #      _prime_model_worker, and prime_engine_model remain UNCHANGED
    #      from the version where the `log_info` was fixed.)
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
        self._notify_queue_updated()


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
            if hasattr(self, 'root') and self.root and self.root.winfo_exists(): 
                self.root.after(0, lambda: messagebox.showerror("Engine Error",
                                 f"Could not load transcription engine: {chosen_engine_name}.\n"
                                 "Please check configuration or ensure selected engine is functional.", parent=self.root))
            else:
                log_error("Root window not available for showing engine error messagebox during init.")


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
        
        self._notify_queue_updated() 
        self._notify_transcribing_status(self._is_transcribing_for_ui)


    def update_commands_list(self, commands: List[CommandEntry]): 
        self.commands_list = commands


    def _notify_transcribing_status(self, is_transcribing: bool): 
        if self._is_transcribing_for_ui != is_transcribing:
            self._is_transcribing_for_ui = is_transcribing
            if self.on_transcribing_status_changed and self.root.winfo_exists():
                self.root.after(0, self.on_transcribing_status_changed, is_transcribing)
    

    def _notify_queue_updated(self): 
        if self.on_queue_updated and self.root.winfo_exists(): 
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
            self.transcription_queue.put(AUDIO_QUEUE_SENTINEL, block=False) 
        except queue.Full:
            log_warning("Transcription queue full when trying to put SENTINEL. Worker might be stuck.")
        
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
                    pass
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
        engine_name_for_log = self.selected_engine.get_name() if self.selected_engine else 'None (Error in init?)'
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
                    time.sleep(0.5) 
                    continue

                if self._stop_worker_event.is_set():
                    log_debug("Transcription worker: Stop event detected before getting from queue.")
                    break

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
                
                if not audio_filepath.exists():
                    log_warning(f"Worker skipping non-existent file: {audio_filepath}")
                    self.persistent_task_queue.mark_task_complete(audio_filepath_str) 
                    self.transcription_queue.task_done()
                    self._notify_queue_updated()
                    continue
                
                if audio_filepath.name.endswith(".transcribed"): 
                    log_debug(f"Worker skipping already processed file (by '.transcribed' suffix): {audio_filepath.name}")
                    self.persistent_task_queue.mark_task_complete(audio_filepath_str)
                    self.transcription_queue.task_done()
                    self._notify_queue_updated()
                    continue


                self._notify_transcribing_status(True)
                self._notify_queue_updated() 
                log_essential(f"Worker processing: {audio_filepath.name} with engine {self.selected_engine.get_name()}")

                transcribed_text = None
                error_msg = None
                
                current_app_settings = self.settings_manager.settings 
                current_prompt = self.settings_manager.prompt
                transcribed_text, error_msg = self.selected_engine.transcribe(
                    audio_filepath, current_app_settings, current_prompt
                )

                if transcribed_text is not None: 
                    parsed_text = self._parse_and_clean_transcription_text(transcribed_text, current_app_settings)
                    
                    if current_app_settings.auto_add_space and parsed_text and not parsed_text.endswith(' '):
                        parsed_text += ' '
                        log_debug("Auto-added space to transcription.")

                    if self.on_transcription_complete and self.root.winfo_exists():
                        self.root.after(0, self.on_transcription_complete, parsed_text, audio_filepath)
                    
                    if not self.persistent_task_queue.mark_task_complete(audio_filepath_str):
                        log_warning(f"Could not mark '{audio_filepath_str}' as complete in persistent queue (it might have been removed already).")
                    
                    try:
                        new_audio_path = audio_filepath.with_suffix(audio_filepath.suffix + ".transcribed")
                        if new_audio_path.exists(): new_audio_path.unlink(missing_ok=True) 
                        shutil.move(str(audio_filepath), str(new_audio_path))
                        log_extended(f"Renamed processed audio to: {new_audio_path.name}")
                    except Exception as e_mv:
                        log_error(f"Error renaming audio file {audio_filepath.name} to {new_audio_path.name if 'new_audio_path' in locals() else 'unknown'}: {e_mv}")
                else: 
                    if self.on_transcription_error and self.root.winfo_exists():
                        self.root.after(0, self.on_transcription_error, audio_filepath, error_msg or "Unknown transcription error.")
                
                self.transcription_queue.task_done() 
                self._notify_queue_updated() 

            except Exception as e: 
                log_path_str = 'unknown file'
                if 'audio_filepath_str' in locals() and \
                   audio_filepath_str is not None and \
                   audio_filepath_str is not AUDIO_QUEUE_SENTINEL:
                    log_path_str = str(audio_filepath_str)
                
                log_error(f"Critical error in transcription worker for {log_path_str}: {e}", exc_info=True)
                
                try: 
                    if 'audio_filepath_str' in locals() and audio_filepath_str is not None:
                         self.transcription_queue.task_done()
                except ValueError: pass 
                except queue.Empty: pass 
                
                self._notify_transcribing_status(False) 
                time.sleep(1) 

        log_essential("Transcription worker loop has exited.")
        self._notify_transcribing_status(False)


    def _parse_and_clean_transcription_text(self, raw_text: str, current_settings: AppSettings) -> str: 
        if not raw_text: return ""
        if not current_settings.clear_text_output and not current_settings.timestamps_disabled:
            return raw_text.strip()

        import re
        cleaned_lines = []
        ts_pattern_general = re.compile(r'^\[\s*([\d:.]+)\s*-->\s*([\d:.]+)\s*\]\s*')

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
            log_extended(f"Subprocess action '{action_string}' completed (or initiated if non-blocking).")
        except Exception as e:
            log_error(f"Error running subprocess action '{action_string}': {e}", exc_info=True)

    # --- NEW Model Priming Functionality ---
    def _prime_model_worker(self, language: str, model_name: str,
                            callback: Optional[Callable[[bool, str], None]]):
        if not self.selected_engine or not hasattr(self.selected_engine, 'prime_model'): 
            err_msg = f"Priming failed: Selected engine does not support priming or not available. Current: {self.selected_engine.get_name() if self.selected_engine else 'None'}"
            log_error(err_msg)
            if callback and self.root.winfo_exists(): self.root.after(0, callback, False, err_msg)
            return

        log_essential(f"Starting model priming via '{self.selected_engine.get_name()}' for language='{language}', model='{model_name}'...")
        
        if not self.test_audio_file.exists():
            err_msg = f"Test audio file '{self.test_audio_file}' not found. Cannot prime model."
            log_error(err_msg)
            if callback and self.root.winfo_exists(): self.root.after(0, callback, False, err_msg)
            return

        priming_base_dir = Path(self.settings_manager.settings.export_folder) / "priming_temp"
        timestamp = time.strftime("%Y%m%d%H%M%S")
        safe_model_name = model_name.replace('/', '_').replace('\\', '_')
        priming_specific_output_dir = priming_base_dir / f"{safe_model_name}_{language}_{timestamp}"

        try:
            priming_specific_output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            err_msg = f"Failed to create priming output directory '{priming_specific_output_dir}': {e}"
            log_error(err_msg, exc_info=True)
            if callback and self.root.winfo_exists(): self.root.after(0, callback, False, err_msg)
            return
        
        success, message = self.selected_engine.prime_model(
            language=language,
            model_name=model_name,
            test_audio_path=self.test_audio_file,
            priming_output_dir=priming_specific_output_dir
        )

        if callback and self.root.winfo_exists():
            self.root.after(0, callback, success, message)
        
        try:
            if priming_specific_output_dir.exists():
                shutil.rmtree(priming_specific_output_dir)
                log_debug(f"Cleaned up priming output directory: {priming_specific_output_dir}")
        except Exception as e_cleanup:
            log_warning(f"Could not clean up priming output directory '{priming_specific_output_dir}': {e_cleanup}")

    def prime_engine_model(self, language: str, model_name: str, 
                           callback: Optional[Callable[[bool, str], None]] = None):
        if not self.selected_engine or not hasattr(self.selected_engine, 'prime_model'):
            msg = f"Model priming skipped: Selected engine ({self.selected_engine.get_name() if self.selected_engine else 'None'}) does not support priming."
            log_warning(msg)
            if callback and self.root.winfo_exists(): self.root.after(0, callback, True, "Priming not applicable for current engine type.")
            return

        if self.priming_thread and self.priming_thread.is_alive():
            msg = f"Model priming for '{model_name}' requested, but another priming is active."
            log_warning(msg)
            if callback and self.root.winfo_exists(): self.root.after(0, callback, False, "Another priming process is active.")
            return

        log_debug(f"Queueing model priming for engine '{self.selected_engine.get_name()}': Lang='{language}', Model='{model_name}'")

        self.priming_thread = threading.Thread(
            target=self._prime_model_worker,
            args=(language, model_name, callback),
            daemon=True
        )
        self.priming_thread.start()