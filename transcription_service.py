import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import queue
import time
import os
import shutil
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any, Tuple # <--- ENSURE Tuple IS HERE
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning
# WHISPER_ENGINES will likely be simplified or removed from constants later
from constants import AUDIO_QUEUE_SENTINEL, WHISPER_ENGINES 
from settings_manager import AppSettings, CommandEntry # For type hinting
# from whisper_lib_integration import FasterWhisperLib, FASTER_WHISPER_AVAILABLE # REMOVE

class TranscriptionService:
    def __init__(self, settings_ref: AppSettings, root_tk_instance: tk.Tk, settings_manager=None):
        self.settings = settings_ref
        self.settings_manager = settings_manager
        self.root = root_tk_instance # For UI updates, messagebox

        self.transcription_queue = queue.Queue()
        self._worker_thread: Optional[threading.Thread] = None
        self._stop_worker_event = threading.Event()
        self.is_queue_processing_paused: bool = False
        self._clear_queue_flag: bool = False # To signal worker to discard items
        self._is_transcribing_for_ui: bool = False # To update UI status

        # Callbacks to be set by main app
        self.on_transcription_complete: Optional[Callable[[str, Path], None]] = None # (text, original_audio_path)
        self.on_transcription_error: Optional[Callable[[Path, str], None]] = None # (audio_path, error_message)
        self.on_queue_updated: Optional[Callable[[int], None]] = None # (queue_size)
        self.on_transcribing_status_changed: Optional[Callable[[bool], None]] = None # (is_transcribing)

        # Command execution related
        self.commands_list: List[CommandEntry] = [] # To be updated by main app

        # Whisper lib integration - REMOVED
        # self.faster_whisper_instance: Optional[FasterWhisperLib] = None
        # if FASTER_WHISPER_AVAILABLE:
        #     self.faster_whisper_instance = FasterWhisperLib(self.settings)

    def set_callbacks(self, on_transcription_complete, on_transcription_error,
                      on_queue_updated, on_transcribing_status_changed):
        self.on_transcription_complete = on_transcription_complete
        self.on_transcription_error = on_transcription_error
        self.on_queue_updated = on_queue_updated
        self.on_transcribing_status_changed = on_transcribing_status_changed

    def update_commands_list(self, commands: List[CommandEntry]):
        self.commands_list = commands

    def _notify_transcribing_status(self, is_transcribing: bool):
        if self._is_transcribing_for_ui != is_transcribing:
            self._is_transcribing_for_ui = is_transcribing
            if self.on_transcribing_status_changed:
                self.root.after(0, self.on_transcribing_status_changed, is_transcribing)
    
    def _notify_queue_updated(self):
        if self.on_queue_updated:
            self.root.after(0, self.on_queue_updated, self.transcription_queue.qsize())

    def add_to_queue(self, audio_filepath: str):
        self.transcription_queue.put(audio_filepath)
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
        self.transcription_queue.put(AUDIO_QUEUE_SENTINEL) # Ensure worker wakes up
        self._worker_thread.join(timeout=3.0) # Wait for graceful shutdown
        if self._worker_thread.is_alive():
            log_error("Transcription worker thread did not stop cleanly.")
        self._worker_thread = None
        self._notify_transcribing_status(False) # Ensure UI reflects stop
        log_essential("Transcription worker stopped.")


    def toggle_pause_queue(self):
        self.is_queue_processing_paused = not self.is_queue_processing_paused
        state = "Paused" if self.is_queue_processing_paused else "Resumed"
        log_essential(f"Transcription queue processing {state}.")
        if not self.is_queue_processing_paused and not self.transcription_queue.empty():
            # If resuming and queue has items, ensure worker is active
            # This is mostly for UI feedback, worker loop handles the pause internally.
            pass

    def clear_queue(self):
        log_essential("Clearing transcription queue...")
        self._clear_queue_flag = True
        cleared_count = 0
        # Drain the queue quickly. Worker will also skip items due to the flag.
        while not self.transcription_queue.empty():
            try:
                item = self.transcription_queue.get_nowait()
                if item is AUDIO_QUEUE_SENTINEL: # Put sentinel back if found
                    self.transcription_queue.put(AUDIO_QUEUE_SENTINEL)
                else:
                    cleared_count += 1
                self.transcription_queue.task_done()
            except queue.Empty:
                break
        # Flag will be reset by worker after it processes its current item (if any)
        # Or, if we want immediate effect, we'd need a lock around queue access in worker.
        # For now, this is mostly good. Worker will see flag and skip.
        log_essential(f"Requested to clear queue. Drained {cleared_count} items directly.")
        self._notify_queue_updated()
        if self.transcription_queue.qsize() == 0:
            self._notify_transcribing_status(False)


    def _transcription_worker_loop(self):
        # is_fw_model_loaded = False # REMOVED
        # if self.settings.whisper_engine_type == WHISPER_ENGINES[1] and self.faster_whisper_instance: # REMOVED
            # is_fw_model_loaded = self.faster_whisper_instance.load_model() # Pre-load model # REMOVED
            # if not is_fw_model_loaded: # REMOVED
                 # log_error("Failed to pre-load faster-whisper model. Transcription with library may fail.") # REMOVED
                 # Optionally switch to CLI mode here if desired as a fallback, # REMOVED
                 # or let it try to load again per file. # REMOVED
        log_debug("Transcription worker loop started (CLI-only mode).")

        while not self._stop_worker_event.is_set():
            try:
                if self.is_queue_processing_paused:
                    if self._is_transcribing_for_ui: # Ensure status is updated if paused mid-transcription
                        self._notify_transcribing_status(False)
                    log_debug("Transcription worker: Queue processing is paused. Sleeping...")
                    time.sleep(0.5) # Sleep a bit longer when paused
                    continue # Loop back to check pause/stop flags

                if self._stop_worker_event.is_set():
                    log_debug("Transcription worker: Stop event detected before getting from queue.")
                    break

                log_debug("Transcription worker: Attempting to get from queue...")
                try:
                    audio_filepath_str = self.transcription_queue.get(timeout=0.5) # Timeout to check stop_event
                except queue.Empty:
                    # log_debug("Transcription worker: Queue empty.") # Can be noisy
                    if self._is_transcribing_for_ui: self._notify_transcribing_status(False)
                    continue # Loop back to check pause/stop flags

                if audio_filepath_str is AUDIO_QUEUE_SENTINEL:
                    self.transcription_queue.task_done()
                    break # Exit signal

                if self._clear_queue_flag:
                    log_extended(f"Worker skipping {audio_filepath_str} due to clear_queue_flag.")
                    self.transcription_queue.task_done()
                    if self.transcription_queue.empty(): self._clear_queue_flag = False # Reset flag when queue is empty
                    self._notify_queue_updated()
                    continue
                
                audio_filepath = Path(audio_filepath_str)
                if not audio_filepath.exists() or audio_filepath.name.endswith(".transcribed"):
                    log_extended(f"Worker skipping non-existent or already processed: {audio_filepath}")
                    self.transcription_queue.task_done()
                    self._notify_queue_updated()
                    continue

                self._notify_transcribing_status(True)
                self._notify_queue_updated() # Update after get
                log_essential(f"Worker processing: {audio_filepath}")

                transcribed_text = None
                error_msg = None

                # Always use CLI method now
                log_debug(f"Processing {audio_filepath.name} with CLI method.")
                transcribed_text, error_msg = self._transcribe_with_cli(audio_filepath)

                # Post-transcription
                if transcribed_text is not None: # Success
                    parsed_text = self._parse_and_clean_transcription_text(transcribed_text)
                    
                    # Auto-add space if enabled and text exists and doesn't already end with a space
                    if self.settings_manager.settings.auto_add_space and parsed_text and not parsed_text.endswith(' '):
                        parsed_text += ' '
                        log_debug("Auto-added space to transcription.")

                    if self.on_transcription_complete:
                        self.root.after(0, self.on_transcription_complete, parsed_text, audio_filepath)
                    
                    # Rename original audio file
                    try:
                        new_audio_path = audio_filepath.with_suffix(audio_filepath.suffix + ".transcribed")
                        shutil.move(str(audio_filepath), str(new_audio_path))
                        log_extended(f"Renamed processed audio to: {new_audio_path.name}")
                    except Exception as e_mv:
                        log_error(f"Error renaming audio file {audio_filepath.name}: {e_mv}")
                else: # Error during transcription
                    if self.on_transcription_error:
                        self.root.after(0, self.on_transcription_error, audio_filepath, error_msg or "Unknown transcription error.")
                
                self.transcription_queue.task_done()
                self._notify_queue_updated() # Update after task_done

            except Exception as e: # Catchall for unexpected errors in the loop
                log_path_str = 'unknown file'
                if 'audio_filepath' in locals() and audio_filepath is not AUDIO_QUEUE_SENTINEL:
                    log_path_str = str(audio_filepath)
                log_error(f"Critical error in transcription worker for {log_path_str}: {e}", exc_info=True)
                try:
                    # Ensure task_done is called if an item was fetched to prevent deadlocks
                    if 'audio_filepath_str' in locals() and audio_filepath_str is not None:
                         self.transcription_queue.task_done()
                except ValueError: # If task_done called too many times
                    pass
                self._notify_transcribing_status(False)
                time.sleep(1) # Brief pause before retrying loop

        log_essential("Transcription worker loop has exited.")
        self._notify_transcribing_status(False)


    def _transcribe_with_cli(self, audio_path: Path) -> Tuple[Optional[str], Optional[str]]:
        """Returns (transcribed_text, error_message)"""
        current_settings = self.settings_manager.settings
        whisper_exec = Path(current_settings.whisper_executable)
        if not whisper_exec.exists() or not whisper_exec.is_file():
            msg = f"Whisper executable not found or is not a file: {whisper_exec}"
            log_error(msg)
            self.root.after(0, lambda: messagebox.showerror("Whisper CLI Error", msg, parent=self.root))
            return None, msg

        export_dir = Path(current_settings.export_folder) # Use current_settings
        export_dir.mkdir(parents=True, exist_ok=True)

        # Output filename for .txt (Whisper CLI creates this)
        # Ensure it matches how Whisper CLI names it: audio_path_stem.txt
        out_txt_filename = audio_path.stem + ".txt"
        out_txt_filepath = export_dir / out_txt_filename

        command = [
            str(whisper_exec),
            str(audio_path),
            "--model", current_settings.model,
            "--language", current_settings.language,
            "--output_format", "txt", # We only care about the text for this method
            "--output_dir", str(export_dir)
        ]
        
        if self.settings_manager and self.settings_manager.prompt: # Get prompt from settings manager
            # Handle quotes in prompt carefully for CLI
            prompt_content = self.settings_manager.prompt
            if '"' in prompt_content and "'" not in prompt_content:
                command.extend(["--initial_prompt", f"'{prompt_content}'"])
            elif "'" in prompt_content and '"' not in prompt_content:
                command.extend(["--initial_prompt", f'"{prompt_content}"'])
            else: # Contains both or neither, use as is or complex escape
                command.extend(["--initial_prompt", prompt_content])


        command.extend(["--task", "translate" if current_settings.translation_enabled else "transcribe"])

        if not current_settings.whisper_cli_beeps_enabled:
            command.append("--beep_off")
        
        # Add any other CLI flags based on self.settings as needed

        log_extended(f"Running Whisper CLI command: {' '.join(command)}")
        try:
            # Hide console window for subprocess
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                command, check=True, capture_output=True, text=True,
                startupinfo=startupinfo, encoding='utf-8', errors='replace'
            )
            
            # Whisper CLI should create the .txt file. We need to read it.
            if out_txt_filepath.exists():
                with open(out_txt_filepath, 'r', encoding='utf-8') as f:
                    transcribed_text = f.read()
                
                # Optionally delete the .txt file created by CLI if we only want to pass text content
                # For now, let's keep it as it's in the designated export_folder.
                # If clear_text_on_exit is true, it will be cleaned up later.
                
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
                      f"Stderr: {e.stderr[:1000]}" # Limit stderr length
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


    def _parse_and_clean_transcription_text(self, raw_text: str) -> str:
        """Cleans transcription text based on settings (timestamps, etc.)."""
        if not raw_text: return ""
        current_settings = self.settings_manager.settings
        # If no cleaning options are enabled, return text as is (after stripping)
        if not current_settings.clear_text_output and not current_settings.timestamps_disabled:
            return raw_text.strip()

        import re # Keep import local if only used here
        cleaned_lines = []
        # Matches [HH:MM:SS.mmm --> HH:MM:SS.mmm]
        # Or [MM:SS.mmm --> MM:SS.mmm] for shorter files from some Whisper versions
        ts_pattern_long = re.compile(r'^\[\s*\d{2}:\d{2}:\d{2}\.\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}\.\d{3}\s*\]\s*')
        ts_pattern_short = re.compile(r'^\[\s*\d{2}:\d{2}\.\d{3}\s*-->\s*\d{2}:\d{2}\.\d{3}\s*\]\s*')
        # General timestamp pattern to catch both + potential variations if whisper output changes
        ts_pattern_general = re.compile(r'^\[\s*[\d:.]+\s*-->\s*[\d:.]+\s*\]\s*')

        for line in raw_text.splitlines():
            line_stripped = line.strip()
            if not line_stripped:
                continue

            # Option to clear metadata like "---" or "===" or date lines sometimes added by tools
            if current_settings.clear_text_output:
                if line_stripped.startswith("---") or \
                   line_stripped.startswith("===") or \
                   re.match(r'^\d{4}-\d{2}-\d{2}', line_stripped): # Matches YYYY-MM-DD
                    continue
            
            match = ts_pattern_general.match(line_stripped)
            if match:
                if current_settings.timestamps_disabled: # Remove timestamp prefix entirely
                    text_part = line_stripped[match.end():].strip()
                    if text_part: # Only add if there's actual text after timestamp
                        cleaned_lines.append(text_part)
                else: # Timestamps are not disabled, keep the line as is
                    cleaned_lines.append(line_stripped)
            else: # Line does not have a timestamp prefix
                cleaned_lines.append(line_stripped)
        
        return "\n".join(cleaned_lines).strip()

    def execute_command_from_text(self, transcription_text: str):
        """Matches transcription_text against loaded commands and executes action."""
        if not transcription_text.strip() or not self.commands_list:
            return
        
        # TODO: This logic could be more sophisticated (fuzzy matching, regex groups, etc.)
        # Current implementation is simple exact phrase matching (case-insensitive) with wildcard.
        import re # Local import
        
        cleaned_transcription = transcription_text.lower().strip()
        log_extended(f"Attempting to match command in: '{cleaned_transcription}'")

        for cmd_entry in self.commands_list:
            voice_trigger = cmd_entry.voice.strip().lower()
            action_to_run = cmd_entry.action.strip()

            if not voice_trigger or not action_to_run:
                continue

            # Handle " FF " wildcard (case-insensitive for " ff ")
            # Replace " ff " with a regex capture group `(.*?)`
            # Ensure spaces around FF are handled: `\s+FF\s+` might be more robust.
            # For now, simple string replace then regex escape.
            
            pattern_str = voice_trigger
            is_wildcard_command = " ff " in pattern_str # Check before escaping
            
            if is_wildcard_command:
                pattern_str = pattern_str.replace(" ff ", r"%%%%WILDCARD%%%%") # Placeholder

            pattern_str = re.escape(pattern_str) # Escape special characters

            if is_wildcard_command:
                pattern_str = pattern_str.replace(re.escape("%%%%WILDCARD%%%%"), r"(.*?)")


            # Add word boundaries if trigger starts/ends with alphanumeric for better matching
            # \bcat\b matches "cat" but not "caterpillar"
            # If trigger is "open " (ends with space), suffix \b is not good.
            # If trigger is " cat" (starts with space), prefix \b is not good.
            prefix = r'\b' if pattern_str and pattern_str[0].isalnum() else ''
            suffix = r'\b' if pattern_str and pattern_str[-1].isalnum() else ''
            
            # Full pattern: ^ (optional non-word chars) prefix pattern suffix (optional non-word chars) $
            # This is to allow matching if transcription has leading/trailing noise or punctuation.
            # For simplicity, let's try a more direct match first:
            # Match the pattern anywhere in the string, but ensure it's a "whole phrase" match
            # by using word boundaries effectively.
            
            # Revised simpler pattern:
            # If command "play music", transcribed "please play music now" should match.
            # If command "play FF", transcribed "play bohemian rhapsody" should match "bohemian rhapsody".
            
            final_pattern = prefix + pattern_str + suffix
            # For commands that are meant to be at the end, like "stop recording."
            # we might add `[.,!?]?$` to allow optional punctuation at the very end.
            # final_pattern += r'[.,!?]?$' # Optional: if commands usually end sentences

            try:
                match = re.search(final_pattern, cleaned_transcription, re.IGNORECASE)
                if match:
                    matched_action = action_to_run
                    if is_wildcard_command and match.groups():
                        wildcard_content = match.group(1).strip()
                        # Replace "FF" (case-insensitive) in the action string
                        # Using re.sub for case-insensitivity in replacement placeholder
                        matched_action = re.sub(r'\bFF\b', wildcard_content, action_to_run, flags=re.IGNORECASE)
                    
                    log_essential(f"Command matched: '{voice_trigger}' -> Action: '{matched_action}'")
                    self._run_subprocess_action(matched_action)
                    return # Execute only the first matched command

            except re.error as re_err:
                log_error(f"Regex error for command '{voice_trigger}' (pattern: {final_pattern}): {re_err}")
                continue # Skip this command
        log_extended("No command matched.")


    def _run_subprocess_action(self, action_string: str):
        """Executes the command action as a subprocess."""
        log_essential(f"Running subprocess action: '{action_string}'")
        try:
            startupinfo = None
            if os.name == 'nt': # Hide console window on Windows
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            # Using shell=True is a security risk if action_string comes from untrusted source.
            # Here, it's from user's own command config, so slightly less risky but still be cautious.
            # For opening files/URLs, consider `webbrowser` or `os.startfile` (Windows).
            # For running specific executables, pass as a list: `subprocess.run(['calc.exe'])`
            subprocess.run(action_string, shell=True, check=False, # check=False, don't raise error on non-zero exit
                           capture_output=True, text=True, # Capture to avoid polluting main console
                           startupinfo=startupinfo, encoding='utf-8', errors='replace')
            log_extended(f"Subprocess action '{action_string}' completed.")
        except Exception as e:
            log_error(f"Error running subprocess action '{action_string}': {e}", exc_info=True)
            # Optionally notify user via messagebox if action fails
            # self.root.after(0, lambda: messagebox.showerror("Command Error", f"Failed to execute: {action_string}\n{e}", parent=self.root))
