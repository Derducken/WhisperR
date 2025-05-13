import tkinter as tk
from tkinter import messagebox
import sounddevice as sd # type: ignore
import numpy as np # type: ignore
import wave
import threading
import time
import queue
from pathlib import Path
from typing import Optional, List, Callable, Tuple, Any
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning
from persistent_queue_service import PersistentTaskQueue # Added import
from constants import (
    AUDIO_QUEUE_SENTINEL, AUDIO_SAMPLE_RATE, AUDIO_CHANNELS,
    AUDIO_DTYPE, AUDIO_BLOCKSIZE,
    DEFAULT_SILENCE_THRESHOLD_SECONDS, DEFAULT_VAD_ENERGY_THRESHOLD,
    DEFAULT_MAX_MEMORY_SEGMENT_DURATION_SECONDS
)
from settings_manager import AppSettings # For type hinting

# For saving in different formats (MP3, AAC)
# try: # REMOVE PYDUB
    # from pydub import AudioSegment # type: ignore # REMOVE PYDUB
    # PYDUB_AVAILABLE = True # REMOVE PYDUB
# except ImportError: # REMOVE PYDUB
PYDUB_AVAILABLE = False # Always False now
log_extended("pydub library integration is currently disabled. MP3/AAC saving will not be available.")


class AudioService:
    PYDUB_AVAILABLE = PYDUB_AVAILABLE # Make module-level variable available as class attribute

    def __init__(self, settings_manager_ref: Any, root_tk_instance: tk.Tk, persistent_task_queue_ref: PersistentTaskQueue, transcription_service_ref: Optional[Any] = None): # Added persistent_task_queue_ref
        self.settings_manager = settings_manager_ref # Store settings_manager
        self.settings = settings_manager_ref.settings # Keep direct ref, but use settings_manager.settings for dynamic values
        self.root = root_tk_instance # For UI updates (messagebox, after calls)
        self.persistent_task_queue = persistent_task_queue_ref # Store ref to persistent queue
        self.transcription_service = transcription_service_ref # Store ref, still needed for queue size logging for now

        self.is_recording_active: bool = False # Master recording state (controlled by user)
        self.is_vad_speaking: bool = False   # VAD determined speech
        self._current_audio_segment_chunks: List[np.ndarray] = []
        self._silence_start_time: Optional[float] = None
        self._last_chunk_time: Optional[float] = None # Initialize _last_chunk_time
        
        self._audio_stream: Optional[sd.InputStream] = None
        self._recording_thread: Optional[threading.Thread] = None
        self._stop_recording_event = threading.Event() # For signaling thread to stop

        # self.transcription_queue is no longer directly used by AudioService.
        # It will call self.transcription_service.add_to_queue()

        # Callbacks to be set by main app
        self.on_vad_status_change: Optional[Callable[[bool], None]] = None # (is_speaking)
        self.on_audio_segment_saved: Optional[Callable[[Path], None]] = None # (filepath)
        self.on_recording_error: Optional[Callable[[str], None]] = None

        # VAD Calibration
        self.is_calibrating_vad: bool = False
        self.calibration_samples: List[float] = []
        self.calibration_duration_seconds: int = 5 # How long to listen for calibration
        self.calibration_update_callback: Optional[Callable[[float, float, bool], None]] = None #(avg_energy, peak_energy, is_done)
        self.calibration_finished_callback: Optional[Callable[[int], None]] = None # (recommended_threshold)

    # def set_transcription_queue(self, q: queue.Queue): # Removed
    #     self.transcription_queue = q # Removed

    def set_callbacks(self, on_vad_status_change, on_audio_segment_saved, on_recording_error):
        self.on_vad_status_change = on_vad_status_change
        self.on_audio_segment_saved = on_audio_segment_saved
        self.on_recording_error = on_recording_error

    def _notify_vad_status_change(self, new_speaking_status: bool):
        if self.is_vad_speaking != new_speaking_status:
            self.is_vad_speaking = new_speaking_status
            if self.on_vad_status_change:
                self.root.after(0, self.on_vad_status_change, self.is_vad_speaking)

    def start_recording(self):
        if self.is_recording_active:
            log_extended("Recording is already active.")
            return
        
        log_essential("Attempting to start recording...")
        self._stop_recording_event.clear()
        self.is_recording_active = True
        self._current_audio_segment_chunks = []
        self.is_vad_speaking = False # Reset VAD state
        self._silence_start_time = None

        if self._recording_thread and self._recording_thread.is_alive():
            log_extended("Waiting for previous recording thread to finish...")
            self._recording_thread.join(timeout=1.0) # Wait briefly

        self._recording_thread = threading.Thread(target=self._record_audio_loop, daemon=True)
        self._recording_thread.start()
        log_essential("Recording thread started.")
        # Initial VAD status update might be needed if starting in VAD mode
        if self.settings_manager.settings.command_mode: # command_mode implies VAD
            self._notify_vad_status_change(False)
        # self.transcription_queue is now managed by TranscriptionService, AudioService will call add_to_queue on it.


    def stop_recording(self, process_final_segment=True):
        if not self.is_recording_active:
            log_extended("Recording is not active.")
            return

        log_essential("Attempting to stop recording...")
        self._stop_recording_event.set() # Signal the thread to stop
        self.is_recording_active = False # Set master state immediately

        # Wait for the recording thread to actually finish
        if self._recording_thread and self._recording_thread.is_alive():
            log_extended("Waiting for recording thread to terminate...")
            self.root.update_idletasks() # Process pending UI events
            self._recording_thread.join(timeout=2.0) # Increased timeout
            if self._recording_thread.is_alive():
                log_error("Recording thread did not terminate cleanly.")
        
        if process_final_segment and self._current_audio_segment_chunks:
            log_extended("Processing final audio segment on manual stop...")
            self._save_current_segment_and_reset_vad_state()
        else:
            self._current_audio_segment_chunks = [] # Clear any remaining chunks
            self.is_vad_speaking = False
            self._silence_start_time = None
            if self.settings_manager.settings.command_mode:
                self._notify_vad_status_change(False) # Ensure UI reflects stopped VAD

        log_essential(f"Recording stopped. Queue size: {self.transcription_service.transcription_queue.qsize() if self.transcription_service else 'N/A'}")


    def _audio_callback(self, indata: np.ndarray, frames: int, time_info: Any, status: sd.CallbackFlags):
        if status:
            log_extended(f"Audio callback status: {status}")
            if status == sd.CallbackFlags.input_overflow or status == sd.CallbackFlags.input_underflow:
                log_error(f"Audio input issue: {status}. Check system load and audio device.")

        # Allow processing if normal recording is active OR VAD calibration is active.
        # The _stop_recording_event should still be respected for both.
        if self._stop_recording_event.is_set():
            log_extended("Audio callback skipped - stop event is set.")
            return
        
        if not self.is_recording_active and not self.is_calibrating_vad:
            log_extended("Audio callback skipped - neither normal recording nor VAD calibration is active.")
            return

        current_time_monotonic = time.monotonic()
        # Ensure indata is copied, as it's a view into a buffer that might be overwritten
        data_copy = indata.copy()
        log_extended(f"Audio callback - received {len(data_copy)} frames, shape: {data_copy.shape}, dtype: {data_copy.dtype}")
        
        # Debug: Print first few samples
        if len(data_copy) > 0:
            log_debug(f"First 5 samples: {data_copy[:5].flatten()}")

        if self.is_calibrating_vad:
            # Store chunks and calculate RMS for calibration
            if data_copy.size > 0:
                self._current_audio_segment_chunks.append(data_copy)
                self._last_chunk_time = current_time_monotonic
                
                audio_data = data_copy.astype(np.float32)
                rms = np.sqrt(np.mean(audio_data**2))
                log_extended(f"Calibration audio chunk - RMS: {rms:.4f}")
                self.calibration_samples.append(rms)
                
                if self.calibration_update_callback:
                    current_avg = np.mean(self.calibration_samples) if self.calibration_samples else 0
                    current_peak = np.max(self.calibration_samples) if self.calibration_samples else 0
                    self.root.after(0, self.calibration_update_callback, current_avg, current_peak, False)
            else:
                log_error("Empty audio data received during calibration")
            return  # Skip normal VAD processing during calibration

        # Accumulate audio data for normal recording modes
        if data_copy.size > 0:
            self._current_audio_segment_chunks.append(data_copy)
        # else:
            # log_debug("Empty audio data chunk received in normal recording mode.") # Optional: for further debugging

        # VAD Logic (only if command_mode is enabled)
        current_settings = self.settings_manager.settings # Get current settings
        if current_settings.command_mode:
            rms = np.sqrt(np.mean(data_copy.astype(np.float32)**2))
            is_currently_loud = rms >= current_settings.vad_energy_threshold
            
            if is_currently_loud:
                if not self.is_vad_speaking: # Transition to speaking
                    self._notify_vad_status_change(True)
                self._silence_start_time = None
            elif self.is_vad_speaking: # Was speaking, now potentially silent
                if self._silence_start_time is None:
                    self._silence_start_time = current_time_monotonic
                
                if current_time_monotonic - self._silence_start_time >= current_settings.silence_threshold_seconds:
                    # Silence duration met, save segment
                    if self.is_recording_active: # Double check master recording state
                        # Schedule save on main thread via root.after to avoid race conditions with UI/queue
                        log_extended("VAD: Silence detected, saving segment.")
                        self.root.after(0, self._save_current_segment_and_reset_vad_state)
        else:
            # When command_mode is disabled, reset VAD state and don't auto-save segments
            if self.is_vad_speaking:
                self._notify_vad_status_change(False)
            self._silence_start_time = None
            
            # Reset VAD state when command_mode is disabled
            if not current_settings.command_mode and self.is_vad_speaking:
                self._notify_vad_status_change(False)

        # Max segment length check (for both VAD and continuous modes)
        # Calculate current segment duration approximately
        # (num_chunks * blocksize) / samplerate
        current_num_frames = sum(len(chunk) for chunk in self._current_audio_segment_chunks)
        segment_duration_seconds = current_num_frames / AUDIO_SAMPLE_RATE
        
        if segment_duration_seconds >= current_settings.max_memory_segment_duration_seconds:
            log_extended(f"Max segment duration ({current_settings.max_memory_segment_duration_seconds}s) reached, saving segment.")
            if self.is_recording_active:
                 self.root.after(0, self._save_current_segment_and_reset_vad_state)


    def _record_audio_loop(self):
        log_essential(f"Audio recording loop started. Device: {self.settings_manager.settings.selected_audio_device_index or 'default'}")
        
        # Reset audio system completely
        try:
            sd._terminate()
            sd._initialize()
            log_extended("Successfully reinitialized PortAudio")
        except Exception as e:
            log_error(f"Failed to reinitialize PortAudio: {e}")
            if self.on_recording_error:
                self.root.after(0, self.on_recording_error, "Failed to initialize audio system. Please restart the application.")
            return

        # Get default device if none selected
        target_device = self.settings_manager.settings.selected_audio_device_index
        if target_device is None:
            try:
                default_input = sd.default.device[0] if isinstance(sd.default.device, (list, tuple)) else sd.default.device
                if isinstance(default_input, int) and default_input >= 0:
                    target_device = default_input
                    log_extended(f"Using system default input device: {target_device}")
            except Exception as e:
                log_error(f"Could not get default input device: {e}")

        # Verify device selection
        if target_device is not None:
            try:
                device_info = sd.query_devices(target_device)
                log_extended(f"Selected device info: {device_info}")
                if device_info['max_input_channels'] <= 0:
                    log_error(f"Device {target_device} has no input channels")
                    target_device = None
            except Exception as e:
                log_error(f"Selected audio device index {target_device} is invalid or unavailable: {e}")
                target_device = None

        if target_device is None:
            # Try to find any working input device
            try:
                devices = sd.query_devices()
                for i, d in enumerate(devices):
                    if d['max_input_channels'] > 0:
                        target_device = i
                        log_extended(f"Falling back to first available input device: {target_device} ({d['name']})")
                        break
            except Exception as e:
                log_error(f"Could not find any input devices: {e}")

        if target_device is None:
            log_error("No valid input devices found")
            if self.on_recording_error:
                self.root.after(0, self.on_recording_error, "No valid audio input devices found. Please check your audio configuration.")
            return

        try:
            log_extended(f"Attempting to open stream with:\n"
                        f"device={target_device} ({sd.query_devices(target_device)['name']})\n"
                        f"samplerate={AUDIO_SAMPLE_RATE}\n"
                        f"channels={AUDIO_CHANNELS}\n"
                        f"dtype={AUDIO_DTYPE}\n"
                        f"blocksize={AUDIO_BLOCKSIZE}")
            
            with sd.InputStream(
                device=target_device,
                samplerate=AUDIO_SAMPLE_RATE,
                channels=AUDIO_CHANNELS,
                dtype=AUDIO_DTYPE,
                blocksize=AUDIO_BLOCKSIZE,
                callback=self._audio_callback,
                finished_callback=self._stream_finished_callback
            ) as stream:
                self._audio_stream = stream
                log_essential(f"Audio stream opened successfully. Device: {stream.device}")
                log_extended(f"Stream info: {stream}")
                while not self._stop_recording_event.is_set() and self.is_recording_active:
                    # The event wait with timeout acts as a heartbeat and allows the loop
                    # to break if the event is set externally or recording state changes.
                    if self._stop_recording_event.wait(timeout=0.1): # Returns True if event is set
                        break 
                log_extended("Exited record_audio_loop's while loop.")
        except sd.PortAudioError as pae:
            log_error(f"PortAudioError in recording thread: {pae}", exc_info=True)
            error_message = f"Audio device error: {pae}.\n"
            if "Invalid device ID" in str(pae) or "Device unavailable" in str(pae):
                error_message += "The selected audio device might be disconnected or invalid. Please check configuration."
            else:
                error_message += "Try selecting a different audio device or restarting the app."
            if self.on_recording_error:
                self.root.after(0, self.on_recording_error, error_message)
        except Exception as e:
            log_error(f"Generic error in recording thread: {e}", exc_info=True)
            if self.on_recording_error:
                self.root.after(0, self.on_recording_error, f"An unexpected audio error occurred: {e}")
        finally:
            self._audio_stream = None # Clear stream reference
            self.is_recording_active = False # Ensure master state reflects loop exit
            # If VAD was active, ensure its UI state is reset
            if self.settings_manager.settings.command_mode and self.is_vad_speaking: # Use settings_manager
                self._notify_vad_status_change(False)
            log_essential("Audio recording loop finished.")


    def _stream_finished_callback(self):
        """Called by sounddevice when the stream is stopped or aborted."""
        log_extended("Audio stream finished_callback invoked.")
        # This callback runs in a separate thread managed by PortAudio.
        # Avoid heavy processing or direct UI updates here.
        # The main loop logic should handle the state changes.
        # If self.is_recording_active is still true here, it means an unexpected stream stop.
        if self.is_recording_active and not self._stop_recording_event.is_set():
            log_error("Audio stream stopped unexpectedly while recording was marked active.")
            if self.on_recording_error:
                # Schedule error handling on the main Tkinter thread
                self.root.after(0, self.on_recording_error, "Audio stream stopped unexpectedly.")
            # Ensure recording is marked as stopped
            self.is_recording_active = False
            if self.settings_manager.settings.command_mode: # Use settings_manager
                self._notify_vad_status_change(False)


    def _save_current_segment_and_reset_vad_state(self):
        """Saves the current audio buffer and resets VAD state. Must be called from main thread or via root.after()."""
        if not self._current_audio_segment_chunks:
            # If VAD was active, ensure UI is reset even if no audio
            if self.settings_manager.settings.command_mode and self.is_vad_speaking: # Use settings_manager
                self._notify_vad_status_change(False)
            self._silence_start_time = None # Reset silence timer
            return

        segment_to_save = list(self._current_audio_segment_chunks) # Copy the list
        self._current_audio_segment_chunks = [] # Clear buffer for next segment
        
        # Reset VAD state immediately after copying buffer
        if self.settings_manager.settings.command_mode: # Use settings_manager
            self._notify_vad_status_change(False)
        self._silence_start_time = None

        # Perform actual saving in a new thread to avoid blocking UI if saving is slow (e.g. MP3 conversion)
        # However, file I/O for WAV is usually fast. For MP3/AAC, it's better.
        # For now, let's do it directly, but be mindful. If using pydub, threading is good.
        threading.Thread(target=self._save_segment_to_file, args=(segment_to_save,), daemon=True).start()


    def _save_segment_to_file(self, segment_data_chunks: List[np.ndarray]):
        """Internal method to handle file saving. Can be run in a thread."""
        if not segment_data_chunks:
            return

        timestamp = time.strftime("%Y%m%d_%H%M%S") + f"_{int(time.time()*1000)%1000:03d}"
        current_settings = self.settings_manager.settings # Get current settings
        export_dir = Path(current_settings.export_folder)
        try:
            export_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            log_error(f"Error creating export directory '{export_dir}': {e}")
            # Optionally notify user or fallback to a default temp location
            return

        file_format = current_settings.audio_segment_format.lower()
        filename_base = f"recording_{timestamp}"
        output_filepath = export_dir / f"{filename_base}.{file_format}"

        try:
            audio_array = np.concatenate(segment_data_chunks, axis=0)
            if audio_array.size == 0:
                log_extended("Attempted to save an empty audio segment.")
                return

            if file_format == "wav":
                with wave.open(str(output_filepath), 'wb') as wf:
                    wf.setnchannels(AUDIO_CHANNELS)
                    wf.setsampwidth(2) # Use 16-bit (2 bytes) sample width
                    wf.setframerate(AUDIO_SAMPLE_RATE)
                    wf.writeframes(audio_array.tobytes())
            elif file_format in ["mp3", "aac"] and PYDUB_AVAILABLE:
                # Convert numpy array to pydub AudioSegment
                # Ensure data is in bytes and correct format for AudioSegment
                audio_segment = AudioSegment(
                    data=audio_array.tobytes(),
                    sample_width=2, # Use 16-bit (2 bytes) sample width
                    frame_rate=AUDIO_SAMPLE_RATE,
                    channels=AUDIO_CHANNELS
                )
                if file_format == "mp3":
                    audio_segment.export(str(output_filepath), format="mp3")
                elif file_format == "aac":
                    # AAC often uses .m4a or .aac extension. pydub might handle .aac.
                    # FFmpeg needs to be compiled with AAC encoder (libfdk_aac is good, faac also common)
                    audio_segment.export(str(output_filepath), format="adts") # adts is raw AAC stream often in .aac files
                    # Or for m4a container: audio_segment.export(str(output_filepath), format="ipod", codec="aac")
                                       
            elif file_format in ["mp3", "aac"] and not PYDUB_AVAILABLE:
                log_error(f"Cannot save as {file_format}: pydub library not available or FFmpeg missing. Saving as WAV instead.")
                # Fallback to WAV
                output_filepath = export_dir / f"{filename_base}.wav"
                with wave.open(str(output_filepath), 'wb') as wf:
                    wf.setnchannels(AUDIO_CHANNELS)
                    wf.setsampwidth(2) # Use 16-bit (2 bytes) sample width
                    wf.setframerate(AUDIO_SAMPLE_RATE)
                    wf.writeframes(audio_array.tobytes())
            else:
                log_error(f"Unsupported audio format: {file_format}. Defaulting to WAV.")
                # Fallback to WAV (same as above)
                output_filepath = export_dir / f"{filename_base}.wav"
                with wave.open(str(output_filepath), 'wb') as wf:
                     # ... (same WAV saving code) ...
                    wf.setnchannels(AUDIO_CHANNELS)
                    wf.setsampwidth(2) # Use 16-bit (2 bytes) sample width
                    wf.setframerate(AUDIO_SAMPLE_RATE)
                    wf.writeframes(audio_array.tobytes())


            log_essential(f"Audio segment saved: {output_filepath}")
            if current_settings.beep_on_save_audio_segment:
                self.play_beep_sound()

            # Add to persistent queue instead of directly to transcription_service's in-memory queue
            if self.persistent_task_queue:
                if self.persistent_task_queue.add_task(str(output_filepath)):
                    log_extended(f"Successfully added {output_filepath.name} to persistent queue.")
                    # Notify TranscriptionService that there might be new items (optional, if it doesn't poll/reload)
                    if self.transcription_service and hasattr(self.transcription_service, 'check_for_new_tasks'):
                        self.transcription_service.check_for_new_tasks() # This method would need to be added
                else:
                    log_error(f"Failed to add {output_filepath.name} to persistent queue.")
            else:
                log_error("PersistentTaskQueue reference not available in AudioService.")
            
            if self.on_audio_segment_saved:
                self.root.after(0, self.on_audio_segment_saved, output_filepath)

        except Exception as e:
            log_error(f"Error saving audio segment to '{output_filepath}': {e}", exc_info=True)
            # Optionally notify user

    def play_beep_sound(self):
        # This should ideally be handled by a central notification manager or main app
        # to avoid direct OS calls from multiple places.
        try:
            import winsound # Keep it local if only for this
            winsound.Beep(500, 150) # Short beep: frequency, duration_ms
        except ImportError:
            log_extended("winsound not available for beep (non-Windows).")
        except Exception as e:
            log_error(f"Error playing beep: {e}")

    def get_available_audio_devices(self) -> List[Tuple[int, str, str]]: # (index, name, host_api_name)
        devices_info = []
        try:
            devices = sd.query_devices()
            default_input_idx = -1
            default_devices = sd.default.device
            if isinstance(default_devices, (list, tuple)) and len(default_devices) > 0: # (input_idx, output_idx)
                default_input_idx = default_devices[0]
            elif isinstance(default_devices, int): # Only one default (could be input or output)
                # We can't be sure it's input, but often is if only one is set.
                # Better to check device capabilities.
                pass


            for i, d in enumerate(devices):
                if d['max_input_channels'] > 0:
                    host_api_info = sd.query_hostapis(d['hostapi'])
                    host_api_name = host_api_info['name'] if host_api_info else "Unknown API"
                    
                    display_name = f"[{i}] {d['name']} ({host_api_name})"
                    if i == default_input_idx:
                        display_name += " (System Default)"
                    devices_info.append((i, display_name, d['name'])) # Store original name too
        except Exception as e:
            log_error(f"Error querying audio devices: {e}", exc_info=True)
            # Return a dummy entry to indicate error
            devices_info.append((-1, "Error querying devices", "Error"))
        
        if not devices_info:
            devices_info.append((-2, "No input devices found", "None"))
            
        return devices_info

    def update_selected_audio_device(self, device_index: Optional[int]):
        """Updates the sounddevice default and internal setting."""
        # This method modifies self.settings directly, which is fine as it's called from main_app
        # which then saves the settings object held by settings_manager.
        # So, direct self.settings access here is okay.
        original_setting = self.settings_manager.settings.selected_audio_device_index
        self.settings_manager.settings.selected_audio_device_index = device_index
        
        is_currently_recording = self.is_recording_active
        if is_currently_recording:
            log_extended("Audio device change requested while recording. Stopping current recording.")
            self.stop_recording(process_final_segment=False) # Stop without processing, will restart if needed

        try:
            if device_index is not None:
                # Validate device index before setting
                devices = sd.query_devices()
                if 0 <= device_index < len(devices) and devices[device_index]['max_input_channels'] > 0:
                    sd.default.device = device_index # This sets both input and output if not specified
                    # Or more specific: sd.default.device[0] = device_index
                    log_essential(f"Audio input device set to: {device_index} ({devices[device_index]['name']})")
                else:
                    log_error(f"Invalid audio device index: {device_index}. Reverting to previous or default.")
                    self.settings_manager.settings.selected_audio_device_index = original_setting # Revert
                    # sd.default.device = original_setting or None # Set back, None will use system default
            else: # User selected "None" or default
                sd.default.device = None # Let PortAudio pick system default
                log_essential("Audio input device set to system default.")
        except Exception as e:
            log_error(f"Failed to set audio device {device_index}: {e}", exc_info=True)
            self.settings_manager.settings.selected_audio_device_index = original_setting # Revert on error
            # sd.default.device = original_setting or None
            if self.root and self.root.winfo_exists(): # Check if UI is available
                 messagebox.showerror("Audio Device Error", f"Could not set audio device: {e}", parent=self.root)


        if is_currently_recording:
            log_extended("Restarting recording after audio device change.")
            self.start_recording()


    # --- VAD Calibration Methods ---
    def start_vad_calibration(self, duration_seconds: int, update_cb: Callable, finished_cb: Callable):
        if self.is_recording_active:
            messagebox.showwarning("VAD Calibration", "Please stop active recording before starting VAD calibration.", parent=self.root)
            return

        log_essential("Starting VAD calibration...")
        self.is_calibrating_vad = True
        self.calibration_samples = []
        self.calibration_duration_seconds = duration_seconds
        self.calibration_update_callback = update_cb
        self.calibration_finished_callback = finished_cb

        self._stop_recording_event.clear() # Use the same event as normal recording
        
        # Start a temporary recording session for calibration
        self._recording_thread = threading.Thread(target=self._vad_calibration_loop, daemon=True)
        self._recording_thread.start()

    def _vad_calibration_loop(self):
        log_extended(f"VAD calibration loop started. Device: {self.settings_manager.settings.selected_audio_device_index or 'default'}")
        
        # Phase 1: Record silence sample
        self.root.after(0, self.calibration_update_callback, 0, 0, False, "Preparing to record silence...")
        time.sleep(0.5)  # Brief pause before countdown
        
        # Countdown for silence recording (3-2-1)
        for i in range(3, 0, -1):
            self.root.after(0, self.calibration_update_callback, 0, 0, False, f"Recording silence in {i}...")
            time.sleep(1)
            
        # Actual recording (5 seconds)
        self.root.after(0, self.calibration_update_callback, 0, 0, False, "Recording silence... (Please remain silent for 5 seconds)")
        silence_file = self._record_calibration_sample(5)
        if silence_file:
            log_extended(f"Silence sample saved to: {silence_file}")
        else:
            log_error("Failed to save silence sample", exc_info=False)
            if self.calibration_update_callback:
                self.root.after(0, self.calibration_update_callback, 0, 0, True, "Calibration failed: Could not record silence.")
            self.is_calibrating_vad = False # Ensure flag is reset
            return
        
        if self._stop_recording_event.is_set(): # Check after recording attempt
            log_extended("VAD calibration cancelled after silence recording.")
            if self.calibration_update_callback: # Notify UI about cancellation
                self.root.after(0, self.calibration_update_callback, 0, 0, True, "Calibration cancelled.")
            self.is_calibrating_vad = False
            return
            
        # Phase 2: Record speech sample  
        self.root.after(0, self.calibration_update_callback, 0, 0, False, "Preparing to record speech...")
        time.sleep(0.5)  # Brief pause before countdown
        
        # Countdown for speech recording (3-2-1)
        for i in range(3, 0, -1):
            self.root.after(0, self.calibration_update_callback, 0, 0, False, f"Recording speech in {i}...")
            time.sleep(1)
            
        # Actual recording (5 seconds)
        self.root.after(0, self.calibration_update_callback, 0, 0, False, "Recording speech... (Please speak normally for 5 seconds)")
        speech_file = self._record_calibration_sample(5)
        if speech_file:
            log_extended(f"Speech sample saved to: {speech_file}")
        else:
            log_error("Failed to save speech sample", exc_info=False)
            if self.calibration_update_callback:
                self.root.after(0, self.calibration_update_callback, 0, 0, True, "Calibration failed: Could not record speech.")
            self.is_calibrating_vad = False # Ensure flag is reset
            return
        
        if self._stop_recording_event.is_set(): # Check after recording attempt
            log_extended("VAD calibration cancelled after speech recording.")
            if self.calibration_update_callback: # Notify UI about cancellation
                self.root.after(0, self.calibration_update_callback, 0, 0, True, "Calibration cancelled.")
            self.is_calibrating_vad = False
            return
            
        # Analyze the recorded files
        self._analyze_calibration_files(silence_file, speech_file)
        
        # Clean up
        self._audio_stream = None
        self.is_calibrating_vad = False
        log_essential("VAD calibration loop finished.")

    def _record_calibration_sample(self, duration: int) -> Optional[Path]:
        """Record a calibration sample using the normal recording path."""
        
        # Attempt to reset PortAudio state before opening a new stream for calibration
        try:
            log_extended("Re-initializing PortAudio for calibration sample recording...")
            sd._terminate()
            sd._initialize()
            log_extended("PortAudio reinitialized successfully for calibration.")
        except Exception as e:
            log_error(f"Failed to reinitialize PortAudio for calibration: {e}")
            # Proceeding anyway, but this might be the source of issues if it fails
            
        temp_dir = Path(self.settings_manager.settings.export_folder) / "calibration_temp"
        try:
            temp_dir.mkdir(parents=True, exist_ok=True)
            log_extended(f"Created calibration temp dir: {temp_dir}")
        except OSError as e:
            log_error(f"Failed to create calibration temp dir '{temp_dir}': {e}")
            return None

        timestamp = time.strftime("%Y%m%d_%H%M%S") + f"_{int(time.time()*1000)%1000:03d}"
        temp_file = temp_dir / f"calibration_{timestamp}.wav"
        log_extended(f"Will save calibration sample to: {temp_file}")

        # Initialize recording state
        self._current_audio_segment_chunks = []
        self._stop_recording_event.clear()
        self._last_chunk_time = time.monotonic()
        
        target_device = self.settings_manager.settings.selected_audio_device_index
        log_extended(f"Starting calibration recording on device {target_device}...")
        
        try:
            # Use basic stream setup without extra_settings
            with sd.InputStream(
                device=target_device,
                samplerate=AUDIO_SAMPLE_RATE,
                channels=AUDIO_CHANNELS,
                dtype=AUDIO_DTYPE,
                blocksize=AUDIO_BLOCKSIZE,
                callback=self._audio_callback
            ) as stream:
                log_extended(f"Calibration stream opened successfully. Device: {stream.device}")
                start_time = time.monotonic()
                while time.monotonic() - start_time < duration:
                    if self._stop_recording_event.is_set():
                        log_extended("Calibration recording cancelled")
                        return None
                    time.sleep(0.1)
                
                if not self._current_audio_segment_chunks:
                    log_error("No audio data collected during calibration recording", exc_info=False)
                    return None
                    
                # Save the recording
                audio_array = np.concatenate(self._current_audio_segment_chunks, axis=0)
                log_extended(f"Saving calibration recording ({len(audio_array)} frames)...")
                
                try:
                    with wave.open(str(temp_file), 'wb') as wf:
                        wf.setnchannels(AUDIO_CHANNELS)
                        wf.setsampwidth(2)
                        wf.setframerate(AUDIO_SAMPLE_RATE)
                        wf.writeframes(audio_array.tobytes())
                    log_extended("Calibration recording saved successfully")
                    return temp_file
                except Exception as save_error:
                    log_error(f"Failed to save calibration recording: {save_error}")
                    return None
                    
        except Exception as e:
            log_error(f"Error during calibration recording: {e}")
            return None
        finally:
            self._current_audio_segment_chunks = []  # Clear buffer

    def _analyze_calibration_files(self, silence_file: Path, speech_file: Path):
        """Analyze recorded calibration files to determine thresholds."""
        try:
            # Read silence sample
            with wave.open(str(silence_file), 'rb') as wf:
                silence_frames = wf.readframes(wf.getnframes())
                silence_data = np.frombuffer(silence_frames, dtype=np.int16)
                silence_data = silence_data.astype(np.float32) / 32768.0  # Normalize
                
            # Read speech sample
            with wave.open(str(speech_file), 'rb') as wf:
                speech_frames = wf.readframes(wf.getnframes())
                speech_data = np.frombuffer(speech_frames, dtype=np.int16)
                speech_data = speech_data.astype(np.float32) / 32768.0  # Normalize
                
            # Calculate RMS energy for both
            silence_rms = np.sqrt(np.mean(silence_data**2))
            speech_rms = np.sqrt(np.mean(speech_data**2))
            
            # Calculate recommended threshold
            recommended_threshold = int((silence_rms + speech_rms) / 2 * 32768)  # Scale back to int16 range
            recommended_threshold = max(recommended_threshold, int(silence_rms * 1.2 * 32768))  # At least 20% above silence
            recommended_threshold = max(recommended_threshold, 50)  # Absolute minimum
            
            log_essential(
                f"VAD Calibration results:\n"
                f"Silence RMS: {silence_rms:.4f}\n"
                f"Speech RMS: {speech_rms:.4f}\n"
                f"Recommended Threshold: {recommended_threshold}"
            )
            
            if self.calibration_update_callback:
                self.root.after(0, self.calibration_update_callback, 
                              silence_rms, np.max(np.abs(silence_data)), True,
                              f"Calibration complete\nRecommended: {recommended_threshold}")
            if self.calibration_finished_callback:
                self.root.after(0, self.calibration_finished_callback, recommended_threshold)
                
        except Exception as e:
            log_error(f"Error analyzing calibration files: {e}")
            self.root.after(0, lambda: messagebox.showerror(
                "Calibration Failed",
                f"Error analyzing calibration files:\n{e}",
                parent=self.root
            ))
            if self.calibration_update_callback:
                self.root.after(0, self.calibration_update_callback, 0, 0, True, "Calibration failed - analysis error")

    def cancel_vad_calibration(self):
        if self.is_calibrating_vad:
            log_extended("Cancelling VAD calibration via external request.")
            self._stop_recording_event.set() # Signal the calibration loop to stop
            # _finish_vad_calibration will be called by the loop's finally block
