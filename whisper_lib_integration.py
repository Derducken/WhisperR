from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
import threading
import time
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
from settings_manager import AppSettings # For type hinting

# Conditional import for faster-whisper
try:
    from faster_whisper import WhisperModel # type: ignore
    FASTER_WHISPER_AVAILABLE = True
except ImportError:
    FASTER_WHISPER_AVAILABLE = False
    log_extended(
        "faster-whisper library not found. Direct library integration will be disabled. "
        "Install with 'pip install faster-whisper'. For GPU support, also install CUDA/cuDNN related packages."
    )

# Model cache and lock
# Stores {model_name_or_path: WhisperModel_instance}
# Or more complex: {(model_name, device, compute_type): WhisperModel_instance}
_model_cache: Dict[str, Any] = {} # Using Any for WhisperModel for now
_model_cache_lock = threading.Lock()


class FasterWhisperLib:
    def __init__(self, settings_ref: AppSettings):
        self.settings = settings_ref
        self.current_model: Optional[WhisperModel] = None
        self.current_model_key: Optional[str] = None # To track which model is loaded based on settings

        if not FASTER_WHISPER_AVAILABLE:
            log_error("FasterWhisperLib initialized, but faster-whisper library is not available.")

    def _get_model_key(self) -> str:
        # Key for caching should include model name, device, and compute type for full specificity
        # For now, primarily using model name from settings.
        # Future: self.settings.faster_whisper_device, self.settings.faster_whisper_compute_type
        model_name = self.settings.faster_whisper_model_name
        device = "cpu" # Default or from settings
        compute_type = "default" # Default or from settings
        return f"{model_name}_{device}_{compute_type}"


    def load_model(self) -> bool:
        """
        Loads the model specified in settings. Uses a cache.
        Returns True if model is successfully loaded or already loaded, False otherwise.
        """
        if not FASTER_WHISPER_AVAILABLE:
            return False

        model_key = self._get_model_key()
        
        with _model_cache_lock:
            if model_key in _model_cache:
                if self.current_model_key != model_key: # Switching to a different cached model
                    self.current_model = _model_cache[model_key]
                    self.current_model_key = model_key
                    log_essential(f"Switched to cached faster-whisper model: {model_key}")
                elif self.current_model is None : # First time, or cache was populated by another instance
                    self.current_model = _model_cache[model_key]
                    self.current_model_key = model_key
                    log_essential(f"Using already cached faster-whisper model: {model_key}")
                return True # Model is in cache and set as current

            if self.current_model and self.current_model_key == model_key:
                 log_extended(f"Model {model_key} already loaded and current.")
                 return True # Model already loaded and is the correct one

        # If not in cache or not current, try to load
        model_name = self.settings.faster_whisper_model_name
        device = "cpu" # TODO: Make configurable (cpu, cuda, auto)
        compute_type = "default" # TODO: Make configurable (int8, float16, etc.)

        log_essential(f"Loading faster-whisper model: {model_name} (Device: {device}, Compute: {compute_type})")
        
        # Progress reporting hook for model download (if model_management_ui implements it)
        # def progress_hook(current_bytes, total_bytes):
        # print(f"Downloading model: {current_bytes / total_bytes * 100:.2f}%")
        # if self.model_download_progress_callback:
        # self.model_download_progress_callback(current_bytes, total_bytes)

        try:
            # Download model if not found locally (faster-whisper handles this)
            # For download_root, consider making it configurable or use faster-whisper's default
            # download_root = Path(self.settings.faster_whisper_model_path or Path.home() / ".cache" / "faster_whisper")
            # download_root.mkdir(parents=True, exist_ok=True)

            new_model = WhisperModel(
                model_name,
                device=device,
                compute_type=compute_type,
                # download_root=str(download_root) # Optional: specify where models are stored
                # local_files_only=False # Set to True if you only want to use local models
            )
            
            with _model_cache_lock:
                _model_cache[model_key] = new_model
            
            self.current_model = new_model
            self.current_model_key = model_key
            log_essential(f"Faster-whisper model '{model_key}' loaded successfully.")
            return True
        except Exception as e:
            log_error(f"Failed to load faster-whisper model '{model_name}': {e}", exc_info=True)
            self.current_model = None
            self.current_model_key = None
            return False

    def transcribe_audio_file(self, audio_path: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Transcribes an audio file using the loaded faster-whisper model.
        Returns (transcribed_text, error_message).
        Text is None on error.
        """
        if not self.current_model:
            if not self.load_model(): # Attempt to load if not already
                return None, "Model not loaded and failed to load."

        if not self.current_model: # Still no model after attempt
             return None, "No model available for transcription."

        log_essential(f"Transcribing with faster-whisper: {audio_path}")
        start_time = time.monotonic()

        try:
            # Parameters for transcription:
            # language: self.settings.language (if model is multilingual)
            # task: "transcribe" or "translate" (from self.settings.translation_enabled)
            # initial_prompt: self.settings.prompt (or prompt text from main app)
            # word_timestamps: True/False (for more detailed output if needed, default False for segments)
            # vad_filter: True (uses Silero VAD built into faster-whisper, could be an option)
            # beam_size: 5 (default)
            
            task = "translate" if self.settings.translation_enabled else "transcribe"
            lang_code = self.settings.language if self.settings.language and self.settings.language.lower() != "auto" else None

            # TODO: Expose more faster-whisper params in advanced settings
            segments, info = self.current_model.transcribe(
                audio_path,
                language=lang_code,
                task=task,
                initial_prompt=self.settings.prompt, # Assuming self.settings has live prompt
                beam_size=5,
                # word_timestamps=True, # If you want word-level detail
                # vad_filter=True, # Can improve accuracy for long files with silence
                # vad_parameters=dict(min_silence_duration_ms=500) 
            )

            full_text = ""
            # Concatenate text from segments
            # The 'segments' generator yields Segment objects
            # Segment(text, start, end, avg_logprob, no_speech_prob, words, ...)
            # For basic transcription text:
            for segment in segments:
                if self.settings.clear_text_output:
                    # Plain text output without any formatting
                    full_text += segment.text.strip() + " "
                elif self.settings.timestamps_disabled:
                    # Just the text with newlines between segments
                    full_text += segment.text.strip() + "\n"
                else:
                    # Format with timestamps similar to CLI output
                    start_ts = time.strftime('%H:%M:%S', time.gmtime(int(segment.start))) + f".{int((segment.start % 1) * 1000):03d}"
                    end_ts = time.strftime('%H:%M:%S', time.gmtime(int(segment.end))) + f".{int((segment.end % 1) * 1000):03d}"
                    full_text += f"[{start_ts} --> {end_ts}] {segment.text.strip()}\n"
            
            full_text = full_text.strip()

            duration = time.monotonic() - start_time
            log_essential(
                f"Faster-whisper transcription complete in {duration:.2f}s. "
                f"Detected language: {info.language} (Prob: {info.language_probability:.2f})"
            )
            return full_text, None

        except Exception as e:
            log_error(f"Error during faster-whisper transcription: {e}", exc_info=True)
            return None, str(e)

    def unload_model(self):
        """Explicitly unload the current model (optional, Python's GC will handle it eventually)."""
        # faster-whisper models are CTranslate2 models, which manage their own memory.
        # Deleting the Python WhisperModel object should release resources.
        # Cache is cleared on key change or if explicitly managed.
        if self.current_model:
            log_extended(f"Unloading faster-whisper model: {self.current_model_key}")
            # If model_key in _model_cache, can remove it, but careful if other threads use it
            # For now, just nullify current_model. Cache keeps it for reuse.
            # To truly free memory, need to remove from cache and ensure no other refs.
            # with _model_cache_lock:
            #    if self.current_model_key in _model_cache:
            #        del _model_cache[self.current_model_key] # This would free it if no other refs
            
            self.current_model = None 
            self.current_model_key = None
            # Python's garbage collector should handle the underlying CTranslate2 model
            # when the WhisperModel object is no longer referenced.
            # For explicit release, one might need to call a del on the CTranslate2 model object
            # if faster-whisper exposes it, or rely on WhisperModel.__del__.

    @staticmethod
    def get_available_models() -> List[str]:
        """
        Returns a list of known/sensible faster-whisper model names.
        Users can also type custom Hugging Face model IDs.
        """
        # This is just a suggestive list. faster-whisper can download many.
        return [
            "tiny", "tiny.en", "base", "base.en", "small", "small.en",
            "medium", "medium.en", "large-v1", "large-v2", "large-v3",
            "distil-large-v2", "distil-medium.en", "distil-small.en",
            # Systran models (often same as above but from their hub)
            "Systran/faster-whisper-large-v3",
            # Add more common ones if desired
        ]

    @staticmethod
    def get_model_storage_path() -> Path:
        """Returns the default path where faster-whisper models are stored."""
        # This depends on faster-whisper's internal logic, often ~/.cache/huggingface/hub or similar
        # Or if download_root is specified during WhisperModel init.
        # For now, let's assume a common cache location.
        # This is an ESTIMATE. Better to let user manage files manually if specific path needed.
        try:
            from huggingface_hub import constants as hf_constants # type: ignore
            return Path(hf_constants.HF_HUB_CACHE)
        except ImportError:
             # Fallback if huggingface_hub is not directly available or changes
            return Path.home() / ".cache" / "huggingface" / "hub"
