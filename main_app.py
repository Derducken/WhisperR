import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
from pathlib import Path
import threading
import time
import traceback
import platform
from typing import Optional, Callable, List, Dict, Any, Tuple

# --- Project specific imports ---
try:
    # Ensure all used logger helper functions are imported
    from app_logger import (
        AppLogger, get_logger, log_error, log_warning,
        log_essential, log_extended, log_debug
    )
    import app_logger # To initialize LOGGER global

    from constants import (
        DEFAULT_HOTKEY_TOGGLE, DEFAULT_HOTKEY_SHOW, APP_ICON_NAME,
        AUDIO_QUEUE_SENTINEL, COLOR_STATUS_IDLE_NOT_RECORDING,
        COLOR_STATUS_RECORDING_CONTINUOUS, COLOR_STATUS_RECORDING_VAD_ACTIVE,
        COLOR_STATUS_RECORDING_VAD_WAITING, COLOR_STATUS_TRANSCRIBING,
        COLOR_STATUS_RECORDING_AND_TRANSCRIBING, WHISPER_ENGINES, Theme as AppThemeEnum # Use alias
    )
    from settings_manager import SettingsManager, get_user_config_dir, get_app_asset_path, AppSettings
    from theme_manager import ThemeManager # Theme enum itself is in constants now
    from hotkey_manager import HotkeyManager
    from tray_icon_manager import TrayIconManager
    from status_bar_manager import StatusBarManager, WINDOWS_FEATURES_AVAILABLE
    from alt_status_indicator import AltStatusIndicator
    from audio_service import AudioService
    from transcription_service import TranscriptionService
    # from whisper_lib_integration import FASTER_WHISPER_AVAILABLE # REMOVE

    from main_window_view import MainWindowView
    from config_window_view import ConfigWindowView
    from scratchpad_view import ScratchpadWindow
    from command_editor_view import CommandEditorWindow
    from vad_calibration_dialog import VADCalibrationDialog
except ImportError as e:
    initial_error_msg = f"Critical Import Error: {e}\n" \
                        "Please ensure all application files are present in the correct directory structure.\n" \
                        f"Missing module might be related to: {e.name}"
    try:
        root_err = tk.Tk()
        root_err.withdraw()
        messagebox.showerror("Application Startup Error", initial_error_msg)
        root_err.destroy()
    except Exception:
        print(initial_error_msg, file=sys.stderr)
    sys.exit(1)


class WhisperRApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.app_asset_path = get_app_asset_path() # For assets like icons
        self.user_config_path = get_user_config_dir("WhisperR") # For configs and logs

        app_logger.LOGGER = AppLogger(self.user_config_path) # Logs go to user config dir
        log_essential("WhisperR Application starting...")
        log_debug(f"Application asset path: {self.app_asset_path}")
        log_debug(f"User config path: {self.user_config_path}")

        self.settings_manager = SettingsManager(self.user_config_path) # Settings manager uses user config dir
        self.settings: AppSettings = self.settings_manager.settings

        self.theme_manager = ThemeManager()
        self.theme_manager.apply_theme(self.root, self.settings.ui_theme)

        self.root.title("WhisperR")
        self.root.minsize(650, 600)
        self._set_app_icon()

        # Initialize TranscriptionService first as AudioService needs a reference to it
        self.transcription_service = TranscriptionService(self.settings_manager, self.root, self.settings_manager)
        self.audio_service = AudioService(self.settings_manager, self.root, self.transcription_service) # Pass transcription_service
        self.hotkey_manager = HotkeyManager(self.root)

        self.status_bar_win_manager: Optional[StatusBarManager] = None
        if WINDOWS_FEATURES_AVAILABLE:
            self.status_bar_win_manager = StatusBarManager(self.root, self.theme_manager)

        self.alt_status_indicator: Optional[AltStatusIndicator] = None
        try:
            self.alt_status_indicator = AltStatusIndicator(self.root, self.app_asset_path, self.theme_manager) # Use asset_path
        except Exception as e_alt:
            log_error(f"Failed to initialize AltStatusIndicator: {e_alt}", exc_info=True)

        self.main_view = MainWindowView(
            self.root,
            self.settings_manager, # Pass SettingsManager instance
            self.settings_manager.prompt, # Pass the actual prompt string
            self.theme_manager
        )
        self.config_window: Optional[ConfigWindowView] = None
        self.scratchpad_window: Optional[ScratchpadWindow] = None
        self.command_editor_window: Optional[CommandEditorWindow] = None

        self.tray_manager = TrayIconManager(
            app_name="WhisperR", root_window=self.root,
            show_window_action=self._action_show_window,
            toggle_recording_action=self._action_toggle_recording_external,
            quit_action=self._action_quit_application,
            base_path=self.app_asset_path # Use asset_path for tray icon resources
        )
        self.tray_thread: Optional[threading.Thread] = None

        self.is_shutting_down = False
        self._ui_transcribing_active = False

        self._connect_signals_and_slots()
        self._initialize_services_and_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._handle_close_button)
        log_essential("WhisperR initialization complete.")

    def _set_app_icon(self):
        try:
            icon_path_obj = self.app_asset_path / APP_ICON_NAME # Use asset_path
            # Fallback for dev mode if icon is not found relative to asset_path (e.g. settings_manager.py)
            # but directly in the script's execution directory.
            if not icon_path_obj.exists() and not hasattr(sys, '_MEIPASS'):
                 icon_path_obj = Path(os.path.dirname(os.path.abspath(__file__))) / APP_ICON_NAME

            if icon_path_obj.exists():
                if platform.system() == "Windows":
                    ico_path = icon_path_obj.with_suffix(".ico")
                    if ico_path.exists():
                        self.root.iconbitmap(default=str(ico_path))
                        log_debug(f"Set .ico window icon: {ico_path}")
                        return
                from PIL import Image, ImageTk
                img = Image.open(icon_path_obj)
                photo_img = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo_img)
                self.root.app_icon_photo_ref = photo_img
                log_debug(f"Set .png window icon: {icon_path_obj}")
            else:
                log_warning(f"Application icon file '{APP_ICON_NAME}' not found at {icon_path_obj}.")
        except Exception as e:
            log_error(f"Error setting application icon: {e}", exc_info=True)

    def _connect_signals_and_slots(self):
        # self.audio_service.set_transcription_queue(self.transcription_service.transcription_queue) # Removed, AudioService now calls add_to_queue on TranscriptionService
        self.audio_service.set_callbacks(
            on_vad_status_change=self._handle_vad_status_change,
            on_audio_segment_saved=self._handle_audio_segment_saved,
            on_recording_error=self._handle_audio_recording_error
        )
        self.transcription_service.set_callbacks(
            on_transcription_complete=self._handle_transcription_complete,
            on_transcription_error=self._handle_transcription_error,
            on_queue_updated=self.main_view.update_queue_indicator_ui,
            on_transcribing_status_changed=self._handle_transcribing_status_change_for_ui
        )
        self.transcription_service.update_commands_list(self.settings_manager.commands)
        self.hotkey_manager.set_callbacks(
            toggle_record_cb=self._action_toggle_recording_external,
            show_window_cb=self._action_show_window
        )
        self.main_view.bind_language_change(lambda lang: self._update_setting_and_save('language', lang))
        self.main_view.bind_model_change(lambda model: self._update_setting_and_save('model', model))
        self.main_view.bind_toggle_change("translation", lambda val: self._update_setting_and_save('translation_enabled', val))
        self.main_view.bind_toggle_change("command_mode", self._handle_command_mode_change)
        self.main_view.bind_toggle_change("timestamps_disabled", lambda val: self._update_setting_and_save('timestamps_disabled', val))
        self.main_view.bind_toggle_change("clear_text_output", lambda val: self._update_setting_and_save('clear_text_output', val))
        self.main_view.bind_prompt_change(self._handle_prompt_change)
        self.main_view.set_button_command("scratchpad", self._action_open_scratchpad)
        self.main_view.set_button_command("start_stop", self._action_toggle_recording_ui)
        self.main_view.set_button_command("ok_hide", self._action_ok_hide_window) # Bind new button
        self.main_view.set_button_command("clear_queue", self._action_clear_queue)
        self.main_view.set_button_command("pause_queue", self._action_toggle_pause_queue)
        self.main_view.add_menu_command("file", "Import Prompt...", self._action_import_prompt_file)
        self.main_view.add_menu_command("file", "Export Prompt...", self._action_export_prompt_file)
        self.main_view.add_menu_command("file", type="separator")
        self.main_view.add_menu_command("file", "Open Scratchpad", self._action_open_scratchpad)
        self.main_view.add_menu_command("file", type="separator")
        self.main_view.add_menu_command("file", "Quit WhisperR", self._action_quit_application)
        self.main_view.add_menu_command("settings", "Configuration...", self._action_open_config_window)
        self.main_view.add_menu_command("settings", "Configure Commands...", self._action_open_command_editor)
        self.main_view.add_menu_command("queue", "Pause/Resume Queue Processing",
                                        self._action_toggle_pause_queue, type="checkbutton",
                                        variable=self.main_view.pause_queue_menu_var)
        self.main_view.add_menu_command("queue", "Clear Transcription Queue", self._action_clear_queue)

    def _initialize_services_and_ui(self):
        log_essential("Initializing services and UI components...")
        self.main_view.update_ui_from_settings()
        self.audio_service.update_selected_audio_device(self.settings.selected_audio_device_index)
        if not self.hotkey_manager.update_hotkeys(self.settings.hotkey_toggle_record, self.settings.hotkey_show_window):
            messagebox.showwarning("Hotkey Error", "Could not register one or more global hotkeys on startup. Please check Configuration.", parent=self.root)
        self.transcription_service.start_worker()
        if self.status_bar_win_manager:
            self.status_bar_win_manager.configure(
                enabled=self.settings.status_bar_enabled,
                position=self.settings.status_bar_position,
                size=self.settings.status_bar_size
            )
        if self.alt_status_indicator:
            self.alt_status_indicator.configure(
                enabled=self.settings.alt_status_indicator_enabled,
                position=self.settings.alt_status_indicator_position,
                size=self.settings.alt_status_indicator_size,
                offset=self.settings.alt_status_indicator_offset
            )
        self._update_all_status_indicators()
        self.tray_thread = threading.Thread(target=self.tray_manager.setup_tray_icon, daemon=True)
        self.tray_thread.start() 
        # log_warning("DEBUG: Tray icon thread start is currently disabled.")

    def _update_setting_and_save(self, setting_name: str, value: Any, save_type: str = "settings"):
        current_value = getattr(self.settings, setting_name, None)
        if current_value != value:
            log_debug(f"Setting '{setting_name}' changed to '{value}' from '{current_value}'. Saving '{save_type}'.")
            setattr(self.settings, setting_name, value)
            if save_type == "settings": self.settings_manager.save_settings()
            elif save_type == "prompt": self.settings_manager.save_prompt()
            elif save_type == "commands": self.settings_manager.save_commands()
            if setting_name == "ui_theme":
                self.theme_manager.apply_theme(self.root, value)
                if self.config_window and self.config_window.winfo_exists():
                    self.theme_manager.apply_theme(self.config_window, value)
                if self.scratchpad_window and self.scratchpad_window.winfo_exists():
                     self.theme_manager.apply_theme(self.scratchpad_window, value)
                     if hasattr(self.scratchpad_window, '_apply_theme'): self.scratchpad_window._apply_theme()
                if self.command_editor_window and self.command_editor_window.winfo_exists():
                    self.theme_manager.apply_theme(self.command_editor_window, value)
                    if hasattr(self.command_editor_window, '_apply_theme'): self.command_editor_window._apply_theme()
                if self.alt_status_indicator: self.alt_status_indicator.update_theme()

    def _handle_command_mode_change(self, new_value: bool):
        self._update_setting_and_save('command_mode', new_value)
        self._update_all_status_indicators()

    def _handle_prompt_change(self, new_prompt: str):
        if self.settings_manager.prompt != new_prompt: # Compare against manager's copy
            self.settings_manager.prompt = new_prompt
            # self.settings.prompt is not a field, so no need to update it on AppSettings.
            self.settings_manager.save_prompt()
            log_essential("Prompt updated and saved.")

    def _action_toggle_recording_ui(self):
        if self.audio_service.is_recording_active:
            self.audio_service.stop_recording()
        else:
            self.audio_service.start_recording()
        self.main_view.update_recording_indicator_ui(self.audio_service.is_recording_active, self.audio_service.is_vad_speaking)
        self._update_all_status_indicators()

    def _action_toggle_recording_external(self):
        log_debug("External toggle recording requested.")
        self.root.after(0, self._action_toggle_recording_ui)

    def _action_clear_queue(self):
        if messagebox.askyesno("Clear Queue", "Remove all pending items from the transcription queue?", parent=self.root):
            self.transcription_service.clear_queue()

    def _action_toggle_pause_queue(self):
        self.transcription_service.toggle_pause_queue()
        self.main_view.update_pause_queue_button_ui(self.transcription_service.is_queue_processing_paused)

    def _action_show_window(self):
        log_debug("Show window hotkey action triggered.")
        try:
            if not self.root.winfo_exists():
                return

            is_main_window_visible = self.root.winfo_viewable()

            if not is_main_window_visible:
                # Show windows
                log_debug("Showing main window and eligible scratchpad.")
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                if self.scratchpad_window and self.scratchpad_window.winfo_exists() and \
                   not self.scratchpad_window.is_explicitly_closed():
                    self.scratchpad_window.show()
            else:
                # Hide windows
                log_debug("Hiding main window and eligible scratchpad via hotkey.")
                self.root.withdraw()
                if self.scratchpad_window and self.scratchpad_window.winfo_exists() and \
                   not self.scratchpad_window.is_explicitly_closed() and \
                   self.scratchpad_window.is_visible(): # Only hide if currently visible
                    self.scratchpad_window.withdraw()
        except Exception as e:
            log_error(f"Error in _action_show_window (toggle): {e}", exc_info=True)

    def _action_ok_hide_window(self):
        """Hides the main application window to the tray. Leaves scratchpad visible if open."""
        log_debug("OK (Hide Window) action triggered.")
        self.root.withdraw()
        # No notification needed here as it's an explicit user action via button.

    def _action_hide_window(self):
        """Hides the main application window to the tray, typically via close button or system action."""
        self.root.withdraw()
        # if self.tray_manager.tray_icon and self.tray_manager.tray_icon.visible: # DEBUG: Disable tray notify
        #      self.tray_manager.notify("WhisperR Hidden", f"{self.tray_manager.app_name} is running in the background.")
        log_warning("DEBUG: Tray notification on hide is currently disabled.")

    def _action_quit_application(self):
        log_essential("Quit action initiated...")
        if self.is_shutting_down: return
        self.is_shutting_down = True
        if self.audio_service.is_recording_active:
            self.audio_service.stop_recording(process_final_segment=True)
        self.hotkey_manager.cleanup()
        self.transcription_service.stop_worker()
        if self.settings.clear_audio_on_exit or self.settings.clear_text_on_exit:
            self._delete_session_files_on_exit()
        self.settings_manager.prompt = self.main_view.get_prompt_text()
        self.settings_manager.save_all()
        if self.status_bar_win_manager: self.status_bar_win_manager.destroy_status_bar()
        if self.alt_status_indicator: self.alt_status_indicator.destroy_indicator()
        if self.tray_manager: self.tray_manager.stop_tray_icon() 
        if self.tray_thread and self.tray_thread.is_alive(): 
            self.tray_thread.join(timeout=1.0) 
        # log_warning("DEBUG: Tray icon stop/join calls in quit_application are currently disabled.")
        get_logger().close()
        if self.root.winfo_exists():
            self.root.after(50, self.root.destroy)
        else:
            sys.exit(0)

    def _handle_close_button(self):
        # Use AppThemeEnum.SYSTEM.value if Theme enum is named AppThemeEnum in constants.py
        # Assuming constants.py defines CloseBehavior enum and DEFAULT_CLOSE_BEHAVIOR
        if self.settings.close_behavior == self.settings_manager.settings.close_behavior: # Compare with current setting
            if self.settings.close_behavior == "Minimize to tray": # Direct string comparison
                self._action_hide_window()
            else: self._action_quit_application() # Exit app
        elif self.settings.close_behavior == "Minimize to tray": # Direct string comparison
             self._action_hide_window()
        else: self._action_quit_application()


    def _handle_vad_status_change(self, is_speaking: bool):
        self.main_view.update_recording_indicator_ui(self.audio_service.is_recording_active, is_speaking)
        self._update_all_status_indicators()

    def _handle_audio_segment_saved(self, audio_filepath: Path):
        log_extended(f"Audio segment saved (callback in MainApp): {audio_filepath.name}")
        if self.settings.beep_on_save_audio_segment:
            self.audio_service.play_beep_sound()

    def _handle_audio_recording_error(self, error_message: str):
        log_error(f"Audio Recording Error: {error_message}")
        messagebox.showerror("Audio Error", error_message, parent=self.root)
        if self.audio_service.is_recording_active:
            self.audio_service.stop_recording(process_final_segment=False)
        self.main_view.update_recording_indicator_ui(False, False)
        self._update_all_status_indicators()

    def _handle_transcription_complete(self, transcribed_text: str, original_audio_path: Path):
        log_essential(f"Transcription complete for {original_audio_path.name}. Length: {len(transcribed_text)}")
        if self.settings.beep_on_transcription:
            self.audio_service.play_beep_sound()
        if self.scratchpad_window and self.scratchpad_window.winfo_exists():
            self.scratchpad_window.add_text(transcribed_text)
        try:
            self.root.clipboard_clear(); self.root.clipboard_append(transcribed_text)
            log_extended("Transcription copied to clipboard.")
        except tk.TclError as e:
            log_error(f"Failed to update clipboard: {e} (Window might be closing)")
        if self.settings.auto_paste:
            delay_ms = int(max(0, self.settings.auto_paste_delay) * 1000)
            self.root.after(delay_ms, self._perform_auto_paste)
        if self.settings.command_mode:
            self.transcription_service.execute_command_from_text(transcribed_text)

    def _handle_transcription_error(self, audio_path: Path, error_message: str):
        log_error(f"Transcription Error for {audio_path.name}: {error_message}")

    def _handle_transcribing_status_change_for_ui(self, is_transcribing: bool):
        self._ui_transcribing_active = is_transcribing
        self._update_all_status_indicators()

    def _perform_auto_paste(self):
        try:
            import keyboard
            keyboard.press_and_release('ctrl+v')
            log_extended("Auto-pasted transcription.")
        except Exception as e:
            log_error(f"Error during auto-paste: {e}")

    def _action_open_scratchpad(self):
        if not self.scratchpad_window or not self.scratchpad_window.winfo_exists():
            self.scratchpad_window = ScratchpadWindow(self.root, self.settings, self.theme_manager)
            # For a newly created window, mark_as_opened_by_user will also call show()
            self.scratchpad_window.mark_as_opened_by_user() 
        else:
            self.scratchpad_window.mark_as_opened_by_user()

    def _action_open_config_window(self):
        if self.config_window and self.config_window.winfo_exists():
            self.config_window.lift(); self.config_window.focus_force(); return
        audio_devices = self.audio_service.get_available_audio_devices()
        self.config_window = ConfigWindowView(
            self.root, # tk_parent
            self, # app_instance
            settings=self.settings, theme_manager=self.theme_manager,
            audio_devices_list=audio_devices,
            save_config_callback=self._save_configuration_from_dialog,
            record_hotkey_callback=lambda: self.hotkey_manager.record_new_hotkey(self.config_window if self.config_window else self.root),
            vad_calibrate_callback=self._trigger_vad_calibration,
            open_command_editor_callback=self._action_open_command_editor,
            delete_session_files_callback=self._action_delete_session_files_now
        )

    def _action_open_command_editor(self):
        if self.command_editor_window and self.command_editor_window.winfo_exists():
            self.command_editor_window.lift(); self.command_editor_window.focus_force(); return
        self.command_editor_window = CommandEditorWindow(
            self.root, # tk_parent
            self, # app_instance
            current_commands=self.settings_manager.commands,
            save_callback=self._save_commands_from_editor, theme_manager=self.theme_manager
        )

    def _save_commands_from_editor(self, new_commands_list):
        self.settings_manager.commands = new_commands_list
        self.settings_manager.save_commands()
        self.transcription_service.update_commands_list(new_commands_list)
        log_essential("Commands updated from editor.")
        if self.command_editor_window: self.command_editor_window.destroy()

    def _save_configuration_from_dialog(self, new_settings_from_dialog: AppSettings) -> bool:
        log_essential("Saving configuration from dialog...")
        old_settings = self.settings # Keep a reference to the old settings object (or its copy)
        self.settings_manager.settings = new_settings_from_dialog
        self.settings = self.settings_manager.settings # Update self.settings to the new object

        hotkeys_ok = self.hotkey_manager.update_hotkeys(
            self.settings.hotkey_toggle_record, self.settings.hotkey_show_window
        )
        self.main_view.update_shortcut_display_ui()
        if old_settings.selected_audio_device_index != self.settings.selected_audio_device_index:
            self.audio_service.update_selected_audio_device(self.settings.selected_audio_device_index)
        # if old_settings.whisper_engine_type != self.settings.whisper_engine_type or \ # REMOVE - engine choice will be gone
        #    (self.settings.whisper_engine_type == WHISPER_ENGINES[1] and \ # REMOVE
        #     old_settings.faster_whisper_model_name != self.settings.faster_whisper_model_name): # REMOVE
            # if self.transcription_service.faster_whisper_instance: # REMOVE
                # self.transcription_service.faster_whisper_instance.unload_model() # REMOVE
                # log_essential("Whisper engine or model changed. Model will be reloaded on next use.") # REMOVE
        log_debug("Faster-Whisper related model change check skipped as it's CLI-only now.")
        if self.status_bar_win_manager:
            if old_settings.status_bar_enabled != self.settings.status_bar_enabled or \
               old_settings.status_bar_position != self.settings.status_bar_position or \
               old_settings.status_bar_size != self.settings.status_bar_size:
                self.status_bar_win_manager.configure(
                    self.settings.status_bar_enabled, self.settings.status_bar_position, self.settings.status_bar_size
                )
        if self.alt_status_indicator:
            if old_settings.alt_status_indicator_enabled != self.settings.alt_status_indicator_enabled or \
               old_settings.alt_status_indicator_position != self.settings.alt_status_indicator_position or \
               old_settings.alt_status_indicator_size != self.settings.alt_status_indicator_size or \
               old_settings.alt_status_indicator_offset != self.settings.alt_status_indicator_offset:
                self.alt_status_indicator.configure(
                    self.settings.alt_status_indicator_enabled, self.settings.alt_status_indicator_position,
                    self.settings.alt_status_indicator_size, self.settings.alt_status_indicator_offset
                )
        if old_settings.ui_theme != self.settings.ui_theme:
            self.theme_manager.apply_theme(self.root, self.settings.ui_theme)
            if self.scratchpad_window and self.scratchpad_window.winfo_exists():
                self.theme_manager.apply_theme(self.scratchpad_window, self.settings.ui_theme)
                if hasattr(self.scratchpad_window, '_apply_theme'): self.scratchpad_window._apply_theme()
            if self.command_editor_window and self.command_editor_window.winfo_exists():
                 self.theme_manager.apply_theme(self.command_editor_window, self.settings.ui_theme)
                 if hasattr(self.command_editor_window, '_apply_theme'): self.command_editor_window._apply_theme()
            if self.alt_status_indicator: self.alt_status_indicator.update_theme()
        get_logger().configure(self.settings.logging_level, self.settings.log_to_file, self.settings.max_log_files)
        self.settings_manager.save_settings()
        self.main_view.update_ui_from_settings()
        self._update_all_status_indicators()
        return hotkeys_ok

    def _trigger_vad_calibration(self, current_threshold: int) -> Optional[int]:
        if self.audio_service.is_recording_active:
            messagebox.showwarning("Calibration Busy", "Cannot start VAD calibration while recording is active.", parent=self.config_window if self.config_window else self.root)
            return None
        recommended_threshold = VADCalibrationDialog.show(
            parent=self.config_window if self.config_window else self.root,
            start_cb=self.audio_service.start_vad_calibration,
            cancel_cb=self.audio_service.cancel_vad_calibration,
            initial_thresh=current_threshold, theme_manager=self.theme_manager
        )
        return recommended_threshold

    def _action_import_prompt_file(self):
        filepath = filedialog.askopenfilename(
            title="Import Prompt File", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")), parent=self.root
        )
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as f: content = f.read()
                self.main_view.set_prompt_widget_text(content) # Use dedicated setter
                self._handle_prompt_change(content)
                log_essential(f"Prompt imported from {filepath}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import prompt: {e}", parent=self.root)

    def _action_export_prompt_file(self):
        content = self.main_view.get_prompt_text()
        filepath = filedialog.asksaveasfilename(
            title="Export Prompt As", filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")),
            defaultextension=".txt", parent=self.root
        )
        if filepath:
            try:
                with open(filepath, 'w', encoding='utf-8') as f: f.write(content)
                log_essential(f"Prompt exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export prompt: {e}", parent=self.root)

    def _action_delete_session_files_now(self):
        export_dir = Path(self.settings.export_folder)
        parent_win = self.config_window if self.config_window and self.config_window.winfo_exists() else self.root
        if not export_dir.is_dir():
            messagebox.showwarning("Cleanup Warning", f"Export directory not found:\n{export_dir}", parent=parent_win); return
        del_audio, del_text = self.settings.clear_audio_on_exit, self.settings.clear_text_on_exit
        if self.config_window and self.config_window.winfo_exists():
            del_audio = self.config_window.clear_audio_on_exit_var.get()
            del_text = self.config_window.clear_text_on_exit_var.get()
        if not (del_audio or del_text):
            messagebox.showinfo("Cleanup Info", "Enable 'Clear Audio' or 'Clear Text' in Config (Advanced tab) to select file types for deletion.", parent=parent_win); return
        self._perform_file_deletion(export_dir, del_audio, del_text, ask_confirm=True, parent_window=parent_win)

    def _delete_session_files_on_exit(self):
        export_dir = Path(self.settings.export_folder)
        if not export_dir.is_dir(): return
        self._perform_file_deletion(export_dir, self.settings.clear_audio_on_exit, self.settings.clear_text_on_exit, ask_confirm=False)

    def _perform_file_deletion(self, directory: Path, delete_audio: bool, delete_text: bool, ask_confirm: bool, parent_window=None):
        files_to_delete = []
        try:
            for item in directory.iterdir():
                if item.is_file():
                    name = item.name
                    if delete_audio and name.startswith("recording_") and \
                       (any(name.endswith(ext) for ext in [".wav", ".mp3", ".aac", ".m4a"]) or \
                        any(name.endswith(ext + ".transcribed") for ext in [".wav", ".mp3", ".aac", ".m4a"])):
                        files_to_delete.append(item)
                    elif delete_text and name.startswith("recording_") and name.endswith(".txt"):
                        files_to_delete.append(item)
        except Exception as e:
            log_error(f"Error reading export directory '{directory}' for cleanup: {e}")
            if ask_confirm: messagebox.showerror("Cleanup Error", f"Error reading export directory:\n{e}", parent=parent_window)
            return
        if not files_to_delete:
            if ask_confirm: messagebox.showinfo("Cleanup Info", f"No matching session files found in '{directory}'.", parent=parent_window)
            return
        confirmed = not ask_confirm
        if ask_confirm:
            types_str = [t for t,b in [("Audio",delete_audio),("Text",delete_text)] if b]
            msg = f"Permanently delete {len(files_to_delete)} session file(s) ({', '.join(types_str)}) from:\n'{directory}'?\nThis cannot be undone."
            if messagebox.askyesno("Confirm Deletion", msg, icon=messagebox.WARNING, parent=parent_window): confirmed = True
        if confirmed:
            deleted_count, error_count = 0, 0
            for f_path in files_to_delete:
                try: f_path.unlink(); deleted_count += 1
                except Exception as e: error_count += 1; log_error(f"Error deleting file {f_path}: {e}")
            result_msg = f"Deleted {deleted_count} file(s)." + (f" Failed: {error_count}." if error_count > 0 else "")
            log_essential(f"File Cleanup: {result_msg}")
            if ask_confirm: messagebox.showinfo("Cleanup Result", result_msg, parent=parent_window)
        elif ask_confirm: log_extended("File deletion cancelled by user.")

    def _get_current_status_indicator_color(self) -> str:
        is_rec, is_vad_speak = self.audio_service.is_recording_active, self.audio_service.is_vad_speaking
        is_trans = self._ui_transcribing_active
        if is_rec and is_trans: return COLOR_STATUS_RECORDING_AND_TRANSCRIBING
        elif is_rec:
            return (COLOR_STATUS_RECORDING_VAD_ACTIVE if is_vad_speak else COLOR_STATUS_RECORDING_VAD_WAITING) \
                   if self.settings.command_mode else COLOR_STATUS_RECORDING_CONTINUOUS
        elif is_trans: return COLOR_STATUS_TRANSCRIBING
        else: return COLOR_STATUS_IDLE_NOT_RECORDING

    def _get_current_alt_indicator_icon_key(self) -> str:
        is_rec, is_vad_speak = self.audio_service.is_recording_active, self.audio_service.is_vad_speaking
        is_trans = self._ui_transcribing_active
        if is_rec and is_trans: return "rec_and_transcribing"
        elif is_rec:
            return ("recording_on" if is_vad_speak else "recording_vad_wait") \
                   if self.settings.command_mode else "recording_on"
        elif is_trans: return "transcribing"
        else: return "idle"

    def _update_all_status_indicators(self):
        if self.is_shutting_down: return
        current_color = self._get_current_status_indicator_color()
        current_icon_key = self._get_current_alt_indicator_icon_key()
        if self.status_bar_win_manager and self.status_bar_win_manager.enabled:
            self.root.after(0, self.status_bar_win_manager.update_bar_color, current_color)
        if self.alt_status_indicator and self.alt_status_indicator.enabled:
            self.root.after(0, self.alt_status_indicator.update_icon_by_state, current_icon_key)

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
        print("DPI awareness set.") # Not logged as logger not set yet
    except Exception as e:
        print(f"Could not set DPI awareness (non-Windows or error): {e}")

    root = tk.Tk()
    app = None
    try:
        app = WhisperRApp(root)
        # log_essential("Application mainloop starting.") # Already logged by app init
        root.mainloop()
    except KeyboardInterrupt:
        if app: log_essential("KeyboardInterrupt detected, shutting down..."); app._action_quit_application()
        else: print("KeyboardInterrupt before app fully initialized.", file=sys.stderr); sys.exit(1)
    except Exception as main_loop_e:
        # Use a basic print for this very critical error as logger itself might be compromised
        # or the error might be in Tkinter's C code.
        print(f"FATAL ERROR in main application: {main_loop_e}\n{traceback.format_exc()}", file=sys.stderr)
        if app and hasattr(app, '_action_quit_application'):
            try: app._action_quit_application()
            except Exception as shutdown_e:
                print(f"Error during emergency shutdown: {shutdown_e}\n{traceback.format_exc()}", file=sys.stderr)
                os._exit(1)
        else:
            try:
                root_fatal = tk.Tk(); root_fatal.withdraw()
                messagebox.showerror("Fatal Error", f"A critical error occurred: {main_loop_e}\nThe application will now close.")
                root_fatal.destroy()
            except: pass # Ignore if even messagebox fails
            os._exit(1)
    finally:
        # Ensure essential is available or use print
        if app_logger.LOGGER: log_essential("Application exiting.")
        else: print("Application exiting.", file=sys.stderr)

        if app_logger.LOGGER and hasattr(app_logger.LOGGER, 'log_file_handle') and app_logger.LOGGER.log_file_handle:
            app_logger.LOGGER.close()
        sys.exit(0)

if __name__ == "__main__":
    main()
