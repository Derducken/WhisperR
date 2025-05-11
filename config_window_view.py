import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import audio_service
from pathlib import Path
from typing import Callable, Optional, List, Dict, Any, Tuple # Ensure all are here
from app_logger import get_logger, log_debug, log_error, log_extended, log_essential # Import helpers
from settings_manager import AppSettings
from theme_manager import ThemeManager
from ui_components import ConfigSection, create_browse_row
from constants import (
    LOG_LEVELS, CLOSE_BEHAVIORS, STATUS_BAR_POSITIONS, WHISPER_ENGINES,
    AUDIO_FORMATS, AUDIO_FORMAT_TOOLTIPS,
    MIN_MAX_MEMORY_SEGMENT_DURATION, MAX_MAX_MEMORY_SEGMENT_DURATION,
    UI_THEMES, ALT_INDICATOR_POSITIONS, MIN_ALT_INDICATOR_SIZE, MAX_ALT_INDICATOR_SIZE,
    MIN_ALT_INDICATOR_OFFSET, MAX_ALT_INDICATOR_OFFSET, CLI_MODEL_OPTIONS
)

class ConfigWindowView(tk.Toplevel):
    def __init__(self, tk_parent, app_instance, 
                 settings: AppSettings,
                 theme_manager: ThemeManager,
                 audio_devices_list: List[Tuple[int, str, str]],
                 save_config_callback: Callable[[AppSettings], bool],
                 record_hotkey_callback: Callable[[], Optional[str]], # This type hint is for the function passed from main_app
                 vad_calibrate_callback: Callable[[int], Optional[int]],
                 open_command_editor_callback: Callable,
                 delete_session_files_callback: Callable):
        
        super().__init__(tk_parent)
        self.app_instance = app_instance # Store app_instance to access hotkey_manager
        self.settings = settings 
        self.initial_settings = AppSettings(**vars(settings)) 
        self.theme_manager = theme_manager
        self.current_theme_colors = theme_manager.get_current_colors(tk_parent, settings.ui_theme)

        self.audio_devices_list = audio_devices_list
        self.save_config_callback = save_config_callback
        # The actual callback passed from main_app will be a lambda calling app_instance.hotkey_manager.record_new_hotkey_dialog_managed
        self.record_hotkey_callback = record_hotkey_callback 
        self.vad_calibrate_callback = vad_calibrate_callback
        self.open_command_editor_callback = open_command_editor_callback
        self.delete_session_files_callback = delete_session_files_callback

        self.title("WhisperR Configuration")
        self.geometry("650x750")
        self.transient(tk_parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close_button)

        # Tkinter Variables
        self.whisper_executable_var = tk.StringVar(value=str(Path(self.settings.whisper_executable)))
        self.export_folder_var = tk.StringVar(value=str(Path(self.settings.export_folder)))
        self.hotkey_toggle_record_var = tk.StringVar(value=self.settings.hotkey_toggle_record)
        self.hotkey_show_window_var = tk.StringVar(value=self.settings.hotkey_show_window)
        self.hotkey_push_to_talk_var = tk.StringVar(value=self.settings.hotkey_push_to_talk) 
        self.selected_audio_device_var = tk.StringVar()
        self.silence_duration_var = tk.DoubleVar(value=self.settings.silence_threshold_seconds)
        self.vad_energy_var = tk.IntVar(value=self.settings.vad_energy_threshold)
        self.max_memory_segment_duration_var = tk.IntVar(value=self.settings.max_memory_segment_duration_seconds)
        self.audio_segment_format_var = tk.StringVar(value=self.settings.audio_segment_format)
        self.beep_on_save_var = tk.BooleanVar(value=self.settings.beep_on_save_audio_segment)
        self.beep_on_transcription_var = tk.BooleanVar(value=self.settings.beep_on_transcription)
        self.whisper_cli_beeps_var = tk.BooleanVar(value=self.settings.whisper_cli_beeps_enabled)
        self.auto_paste_var = tk.BooleanVar(value=self.settings.auto_paste)
        self.auto_paste_delay_var = tk.DoubleVar(value=self.settings.auto_paste_delay)
        self.status_bar_enabled_var = tk.BooleanVar(value=self.settings.status_bar_enabled)
        self.status_bar_pos_var = tk.StringVar(value=self.settings.status_bar_position)
        self.status_bar_size_var = tk.IntVar(value=self.settings.status_bar_size)
        self.alt_status_indicator_enabled_var = tk.BooleanVar(value=self.settings.alt_status_indicator_enabled)
        self.alt_status_indicator_pos_var = tk.StringVar(value=self.settings.alt_status_indicator_position)
        self.alt_status_indicator_size_var = tk.IntVar(value=self.settings.alt_status_indicator_size)
        self.alt_status_indicator_offset_var = tk.IntVar(value=self.settings.alt_status_indicator_offset)
        self.logging_level_var = tk.StringVar(value=self.settings.logging_level)
        self.log_to_file_var = tk.BooleanVar(value=self.settings.log_to_file)
        self.versioning_var = tk.BooleanVar(value=self.settings.versioning_enabled)
        self.backup_folder_var = tk.StringVar(value=str(Path(self.settings.backup_folder)))
        self.max_backups_var = tk.IntVar(value=self.settings.max_backups)
        self.clear_audio_on_exit_var = tk.BooleanVar(value=self.settings.clear_audio_on_exit)
        self.clear_text_on_exit_var = tk.BooleanVar(value=self.settings.clear_text_on_exit)
        self.close_behavior_var = tk.StringVar(value=self.settings.close_behavior)
        self.ui_theme_var = tk.StringVar(value=self.settings.ui_theme)
        self.audio_format_tooltip_var = tk.StringVar() 
        self.max_log_files_var = tk.IntVar(value=self.settings.max_log_files) 
        self.auto_add_space_var = tk.BooleanVar(value=self.settings.auto_add_space)

        self._apply_theme()
        self._create_widgets()
        self._populate_audio_devices()
        self._on_whisper_engine_change() 
        self._show_audio_format_tooltip()

    def _apply_theme(self):
        self.configure(bg=self.current_theme_colors["bg"])

    def _create_widgets(self):
        main_notebook = ttk.Notebook(self, style='TNotebook', padding=(5,5))
        main_notebook.pack(expand=True, fill='both', padx=5, pady=5)

        tab_general = ttk.Frame(main_notebook, padding=(10, 10, 10, 0), style='TFrame')
        tab_audio = ttk.Frame(main_notebook, padding=(10, 10, 10, 0), style='TFrame')
        tab_notifications = ttk.Frame(main_notebook, padding=(10, 10, 10, 0), style='TFrame')
        tab_status_indication = ttk.Frame(main_notebook, padding=(10,10,10,0), style='TFrame')
        tab_advanced = ttk.Frame(main_notebook, padding=(10, 10, 10, 0), style='TFrame')

        main_notebook.add(tab_general, text=' General & Hotkeys ')
        main_notebook.add(tab_audio, text=' Audio & VAD ')
        main_notebook.add(tab_notifications, text=' Notifications & Output ')
        main_notebook.add(tab_status_indication, text=' Status Indicators ') 
        main_notebook.add(tab_advanced, text=' Advanced & App Behavior ') 
        
        self._create_general_tab(tab_general)
        self._create_audio_tab(tab_audio)
        self._create_notifications_tab(tab_notifications)
        self._create_status_indication_tab(tab_status_indication) 
        self._create_advanced_tab(tab_advanced) 
        
        bottom_button_frame = ttk.Frame(self, style='TFrame', padding=(10,10))
        bottom_button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(bottom_button_frame, text="Save Configuration", command=self._save_configuration_and_close, style='TButton').pack(side=tk.RIGHT, padx=(5,0))
        ttk.Button(bottom_button_frame, text="Cancel", command=self._on_close_button, style='TButton').pack(side=tk.RIGHT)
        ttk.Button(bottom_button_frame, text="Configure Commands...", command=self.open_command_editor_callback, style='TButton').pack(side=tk.LEFT, padx=(0,5))

    def _create_general_tab(self, parent_tab: ttk.Frame):
        sec_engine = ConfigSection(parent_tab, "Whisper Executable", self.theme_manager) 
        engine_frame = sec_engine.get_inner_frame()
        
        self.exec_path_label = ttk.Label(engine_frame, text="Executable Path:", style='TLabel')
        self.exec_path_label.grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=2)
        self.exec_path_entry_frame = ttk.Frame(engine_frame, style='TFrame')
        self.exec_path_entry_frame.grid(row=0, column=1, sticky=tk.EW, pady=2) 
        self.exec_path_entry = ttk.Entry(self.exec_path_entry_frame, textvariable=self.whisper_executable_var)
        self.exec_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        self.exec_path_browse_btn = ttk.Button(self.exec_path_entry_frame, text="Browse...", command=self._browse_whisper_executable, style='TButton')
        self.exec_path_browse_btn.pack(side=tk.LEFT)
        engine_frame.columnconfigure(1, weight=1)

        sec_export = ConfigSection(parent_tab, "File Export", self.theme_manager)
        export_frame = sec_export.get_inner_frame()
        create_browse_row(export_frame, "Export Folder (Audio/Text):", self.export_folder_var, self._browse_export_folder)

        sec_hotkeys = ConfigSection(parent_tab, "Global Hotkeys", self.theme_manager)
        hotkey_frame = sec_hotkeys.get_inner_frame()
        ttk.Label(hotkey_frame, text="Toggle Recording:", style='TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=2)
        toggle_hk_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_toggle_record_var, width=25)
        toggle_hk_entry.grid(row=0, column=1, sticky=tk.EW, pady=2, padx=(0,5))
        ttk.Button(hotkey_frame, text="Record", command=lambda: self._record_hotkey_ui(self.hotkey_toggle_record_var), style='TButton').grid(row=0, column=2, pady=2)
        ttk.Label(hotkey_frame, text="Show Window:", style='TLabel').grid(row=1, column=0, sticky=tk.W, padx=(0,5), pady=2)
        show_hk_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_show_window_var, width=25)
        show_hk_entry.grid(row=1, column=1, sticky=tk.EW, pady=2, padx=(0,5))
        ttk.Button(hotkey_frame, text="Record", command=lambda: self._record_hotkey_ui(self.hotkey_show_window_var), style='TButton').grid(row=1, column=2, pady=2)
        
        ttk.Label(hotkey_frame, text="Push To Talk:", style='TLabel').grid(row=2, column=0, sticky=tk.W, padx=(0,5), pady=2)
        ptt_hk_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_push_to_talk_var, width=25)
        ptt_hk_entry.grid(row=2, column=1, sticky=tk.EW, pady=2, padx=(0,5))
        ttk.Button(hotkey_frame, text="Record", command=lambda: self._record_hotkey_ui(self.hotkey_push_to_talk_var), style='TButton').grid(row=2, column=2, pady=2)

        hotkey_frame.columnconfigure(1, weight=1)
        ttk.Label(hotkey_frame, text="Press 'Record' then type combination. (e.g., ctrl+alt+space)", style='TLabel', wraplength=400).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5,0))

    def _create_audio_tab(self, parent_tab: ttk.Frame):
        sec_input = ConfigSection(parent_tab, "Audio Input Device", self.theme_manager)
        input_frame = sec_input.get_inner_frame()
        ttk.Label(input_frame, text="Device:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        self.audio_device_combobox = ttk.Combobox(input_frame, textvariable=self.selected_audio_device_var,
                                                  state="readonly", width=50)
        self.audio_device_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        sec_vad = ConfigSection(parent_tab, "Auto-Pause (VAD) Settings", self.theme_manager)
        vad_frame = sec_vad.get_inner_frame()
        vad_grid = ttk.Frame(vad_frame, style='TFrame')
        vad_grid.pack(fill=tk.X)
        ttk.Label(vad_grid, text="Silence Duration (s):", style='TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=3)
        ttk.Entry(vad_grid, textvariable=self.silence_duration_var, width=7).grid(row=0, column=1, sticky=tk.W, pady=3)
        ttk.Label(vad_grid, text="Energy Threshold:", style='TLabel').grid(row=1, column=0, sticky=tk.W, padx=(0,5), pady=3)
        vad_energy_entry = ttk.Entry(vad_grid, textvariable=self.vad_energy_var, width=7)
        vad_energy_entry.grid(row=1, column=1, sticky=tk.W, pady=3)
        ttk.Button(vad_grid, text="Calibrate...", command=self._calibrate_vad_ui, style='TButton').grid(row=1, column=2, sticky=tk.W, padx=(10,0), pady=3)
        
        sec_segment = ConfigSection(parent_tab, "Audio Segment Handling", self.theme_manager)
        segment_frame = sec_segment.get_inner_frame()
        segment_grid = ttk.Frame(segment_frame, style='TFrame')
        segment_grid.pack(fill=tk.X)
        ttk.Label(segment_grid, text="Max In-Memory Duration (s):", style='TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=3)
        self.max_mem_spinbox = ttk.Spinbox(segment_grid, textvariable=self.max_memory_segment_duration_var,
                                            from_=MIN_MAX_MEMORY_SEGMENT_DURATION, to=MAX_MAX_MEMORY_SEGMENT_DURATION,
                                            increment=10, width=7, wrap=True)
        self.max_mem_spinbox.grid(row=0, column=1, sticky=tk.W, pady=3)
        ttk.Label(segment_grid, text=f"(Range: {MIN_MAX_MEMORY_SEGMENT_DURATION}-{MAX_MAX_MEMORY_SEGMENT_DURATION}s)", style='TLabel').grid(row=0, column=2, sticky=tk.W, padx=(5,0))
        ttk.Label(segment_grid, text="Saved Audio Format:", style='TLabel').grid(row=1, column=0, sticky=tk.W, padx=(0,5), pady=3)
        format_combo = ttk.Combobox(segment_grid, textvariable=self.audio_segment_format_var,
                                    values=AUDIO_FORMATS, state="readonly", width=7)
        format_combo.grid(row=1, column=1, sticky=tk.W, pady=3)
        format_combo.bind("<<ComboboxSelected>>", self._show_audio_format_tooltip)
        ttk.Label(segment_grid, textvariable=self.audio_format_tooltip_var, style='TLabel', wraplength=300).grid(row=1, column=2, sticky=tk.W, padx=(5,0), columnspan=2)

    def _create_notifications_tab(self, parent_tab: ttk.Frame):
        # Section 1: Transcription Feedback
        sec_feedback = ConfigSection(parent_tab, "Transcription Feedback", self.theme_manager)
        feedback_frame = sec_feedback.get_inner_frame()
        
        ttk.Checkbutton(feedback_frame, text="Beep on Audio Segment Save (.wav/.mp3/.aac created)", 
                        variable=self.beep_on_save_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(feedback_frame, text="Beep on Transcription Completion", 
                        variable=self.beep_on_transcription_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        self.cli_beep_checkbox = ttk.Checkbutton(feedback_frame, 
                                                 text="Enable Whisper CLI's internal beeps (removes --beep_off flag from CLI engine)",
                                                 variable=self.whisper_cli_beeps_var, style='TCheckbutton')
        self.cli_beep_checkbox.pack(anchor=tk.W, pady=(5,2))

        # Section 2: Output Actions
        sec_actions = ConfigSection(parent_tab, "Output Actions", self.theme_manager)
        actions_frame = sec_actions.get_inner_frame()

        # Auto-Paste options
        auto_paste_sub_frame = ttk.Frame(actions_frame, style='TFrame') 
        auto_paste_sub_frame.pack(fill=tk.X, anchor=tk.W, pady=(0, 5)) 
        ttk.Checkbutton(auto_paste_sub_frame, text="Auto-Paste After Transcription", 
                        variable=self.auto_paste_var, style='TCheckbutton').pack(side=tk.LEFT, anchor=tk.W, padx=(0,10))
        ttk.Label(auto_paste_sub_frame, text="Delay (s):", style='TLabel').pack(side=tk.LEFT, padx=(5,5))
        ttk.Entry(auto_paste_sub_frame, textvariable=self.auto_paste_delay_var, width=5).pack(side=tk.LEFT)
        
        # Auto-Add Space option (directly in actions_frame, below auto_paste_sub_frame)
        ttk.Checkbutton(actions_frame, text="Auto-Add Space After Transcription", 
                        variable=self.auto_add_space_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)

    def _create_status_indication_tab(self, parent_tab: ttk.Frame): 
        sec_edge_bar = ConfigSection(parent_tab, "Screen Edge Status Bar (Windows Only)", self.theme_manager)
        edge_bar_frame = sec_edge_bar.get_inner_frame()
        ttk.Checkbutton(edge_bar_frame, text="Enable Screen Edge Status Bar", variable=self.status_bar_enabled_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        sb_pos_frame = ttk.Frame(edge_bar_frame, style='TFrame', padding=(20,5,0,0))
        sb_pos_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(sb_pos_frame, text="Position:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(sb_pos_frame, textvariable=self.status_bar_pos_var, values=STATUS_BAR_POSITIONS,
                     state="readonly", width=10).pack(side=tk.LEFT)
        sb_size_frame = ttk.Frame(edge_bar_frame, style='TFrame', padding=(20,5,0,0))
        sb_size_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(sb_size_frame, text="Size (pixels):", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Spinbox(sb_size_frame, textvariable=self.status_bar_size_var, from_=1, to=50, increment=1, width=5).pack(side=tk.LEFT)
        ttk.Label(sb_size_frame, text="(Height for Top/Bottom, Width for Left/Right)", style='TLabel').pack(side=tk.LEFT, padx=5)
        ttk.Label(edge_bar_frame, text="This bar is click-through and aims to be minimally intrusive.", style='TLabel', foreground="gray", wraplength=500).pack(anchor=tk.W, pady=(10,0), padx=20)

        sec_alt_icon = ConfigSection(parent_tab, "Corner Status Icon (Cross-Platform)", self.theme_manager)
        alt_icon_frame = sec_alt_icon.get_inner_frame()
        ttk.Checkbutton(alt_icon_frame, text="Enable Corner Status Icon", variable=self.alt_status_indicator_enabled_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        alt_options_frame = ttk.Frame(alt_icon_frame, style='TFrame', padding=(20,5,0,0))
        alt_options_frame.pack(fill=tk.X, anchor=tk.W)
        alt_pos_frame = ttk.Frame(alt_options_frame, style='TFrame')
        alt_pos_frame.pack(fill=tk.X, pady=2)
        ttk.Label(alt_pos_frame, text="Position:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(alt_pos_frame, textvariable=self.alt_status_indicator_pos_var, values=ALT_INDICATOR_POSITIONS,
                     state="readonly", width=15).pack(side=tk.LEFT)
        alt_size_frame = ttk.Frame(alt_options_frame, style='TFrame')
        alt_size_frame.pack(fill=tk.X, pady=2)
        ttk.Label(alt_size_frame, text="Size (px):", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Spinbox(alt_size_frame, textvariable=self.alt_status_indicator_size_var, 
                    from_=MIN_ALT_INDICATOR_SIZE, to=MAX_ALT_INDICATOR_SIZE, increment=4, width=7).pack(side=tk.LEFT)
        alt_offset_frame = ttk.Frame(alt_options_frame, style='TFrame')
        alt_offset_frame.pack(fill=tk.X, pady=2)
        ttk.Label(alt_offset_frame, text="Offset from Corner (px):", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Spinbox(alt_offset_frame, textvariable=self.alt_status_indicator_offset_var,
                    from_=MIN_ALT_INDICATOR_OFFSET, to=MAX_ALT_INDICATOR_OFFSET, increment=2, width=7).pack(side=tk.LEFT)
        ttk.Label(alt_icon_frame, text="Uses custom icons from 'status_icons' folder. Transparency depends on OS/WM.", style='TLabel', foreground="gray", wraplength=500).pack(anchor=tk.W, pady=(10,0), padx=20)

    def _create_advanced_tab(self, parent_tab: ttk.Frame): 
        sec_logging = ConfigSection(parent_tab, "Logging", self.theme_manager)
        log_frame = sec_logging.get_inner_frame()
        log_level_frame = ttk.Frame(log_frame, style='TFrame')
        log_level_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(log_level_frame, text="Logging Level:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(log_level_frame, textvariable=self.logging_level_var, values=LOG_LEVELS,
                     state="readonly", width=12).pack(side=tk.LEFT, padx=(0,15))
        ttk.Checkbutton(log_level_frame, text="Log to File", variable=self.log_to_file_var, style='TCheckbutton').pack(side=tk.LEFT)

        log_file_management_frame = ttk.Frame(log_frame, style='TFrame', padding=(0,5,0,0))
        log_file_management_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(log_file_management_frame, text="Max Log Files to Keep:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Spinbox(log_file_management_frame, textvariable=self.max_log_files_var, from_=1, to=100, increment=1, width=5).pack(side=tk.LEFT)
        ttk.Label(log_file_management_frame, text="(1-100, 0 for unlimited)", style='TLabel').pack(side=tk.LEFT, padx=(5,0)) 
        
        sec_versioning = ConfigSection(parent_tab, "File Versioning (Backups)", self.theme_manager)
        ver_frame = sec_versioning.get_inner_frame()
        ttk.Checkbutton(ver_frame, text="Enable Backups for Config/Prompt/Commands", variable=self.versioning_var, style='TCheckbutton').pack(anchor=tk.W, pady=(0,5))
        create_browse_row(ver_frame, "Backup Folder:", self.backup_folder_var, self._browse_backup_folder)
        max_backups_frame = ttk.Frame(ver_frame, style='TFrame', padding=(0,5,0,0))
        max_backups_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(max_backups_frame, text="Max Backups per File Type:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Entry(max_backups_frame, textvariable=self.max_backups_var, width=5).pack(side=tk.LEFT)

        sec_behavior = ConfigSection(parent_tab, "Application Behavior", self.theme_manager)
        behav_frame = sec_behavior.get_inner_frame()
        ttk.Checkbutton(behav_frame, text="Clear Audio (.wav/.mp3/.transcribed) on Application Exit", variable=self.clear_audio_on_exit_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(behav_frame, text="Clear Text (.txt) on Application Exit", variable=self.clear_text_on_exit_var, style='TCheckbutton').pack(anchor=tk.W, pady=2)
        close_behav_frame = ttk.Frame(behav_frame, style='TFrame', padding=(0,5,0,0))
        close_behav_frame.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(close_behav_frame, text="Window Close (X) Action:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(close_behav_frame, textvariable=self.close_behavior_var, values=CLOSE_BEHAVIORS,
                     state="readonly", width=20).pack(side=tk.LEFT)

        sec_theme = ConfigSection(parent_tab, "User Interface Theme", self.theme_manager)
        theme_frame = sec_theme.get_inner_frame()
        ttk.Label(theme_frame, text="Select Theme:", style='TLabel').pack(side=tk.LEFT, padx=(0,5))
        theme_combo = ttk.Combobox(theme_frame, textvariable=self.ui_theme_var, values=UI_THEMES,
                                   state="readonly", width=15)
        theme_combo.pack(side=tk.LEFT)
        ttk.Label(theme_frame, text="(Restart may be needed for full effect on some elements)", style='TLabel', foreground="gray").pack(side=tk.LEFT, padx=10)

        delete_button_frame = ttk.Frame(parent_tab, style='TFrame', padding=(5,20,5,5))
        delete_button_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(delete_button_frame, text="Delete Session Files Now...", command=self.delete_session_files_callback, style='TButton').pack(fill=tk.X, ipady=5)

    def _browse_file(self, title: str, filetypes: List[Tuple[str,str]], entry_var: tk.StringVar, is_folder: bool = False):
        initial_dir_str = entry_var.get()
        initial_dir = None
        if initial_dir_str:
            path_obj = Path(initial_dir_str)
            if path_obj.is_dir():
                initial_dir = str(path_obj)
            elif path_obj.parent.exists():
                initial_dir = str(path_obj.parent)
        
        if not initial_dir and self.settings.export_folder: 
            initial_dir = self.settings.export_folder
        if not initial_dir: initial_dir = "."


        if is_folder:
            selected_path = filedialog.askdirectory(title=title, initialdir=initial_dir, parent=self)
        else:
            selected_path = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initial_dir, parent=self)
        if selected_path: entry_var.set(str(Path(selected_path)))

    def _browse_whisper_executable(self): self._browse_file("Select Whisper Executable", [("Executables", "*.exe"), ("All files", "*.*")], self.whisper_executable_var)
    def _browse_export_folder(self): self._browse_file("Select Export Folder", [], self.export_folder_var, is_folder=True)
    def _browse_backup_folder(self): self._browse_file("Select Backup Folder", [], self.backup_folder_var, is_folder=True)

    def _record_hotkey_ui(self, target_var: tk.StringVar):
        self.focus_set(); self.update_idletasks()
        # This now correctly calls the lambda which calls app_instance.hotkey_manager.record_new_hotkey_dialog_managed
        recorded_hotkey = self.record_hotkey_callback() 
        if recorded_hotkey is not None: target_var.set(recorded_hotkey)
        self.lift()

    def _populate_audio_devices(self):
        display_names = [name for _, name, _ in self.audio_devices_list]
        if not display_names or (len(display_names)==1 and display_names[0] in ["Error querying devices", "No input devices found"]):
            self.audio_device_combobox['values'] = ["No input devices found"]
            self.selected_audio_device_var.set("No input devices found")
            self.audio_device_combobox.config(state=tk.DISABLED)
            return
        
        self.audio_device_combobox.config(state="readonly") 
        self.audio_device_combobox['values'] = display_names
        current_idx = self.settings.selected_audio_device_index
        selected_display_name = None
        if current_idx is not None:
            for idx, disp_name, _ in self.audio_devices_list:
                if idx == current_idx: selected_display_name = disp_name; break
        if selected_display_name: self.selected_audio_device_var.set(selected_display_name)
        elif display_names: self.selected_audio_device_var.set(display_names[0])

    def _calibrate_vad_ui(self):
        self.focus_set(); self.update_idletasks()
        current_threshold = self.vad_energy_var.get()
        recommended = self.vad_calibrate_callback(current_threshold)
        if recommended is not None:
            self.vad_energy_var.set(recommended)
            messagebox.showinfo("VAD Calibration", f"VAD Threshold updated to: {recommended}", parent=self)
        self.lift()

    def _show_audio_format_tooltip(self, event=None):
        selected_format = self.audio_segment_format_var.get()
        tooltip_text = AUDIO_FORMAT_TOOLTIPS.get(selected_format, "Select an audio format.")
        if selected_format in ["MP3", "AAC"]:
            if not audio_service.PYDUB_AVAILABLE:
                tooltip_text += "\n(Requires pydub library and FFmpeg in PATH)"
        self.audio_format_tooltip_var.set(tooltip_text)

    def _on_whisper_engine_change(self, *args):
        is_executable_engine = True 
        if hasattr(self, 'exec_path_label') and hasattr(self, 'exec_path_entry_frame'):
            for widget in [self.exec_path_label, self.exec_path_entry_frame]: 
                if widget: widget.grid() 
        if hasattr(self, 'fw_model_label') and hasattr(self, 'fw_model_entry_frame'):
            for widget in [self.fw_model_label, self.fw_model_entry_frame]: 
                if widget: widget.grid_remove()
        if hasattr(self, 'cli_beep_checkbox') and self.cli_beep_checkbox: 
            self.cli_beep_checkbox.config(state=tk.NORMAL) 
        if hasattr(self, 'fw_warning_label') and self.fw_warning_label: 
            self.fw_warning_label.grid_remove()

    def _collect_settings_from_ui(self) -> AppSettings:
        s = AppSettings() 
        initial_s = self.initial_settings 

        s.whisper_engine_type = WHISPER_ENGINES[0] 
        s.whisper_executable = self.whisper_executable_var.get() or initial_s.whisper_executable
        s.export_folder = self.export_folder_var.get() or initial_s.export_folder
        s.hotkey_toggle_record = self.hotkey_toggle_record_var.get() or initial_s.hotkey_toggle_record
        s.hotkey_show_window = self.hotkey_show_window_var.get() or initial_s.hotkey_show_window
        s.hotkey_push_to_talk = self.hotkey_push_to_talk_var.get() or initial_s.hotkey_push_to_talk
        
        selected_device_display_name = self.selected_audio_device_var.get()
        s.selected_audio_device_index = None
        for idx, disp_name, _ in self.audio_devices_list:
            if disp_name == selected_device_display_name: s.selected_audio_device_index = idx; break
        if selected_device_display_name in ["Error querying devices", "No input devices found", ""]: s.selected_audio_device_index = initial_s.selected_audio_device_index

        try: s.silence_threshold_seconds = float(self.silence_duration_var.get())
        except ValueError: s.silence_threshold_seconds = initial_s.silence_threshold_seconds
        try: s.vad_energy_threshold = int(self.vad_energy_var.get())
        except ValueError: s.vad_energy_threshold = initial_s.vad_energy_threshold
        try: s.max_memory_segment_duration_seconds = int(self.max_memory_segment_duration_var.get())
        except ValueError: s.max_memory_segment_duration_seconds = initial_s.max_memory_segment_duration_seconds
        s.audio_segment_format = self.audio_segment_format_var.get() or initial_s.audio_segment_format
        
        s.beep_on_save_audio_segment = self.beep_on_save_var.get()
        s.beep_on_transcription = self.beep_on_transcription_var.get()
        s.whisper_cli_beeps_enabled = self.whisper_cli_beeps_var.get()
        s.auto_paste = self.auto_paste_var.get()
        try: s.auto_paste_delay = float(self.auto_paste_delay_var.get())
        except ValueError: s.auto_paste_delay = initial_s.auto_paste_delay

        s.status_bar_enabled = self.status_bar_enabled_var.get()
        s.status_bar_position = self.status_bar_pos_var.get() or initial_s.status_bar_position
        try: s.status_bar_size = int(self.status_bar_size_var.get())
        except ValueError: s.status_bar_size = initial_s.status_bar_size
        s.alt_status_indicator_enabled = self.alt_status_indicator_enabled_var.get()
        s.alt_status_indicator_position = self.alt_status_indicator_pos_var.get() or initial_s.alt_status_indicator_position
        try: s.alt_status_indicator_size = int(self.alt_status_indicator_size_var.get())
        except ValueError: s.alt_status_indicator_size = initial_s.alt_status_indicator_size
        try: s.alt_status_indicator_offset = int(self.alt_status_indicator_offset_var.get())
        except ValueError: s.alt_status_indicator_offset = initial_s.alt_status_indicator_offset

        s.logging_level = self.logging_level_var.get() or initial_s.logging_level
        s.log_to_file = self.log_to_file_var.get()
        s.versioning_enabled = self.versioning_var.get()
        s.backup_folder = self.backup_folder_var.get() or initial_s.backup_folder
        try: s.max_backups = int(self.max_backups_var.get())
        except ValueError: s.max_backups = initial_s.max_backups
        s.clear_audio_on_exit = self.clear_audio_on_exit_var.get()
        s.clear_text_on_exit = self.clear_text_on_exit_var.get()
        s.close_behavior = self.close_behavior_var.get() or initial_s.close_behavior
        s.ui_theme = self.ui_theme_var.get() or initial_s.ui_theme
        try: s.max_log_files = int(self.max_log_files_var.get())
        except ValueError: s.max_log_files = initial_s.max_log_files
        s.auto_add_space = self.auto_add_space_var.get()

        s.language = initial_s.language
        s.model = initial_s.model
        s.translation_enabled = initial_s.translation_enabled
        s.command_mode = initial_s.command_mode
        s.timestamps_disabled = initial_s.timestamps_disabled
        s.clear_text_output = initial_s.clear_text_output
        s.scratchpad_append_mode = initial_s.scratchpad_append_mode

        return s

    def _has_changes(self) -> bool:
        current_ui_settings = self._collect_settings_from_ui()
        initial_dict = vars(self.initial_settings)
        current_dict = vars(current_ui_settings)
        for key, initial_value in initial_dict.items():
            if key == "prompt": continue 
            current_value = current_dict.get(key) 
            if isinstance(initial_value, str) and (key.endswith("folder") or key.endswith("executable")): 
                if Path(current_value or "").resolve() != Path(initial_value or "").resolve():
                    log_debug(f"Config changed (path): {key} from '{initial_value}' to '{current_value}'")
                    return True
            elif current_value != initial_value:
                log_debug(f"Config changed: {key} from '{initial_value}' to '{current_value}'")
                return True
        return False

    def _save_configuration_and_close(self):
        new_settings = self._collect_settings_from_ui()
        if new_settings.silence_threshold_seconds < 0: new_settings.silence_threshold_seconds = self.initial_settings.silence_threshold_seconds
        if new_settings.vad_energy_threshold < 0: new_settings.vad_energy_threshold = self.initial_settings.vad_energy_threshold
        if not (MIN_MAX_MEMORY_SEGMENT_DURATION <= new_settings.max_memory_segment_duration_seconds <= MAX_MAX_MEMORY_SEGMENT_DURATION):
            new_settings.max_memory_segment_duration_seconds = self.initial_settings.max_memory_segment_duration_seconds
        if new_settings.auto_paste_delay < 0: new_settings.auto_paste_delay = self.initial_settings.auto_paste_delay
        if not (1 <= new_settings.status_bar_size <= 100): new_settings.status_bar_size = self.initial_settings.status_bar_size
        if not (MIN_ALT_INDICATOR_SIZE <= new_settings.alt_status_indicator_size <= MAX_ALT_INDICATOR_SIZE):
             new_settings.alt_status_indicator_size = self.initial_settings.alt_status_indicator_size
        if not (MIN_ALT_INDICATOR_OFFSET <= new_settings.alt_status_indicator_offset <= MAX_ALT_INDICATOR_OFFSET):
            new_settings.alt_status_indicator_offset = self.initial_settings.alt_status_indicator_offset
        if new_settings.max_backups < 0: new_settings.max_backups = self.initial_settings.max_backups
        if new_settings.max_log_files < 0: new_settings.max_log_files = self.initial_settings.max_log_files 

        hotkeys_ok = self.save_config_callback(new_settings)
        if hotkeys_ok:
            log_essential("Configuration saved from dialog.")
            self.destroy()
        else:
            messagebox.showerror("Hotkey Error", "One or more hotkeys could not be registered. Please check syntax and try again. Other settings were saved.", parent=self)

    def _on_close_button(self):
        if self._has_changes():
            response = messagebox.askyesnocancel("Unsaved Changes", "You have unsaved changes. Save them before closing?", parent=self)
            if response is True: self._save_configuration_and_close()
            elif response is False: self.destroy()
        else: self.destroy()
