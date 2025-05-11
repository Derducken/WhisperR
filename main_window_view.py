import tkinter as tk
from tkinter import ttk, Menu
from typing import Callable, List, Tuple, Optional # Ensure Optional and Callable are here
from app_logger import get_logger, log_extended, log_error # Ensure log_error is imported
from constants import DEFAULT_LANGUAGE, DEFAULT_MODEL, EXTENDED_MODEL_OPTIONS, Theme
from settings_manager import AppSettings, SettingsManager # Import SettingsManager
from theme_manager import ThemeManager

class MainWindowView:
    def __init__(self, root: tk.Tk, settings_manager: SettingsManager, initial_prompt: str, theme_manager: ThemeManager): # Changed settings to settings_manager
        self.root = root
        self.settings_manager = settings_manager # Store settings_manager
        # self.settings is now a property-like access to current settings via settings_manager
        self.initial_prompt_val = initial_prompt
        self.theme_manager = theme_manager

        # Access settings via self.settings_manager.settings for initialization
        current_settings = self.settings_manager.settings
        self.language_var = tk.StringVar(value=current_settings.language)
        self.model_var = tk.StringVar(value=current_settings.model)
        self.translation_var = tk.BooleanVar(value=current_settings.translation_enabled)
        self.command_mode_var = tk.BooleanVar(value=current_settings.command_mode)
        self.timestamps_disabled_var = tk.BooleanVar(value=current_settings.timestamps_disabled)
        self.clear_text_output_var = tk.BooleanVar(value=current_settings.clear_text_output)
        
        self.shortcut_display_var = tk.StringVar()
        self.queue_indicator_var = tk.StringVar(value="Queue: 0")
        self.pause_queue_menu_var = tk.BooleanVar(value=False) 

        self.is_recording_visual_indicator = False 

        self.prompt_text_widget: Optional[tk.Text] = None
        self.start_stop_button: Optional[ttk.Button] = None
        self.recording_indicator_label: Optional[ttk.Label] = None
        self.pause_queue_button: Optional[ttk.Button] = None
        self.model_combobox: Optional[ttk.Combobox] = None
        self.language_combobox: Optional[ttk.Combobox] = None
        self.ok_hide_button: Optional[ttk.Button] = None # New button for OK (Hide Window)

        self._create_widgets()

        if self.prompt_text_widget:
            self.prompt_text_widget.delete("1.0", tk.END) 
            self.prompt_text_widget.insert("1.0", self.initial_prompt_val)
        
        self.update_ui_from_settings() # This will now use self.settings_manager.settings

    def _create_widgets(self):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        self.file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=self.file_menu)
        self.settings_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=self.settings_menu)
        self.queue_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Queue", menu=self.queue_menu)

        top_frame = ttk.Frame(self.root, style='TFrame', padding=(10,10))
        top_frame.pack(fill=tk.X)

        lang_frame = ttk.Frame(top_frame, style='TFrame')
        lang_frame.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(lang_frame, text="Language:", style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.language_options = ["auto", "en", "es", "fr", "de", "it", "ja", "zh", "ko", "ru", "pt", "el"]
        self.language_combobox = ttk.Combobox(lang_frame, textvariable=self.language_var,
                                              values=self.language_options, state="readonly", width=10)
        self.language_combobox.pack(side=tk.LEFT)

        model_frame = ttk.Frame(top_frame, style='TFrame')
        model_frame.pack(side=tk.RIGHT, padx=(0, 0)) # Changed side to RIGHT
        ttk.Label(model_frame, text="Model (CLI/Lib Fallback):", style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.model_combobox = ttk.Combobox(model_frame, textvariable=self.model_var,
                                           values=EXTENDED_MODEL_OPTIONS, width=28)
        self.model_combobox.pack(side=tk.LEFT)

        toggle_frame1 = ttk.Frame(self.root, style='TFrame', padding=(10,5))
        toggle_frame1.pack(fill=tk.X)
        ttk.Checkbutton(toggle_frame1, text="Enable Translation", variable=self.translation_var, style='TCheckbutton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(toggle_frame1, text="Auto-Pause / Commands (VAD)", variable=self.command_mode_var, style='TCheckbutton').pack(side=tk.RIGHT, padx=(0,0)) # Changed side to RIGHT

        toggle_frame2 = ttk.Frame(self.root, style='TFrame', padding=(10,2,10,5))
        toggle_frame2.pack(fill=tk.X)
        ttk.Checkbutton(toggle_frame2, text="Hide Timestamps (Output)", variable=self.timestamps_disabled_var, style='TCheckbutton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(toggle_frame2, text="Clean Metadata (Output)", variable=self.clear_text_output_var, style='TCheckbutton').pack(side=tk.RIGHT, padx=(0,0)) # Changed side to RIGHT

        prompt_label_frame = ttk.Frame(self.root, style='TFrame', padding=(10,10,10,0))
        prompt_label_frame.pack(fill=tk.X)
        ttk.Label(prompt_label_frame, text="Whisper Initial Prompt:", style='TLabel').pack(anchor=tk.W)
        
        self.prompt_text_widget = tk.Text(self.root, height=8, wrap=tk.WORD, undo=True) # Increased height to 8
        self.prompt_text_widget.pack(pady=5, padx=10, fill=tk.X, expand=False)

        self.scratchpad_button = ttk.Button(self.root, text="Open Scratchpad")
        self.scratchpad_button.pack(pady=(5, 10), padx=10, fill=tk.X, ipady=8)

        self.ok_hide_button = ttk.Button(self.root, text="OK (Hide Window)") # New button
        self.ok_hide_button.pack(pady=(0, 10), padx=10, fill=tk.X, ipady=8) # New button packing

        self.start_stop_frame = ttk.Frame(self.root, style='TFrame', padding=(10,0,10,0)) # Reduced bottom padding
        self.start_stop_frame.pack(fill=tk.X)
        self.start_stop_button = ttk.Button(self.start_stop_frame, text="Start Recording")
        self.start_stop_button.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=10)
        self.recording_indicator_label = ttk.Label(self.start_stop_frame, text="●", font=("Arial", 16), style='TLabel')
        self.recording_indicator_label.pack(side=tk.LEFT, padx=10)
        self.update_recording_indicator_ui()

        bottom_controls_frame = ttk.Frame(self.root, style='TFrame', padding=(10,0,10,10)) # Reduced top padding
        bottom_controls_frame.pack(fill=tk.X, side=tk.BOTTOM, anchor=tk.S)
        self.hotkey_display_label = ttk.Label(bottom_controls_frame, textvariable=self.shortcut_display_var,
                                             justify=tk.LEFT, style='TLabel', wraplength=250)
        self.hotkey_display_label.pack(side=tk.LEFT, anchor=tk.W, expand=True, fill=tk.X)
        
        queue_controls_subframe = ttk.Frame(bottom_controls_frame, style='TFrame')
        queue_controls_subframe.pack(side=tk.RIGHT, anchor=tk.E)
        self.queue_indicator_label = ttk.Label(queue_controls_subframe, textvariable=self.queue_indicator_var,
                                               justify=tk.RIGHT, style='TLabel')
        self.queue_indicator_label.pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))
        self.clear_queue_button = ttk.Button(queue_controls_subframe, text="Clear Q", width=8)
        self.clear_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))
        self.pause_queue_button = ttk.Button(queue_controls_subframe, text="Pause Q", width=9)
        self.pause_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))

    def update_ui_from_settings(self):
        current_settings = self.settings_manager.settings # Use current settings
        self.language_var.set(current_settings.language)
        self.model_var.set(current_settings.model)
        self.translation_var.set(current_settings.translation_enabled)
        self.command_mode_var.set(current_settings.command_mode)
        self.timestamps_disabled_var.set(current_settings.timestamps_disabled)
        self.clear_text_output_var.set(current_settings.clear_text_output)
        
        self.update_shortcut_display_ui()
        log_extended("Main window UI updated from settings.")

    def update_recording_indicator_ui(self, is_recording: Optional[bool] = None, vad_is_speaking: Optional[bool] = None):
        if is_recording is not None:
            self.is_recording_visual_indicator = is_recording

        if self.start_stop_button and self.start_stop_button.winfo_exists():
            self.start_stop_button.config(text="Stop Recording" if self.is_recording_visual_indicator else "Start Recording")
        
        if self.recording_indicator_label and self.recording_indicator_label.winfo_exists():
            indicator_color = self.theme_manager.themes[Theme.LIGHT.value]["disabled_fg"] 
            current_settings = self.settings_manager.settings # Use current settings
            try:
                 current_theme_name = current_settings.ui_theme
                 colors = self.theme_manager.get_current_colors(self.root, current_theme_name)
                 indicator_color = colors.get("disabled_fg", indicator_color)
            except Exception: pass

            if self.is_recording_visual_indicator:
                if current_settings.command_mode: # Use current settings
                    indicator_color = "orange" if vad_is_speaking else "darkkhaki" 
                else:
                    indicator_color = "red"
            self.recording_indicator_label.config(foreground=indicator_color)

    def update_shortcut_display_ui(self):
        current_settings = self.settings_manager.settings # Use current settings
        ptt_key = current_settings.hotkey_push_to_talk or "[Not Set]"
        toggle_key = current_settings.hotkey_toggle_record or "[Not Set]"
        show_key = current_settings.hotkey_show_window or "[Not Set]"
        self.shortcut_display_var.set(f"PTT: {ptt_key}\nToggle: {toggle_key}\nShow: {show_key}")

    def update_queue_indicator_ui(self, queue_size: int):
        self.queue_indicator_var.set(f"Queue: {queue_size}")

    def update_pause_queue_button_ui(self, is_paused: bool):
        if self.pause_queue_button and self.pause_queue_button.winfo_exists():
            self.pause_queue_button.config(text="Resume Q" if is_paused else "Pause Q")
        self.pause_queue_menu_var.set(is_paused)

    def get_prompt_text(self) -> str:
        if self.prompt_text_widget and self.prompt_text_widget.winfo_exists():
            return self.prompt_text_widget.get("1.0", tk.END).strip()
        return "" 

    def set_prompt_widget_text(self, text: str):
        if self.prompt_text_widget and self.prompt_text_widget.winfo_exists():
            self.prompt_text_widget.delete("1.0", tk.END)
            self.prompt_text_widget.insert("1.0", text)

    def bind_language_change(self, callback: Callable[[str], None]):
        self.language_combobox.bind("<<ComboboxSelected>>", lambda e: callback(self.language_var.get()))

    def bind_model_change(self, callback: Callable[[str], None]):
        self.model_combobox.bind("<<ComboboxSelected>>", lambda e: callback(self.model_var.get()))
        self.model_combobox.bind("<Return>", lambda e: callback(self.model_var.get()))
        self.model_combobox.bind("<FocusOut>", lambda e: callback(self.model_var.get()))

    def bind_toggle_change(self, var_name: str, callback: Callable[[bool], None]):
        var_map = {
            "translation": self.translation_var,
            "command_mode": self.command_mode_var,
            "timestamps_disabled": self.timestamps_disabled_var,
            "clear_text_output": self.clear_text_output_var,
        }
        if var_name in var_map:
            var_map[var_name].trace_add("write", lambda *args: callback(var_map[var_name].get()))
        else:
            log_error(f"Cannot bind toggle change for unknown var_name: {var_name}")

    def bind_prompt_change(self, callback: Callable[[str], None]):
        if self.prompt_text_widget:
            self._prompt_update_job = None
            def on_key_release(event):
                if self._prompt_update_job:
                    self.root.after_cancel(self._prompt_update_job)
                self._prompt_update_job = self.root.after(750, lambda: callback(self.get_prompt_text()))
            
            self.prompt_text_widget.bind("<KeyRelease>", on_key_release)
            self.prompt_text_widget.bind("<FocusOut>", lambda e: callback(self.get_prompt_text()))

    def set_button_command(self, button_name: str, command: Callable):
        button_map = {
            "scratchpad": self.scratchpad_button,
            "ok_hide": self.ok_hide_button, # Added new button to map
            "start_stop": self.start_stop_button,
            "clear_queue": self.clear_queue_button,
            "pause_queue": self.pause_queue_button,
        }
        if button_name in button_map and button_map[button_name]:
            button_map[button_name].config(command=command)
        else:
             log_error(f"Cannot set command for unknown button: {button_name}")

    def add_menu_command(self, menu_type: str, label: Optional[str] = None, command: Optional[Callable] = None, **kwargs):
        menu_map = { "file": self.file_menu, "settings": self.settings_menu, "queue": self.queue_menu }
        target_menu = menu_map.get(menu_type.lower())
        
        if not target_menu:
            log_error(f"Cannot add command to unknown menu type: {menu_type}")
            return

        menu_item_type = kwargs.get("type")

        if menu_item_type == "separator":
            target_menu.add_separator()
        elif menu_item_type == "checkbutton":
            if label is None or command is None:
                log_error(f"Checkbutton menu item for '{menu_type}' menu is missing label or command.")
                return
            target_menu.add_checkbutton(label=label, command=command, variable=kwargs.get("variable"))
        else: 
            if label is None or command is None:
                log_error(f"Regular menu command for '{menu_type}' menu is missing label or command.")
                return
            target_menu.add_command(label=label, command=command, accelerator=kwargs.get("accelerator"))
