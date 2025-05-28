import tkinter as tk
from tkinter import ttk, Menu
from typing import Callable, List, Tuple, Optional 
from app_logger import get_logger, log_extended, log_error, log_debug
from constants import DEFAULT_LANGUAGE, DEFAULT_MODEL, EXTENDED_MODEL_OPTIONS, Theme
from settings_manager import AppSettings, SettingsManager
from theme_manager import ThemeManager

class MainWindowView:
    def __init__(self, root: tk.Tk, settings_manager: SettingsManager, initial_prompt: str, theme_manager: ThemeManager):
        self.root = root
        self.settings_manager = settings_manager
        self.initial_prompt_val = initial_prompt
        self.theme_manager = theme_manager

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
        self.ok_hide_button: Optional[ttk.Button] = None
        self.queue_status_label: Optional[ttk.Label] = None
        self.clear_queue_button: Optional[ttk.Button] = None

        self.model_priming_status_var = tk.StringVar(value="")
        self.model_priming_status_label: Optional[ttk.Label] = None
        self._model_priming_status_job_id: Optional[str] = None
        
        # --- NEW for precise model change detection ---
        self._previous_model_value: Optional[str] = self.model_var.get() # Initialize with current value
        # --- END NEW ---


        self._create_widgets()

        if self.prompt_text_widget:
            self.prompt_text_widget.delete("1.0", tk.END) 
            self.prompt_text_widget.insert("1.0", self.initial_prompt_val)
        
        self.update_ui_from_settings()

    def _create_widgets(self):
        # ... (This method is UNCHANGED from the version I last sent you with the priming status label) ...
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        self.file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=self.file_menu)
        self.settings_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=self.settings_menu)
        self.queue_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Queue", menu=self.queue_menu)

        top_controls_frame = ttk.Frame(self.root, style='TFrame', padding=(10,10,10,0))
        top_controls_frame.pack(fill=tk.X)

        lang_frame = ttk.Frame(top_controls_frame, style='TFrame')
        lang_frame.pack(side=tk.LEFT, padx=(0, 10), anchor=tk.NW) 
        ttk.Label(lang_frame, text="Language:", style='TLabel').pack(side=tk.TOP, anchor=tk.W)
        self.language_options = ["auto", "en", "es", "fr", "de", "it", "ja", "zh", "ko", "ru", "pt", "el"] 
        self.language_combobox = ttk.Combobox(lang_frame, textvariable=self.language_var,
                                              values=self.language_options, state="readonly", width=10)
        self.language_combobox.pack(side=tk.TOP, anchor=tk.W)

        model_outer_frame = ttk.Frame(top_controls_frame, style='TFrame') 
        model_outer_frame.pack(side=tk.RIGHT, padx=(0,0), anchor=tk.NE, fill=tk.X, expand=True)

        model_select_frame = ttk.Frame(model_outer_frame, style='TFrame')
        model_select_frame.pack(side=tk.TOP, anchor=tk.E) 
        ttk.Label(model_select_frame, text="Model (CLI/Lib Fallback):", style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.model_combobox = ttk.Combobox(model_select_frame, textvariable=self.model_var,
                                           values=EXTENDED_MODEL_OPTIONS, width=28) 
        self.model_combobox.pack(side=tk.LEFT)
        
        self.model_priming_status_label = ttk.Label(
            model_outer_frame, 
            textvariable=self.model_priming_status_var,
            style="TLabel", # Using default TLabel style for now
            anchor=tk.E, 
            justify=tk.RIGHT,
            wraplength=300 
        )
        self.model_priming_status_label.pack(side=tk.TOP, anchor=tk.E, pady=(2,0), fill=tk.X)


        toggle_frame1 = ttk.Frame(self.root, style='TFrame', padding=(10,5))
        toggle_frame1.pack(fill=tk.X)
        ttk.Checkbutton(toggle_frame1, text="Enable Translation", variable=self.translation_var, style='TCheckbutton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(toggle_frame1, text="Auto-Pause / Commands (VAD)", variable=self.command_mode_var, style='TCheckbutton').pack(side=tk.RIGHT, padx=(0,0))

        toggle_frame2 = ttk.Frame(self.root, style='TFrame', padding=(10,2,10,5))
        toggle_frame2.pack(fill=tk.X)
        ttk.Checkbutton(toggle_frame2, text="Hide Timestamps (Output)", variable=self.timestamps_disabled_var, style='TCheckbutton').pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(toggle_frame2, text="Clean Metadata (Output)", variable=self.clear_text_output_var, style='TCheckbutton').pack(side=tk.RIGHT, padx=(0,0))

        prompt_label_frame = ttk.Frame(self.root, style='TFrame', padding=(10,10,10,0))
        prompt_label_frame.pack(fill=tk.X)
        ttk.Label(prompt_label_frame, text="Whisper Initial Prompt:", style='TLabel').pack(anchor=tk.W)
        
        self.prompt_text_widget = tk.Text(self.root, height=8, wrap=tk.WORD, undo=True)
        self.prompt_text_widget.pack(pady=5, padx=10, fill=tk.X, expand=False)

        self.scratchpad_button = ttk.Button(self.root, text="Open Scratchpad")
        self.scratchpad_button.pack(pady=(5, 10), padx=10, fill=tk.X, ipady=8)

        self.ok_hide_button = ttk.Button(self.root, text="OK (Hide Window)")
        self.ok_hide_button.pack(pady=(0, 10), padx=10, fill=tk.X, ipady=8)

        self.start_stop_frame = ttk.Frame(self.root, style='TFrame', padding=(10,0,10,0))
        self.start_stop_frame.pack(fill=tk.X)
        self.start_stop_button = ttk.Button(self.start_stop_frame, text="Start Recording")
        self.start_stop_button.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=10)
        self.recording_indicator_label = ttk.Label(self.start_stop_frame, text="●", font=("Arial", 16), style='TLabel')
        self.recording_indicator_label.pack(side=tk.LEFT, padx=10)
        self.update_recording_indicator_ui()

        self.bottom_frame = ttk.Frame(self.root, style='TFrame', padding=(10,0,10,10)) 
        self.bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, anchor=tk.S)
        
        self.hotkey_display_label = ttk.Label(self.bottom_frame, textvariable=self.shortcut_display_var,
                                             justify=tk.LEFT, style='TLabel', wraplength=250) 
        self.hotkey_display_label.pack(side=tk.LEFT, anchor=tk.W, expand=True, fill=tk.X)
        
        queue_controls_subframe = ttk.Frame(self.bottom_frame, style='TFrame')
        queue_controls_subframe.pack(side=tk.RIGHT, anchor=tk.E)
        
        self.queue_status_label = ttk.Label(queue_controls_subframe, textvariable=self.queue_indicator_var,
                                               justify=tk.RIGHT, style='TLabel')
        self.queue_status_label.pack(side=tk.RIGHT, anchor=tk.E, padx=(10, 0))
        
        self.clear_queue_button = ttk.Button(queue_controls_subframe, text="Clear Q", width=8)
        self.clear_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))
        self.pause_queue_button = ttk.Button(queue_controls_subframe, text="Pause Q", width=9)
        self.pause_queue_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(5, 0))

    def set_transient_status_message(self, message: str, duration_ms: int = 5000): # UNCHANGED
        if not self.model_priming_status_label or not self.model_priming_status_label.winfo_exists():
            log_debug(f"Transient status label not available for message: {message}")
            return

        self.model_priming_status_var.set(message)
        
        if self._model_priming_status_job_id:
            self.root.after_cancel(self._model_priming_status_job_id)
            self._model_priming_status_job_id = None
        
        if duration_ms > 0:
            self._model_priming_status_job_id = self.root.after(
                duration_ms, 
                lambda: self.model_priming_status_var.set("") if self.model_priming_status_var else None
            )

    def update_ui_from_settings(self): # UNCHANGED
        current_settings = self.settings_manager.settings
        self.language_var.set(current_settings.language)
        self.model_var.set(current_settings.model)
        self._previous_model_value = current_settings.model # Ensure previous value is synced
        self.translation_var.set(current_settings.translation_enabled)
        self.command_mode_var.set(current_settings.command_mode)
        self.timestamps_disabled_var.set(current_settings.timestamps_disabled)
        self.clear_text_output_var.set(current_settings.clear_text_output)
        
        self.update_shortcut_display_ui()
        log_extended("Main window UI updated from settings.")

    def update_recording_indicator_ui(self, is_recording: Optional[bool] = None, vad_is_speaking: Optional[bool] = None): # UNCHANGED
        if is_recording is not None:
            self.is_recording_visual_indicator = is_recording

        if self.start_stop_button and self.start_stop_button.winfo_exists():
            self.start_stop_button.config(text="Stop Recording" if self.is_recording_visual_indicator else "Start Recording")
        
        if self.recording_indicator_label and self.recording_indicator_label.winfo_exists():
            indicator_color = "#808080" # Default gray for non-recording
            current_settings = self.settings_manager.settings
            colors = self.theme_manager.get_current_colors(self.root, current_settings.ui_theme)
            indicator_color = colors.get("disabled_fg", indicator_color) 

            if self.is_recording_visual_indicator:
                if current_settings.command_mode: 
                    indicator_color = colors.get("vad_active_fg", "orange") if vad_is_speaking else colors.get("vad_waiting_fg", "darkkhaki")
                else: 
                    indicator_color = colors.get("recording_fg", "red")
            self.recording_indicator_label.config(foreground=indicator_color)

    def update_shortcut_display_ui(self): # UNCHANGED
        current_settings = self.settings_manager.settings
        ptt_key = current_settings.hotkey_push_to_talk or "[Not Set]"
        toggle_key = current_settings.hotkey_toggle_record or "[Not Set]"
        show_key = current_settings.hotkey_show_window or "[Not Set]"
        self.shortcut_display_var.set(f"PTT: {ptt_key}\nToggle: {toggle_key}\nShow: {show_key}")

    def update_queue_indicator_ui(self, queue_size: int, is_paused: bool): # UNCHANGED
        status_parts = []
        if queue_size > 0:
            status_parts.append(f"{queue_size}")
        else:
            status_parts.append("Empty")

        if is_paused and queue_size > 0:
            status_parts.append("(Paused)")
        
        final_text = f"Queue: {' '.join(status_parts)}"
        self.queue_indicator_var.set(final_text)
        
        self.update_pause_queue_button_ui(is_paused) 
        log_debug(f"UI Queue Indicator Updated: Size={queue_size}, Paused={is_paused}, Text='{final_text}'")


    def update_pause_queue_button_ui(self, is_paused: bool): # UNCHANGED
        if self.pause_queue_button and self.pause_queue_button.winfo_exists():
            self.pause_queue_button.config(text="Resume Q" if is_paused else "Pause Q")
        self.pause_queue_menu_var.set(is_paused)

    def get_prompt_text(self) -> str: # UNCHANGED
        if self.prompt_text_widget and self.prompt_text_widget.winfo_exists():
            return self.prompt_text_widget.get("1.0", tk.END).strip()
        return "" 

    def set_prompt_widget_text(self, text: str): # UNCHANGED
        if self.prompt_text_widget and self.prompt_text_widget.winfo_exists():
            self.prompt_text_widget.delete("1.0", tk.END)
            self.prompt_text_widget.insert("1.0", text)

    def bind_language_change(self, callback: Callable[[str], None]): # UNCHANGED
        if self.language_combobox:
            self.language_combobox.bind("<<ComboboxSelected>>", lambda e: callback(self.language_var.get()))

    # --- MODIFIED ---
    def bind_model_change(self, callback: Callable[[str], None]):
        if self.model_combobox:
            # Ensure _previous_model_value is initialized if not already
            if not hasattr(self, '_previous_model_value') or self._previous_model_value is None:
                 self._previous_model_value = self.model_var.get()

            def on_model_potentially_changed(event_type: str, event_widget=None):
                # For ComboboxSelected, the value in model_var is already updated.
                # For Return/FocusOut on a non-readonly Combobox, model_var.get() is also current.
                current_value = self.model_var.get()
                
                # If the combobox is readonly, any event means a selection change.
                # If it's not readonly, user might type.
                # <<ComboboxSelected>> is the most reliable for actual selection change.
                # For Return/FocusOut, we must check if the value actually changed.

                if event_type == "<<ComboboxSelected>>":
                    if current_value != self._previous_model_value:
                        log_debug(f"Model selected via dropdown: '{current_value}'. Previous: '{self._previous_model_value}'. Calling callback.")
                        self._previous_model_value = current_value
                        callback(current_value)
                    # else:
                        # log_debug(f"<<ComboboxSelected>> but value '{current_value}' is same as previous. No callback.")
                
                elif event_type in ["<Return>", "<FocusOut>"]:
                    # This logic is for when users can TYPE into the combobox.
                    # If combobox were 'readonly', these might not be as necessary or behave differently.
                    if current_value != self._previous_model_value:
                        log_debug(f"Model potentially changed by {event_type} to '{current_value}'. Previous: '{self._previous_model_value}'. Calling callback.")
                        self._previous_model_value = current_value
                        callback(current_value)
                    # else:
                        # log_debug(f"Model event {event_type} but value '{current_value}' is same as previous. No callback.")
                
            self.model_combobox.bind("<<ComboboxSelected>>", lambda e: on_model_potentially_changed("<<ComboboxSelected>>", e.widget))
            self.model_combobox.bind("<Return>", lambda e: on_model_potentially_changed("<Return>", e.widget))
            # FocusOut can be problematic if the user is just clicking around or tabbing without intending to change.
            # Only bind FocusOut if really necessary and after thorough testing.
            # For now, let's remove it to avoid over-eager priming. If the user types and hits Enter, <Return> covers it.
            # If they type and click something else, they might expect the change to take.
            # Let's keep FocusOut but the check against _previous_model_value is critical.
            self.model_combobox.bind("<FocusOut>", lambda e: on_model_potentially_changed("<FocusOut>", e.widget))
            log_debug("Model change events bound to <<ComboboxSelected>>, <Return>, and <FocusOut> with change detection.")
    # --- END MODIFIED ---

    def bind_toggle_change(self, var_name: str, callback: Callable[[bool], None]): # UNCHANGED
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

    def bind_prompt_change(self, callback: Callable[[str], None]): # UNCHANGED
        if self.prompt_text_widget:
            self._prompt_update_job: Optional[str] = None 
            def on_key_release(event): 
                if self._prompt_update_job:
                    self.root.after_cancel(self._prompt_update_job)
                self._prompt_update_job = self.root.after(750, lambda: callback(self.get_prompt_text()))
            
            self.prompt_text_widget.bind("<KeyRelease>", on_key_release)
            self.prompt_text_widget.bind("<FocusOut>", lambda e: callback(self.get_prompt_text()))

    def set_button_command(self, button_name: str, command: Callable): # UNCHANGED
        button_map = {
            "scratchpad": self.scratchpad_button,
            "ok_hide": self.ok_hide_button,
            "start_stop": self.start_stop_button,
            "clear_queue": self.clear_queue_button,
            "pause_queue": self.pause_queue_button,
        }
        if button_name in button_map and button_map[button_name]: 
            button_map[button_name].config(command=command)
        else:
             log_error(f"Cannot set command for unknown or uninitialized button: {button_name}")


    def add_menu_command(self, menu_type: str, label: Optional[str] = None, command: Optional[Callable] = None, **kwargs): # UNCHANGED
        menu_map = { "file": self.file_menu, "settings": self.settings_menu, "queue": self.queue_menu }
        target_menu = menu_map.get(menu_type.lower())
        
        if not target_menu:
            log_error(f"Cannot add command to unknown menu type: {menu_type}")
            return

        menu_item_type = kwargs.get("type")

        if menu_item_type == "separator":
            target_menu.add_separator()
        elif menu_item_type == "checkbutton":
            if label is None: 
                log_error(f"Checkbutton menu item for '{menu_type}' menu is missing label.")
                return
            target_menu.add_checkbutton(label=label, command=command, variable=kwargs.get("variable"))
        else: 
            if label is None or command is None:
                log_error(f"Regular menu command for '{menu_type}' menu is missing label or command.")
                return
            target_menu.add_command(label=label, command=command, accelerator=kwargs.get("accelerator"))
