import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable, Optional
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
# Assuming AudioService is accessible or its methods are passed via callbacks
# from audio_service import AudioService

class VADCalibrationDialog(tk.Toplevel):
    def __init__(self, parent,
                 start_calibration_callback: Callable[[int, Callable, Callable], None],
                 cancel_calibration_callback: Callable[[], None],
                 initial_threshold: int,
                 theme_manager_ref):
        super().__init__(parent)
        self.transient(parent)
        self.title("VAD Energy Threshold Calibration")
        self.geometry("450x350") # Adjusted size
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.parent_app = parent # To access theme manager if needed via parent
        self.theme_manager = theme_manager_ref

        self.start_calibration_cb = start_calibration_callback
        self.cancel_calibration_cb = cancel_calibration_callback
        self.initial_threshold = initial_threshold
        self.recommended_threshold: Optional[int] = None
        self.is_calibrating = False

        self._apply_theme()
        self._create_widgets()

        # self.lift() # Already handled by transient + grab_set
        # self.focus_force() # wait_window takes care of focus
        self.wait_window() # Block until dialog is closed

    def _apply_theme(self):
        if hasattr(self.parent_app, 'settings_manager'): # MainApp should have this
            current_theme = self.parent_app.settings_manager.settings.ui_theme
            colors = self.theme_manager.get_current_colors(self, current_theme)
            self.configure(bg=colors["bg"])
            # Also update style for ttk widgets used in this dialog
            self.style = ttk.Style(self) # Ensure style is configured for this Toplevel
            # self.theme_manager.apply_theme(self, current_theme) # This might be too broad
            # Instead, configure specific styles if needed or rely on parent's styling.
            self.style.configure('VADDialog.TLabel', background=colors["bg"], foreground=colors["fg"])
            self.style.configure('VADDialog.TButton', background=colors["button_bg"], foreground=colors["button_fg"])
            self.style.map('VADDialog.TButton',
                background=[('active', colors["select_bg"]), ('disabled', colors["bg"])],
                foreground=[('active', colors["select_fg"]), ('disabled', colors["disabled_fg"])])

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding=10, style='TFrame')
        main_frame.pack(expand=True, fill=tk.BOTH)

        instructions_text = (
            "This tool will help you find a suitable VAD energy threshold.\n\n"
            "1. Click 'Start Calibration'.\n"
            "2. Remain SILENT for about 5 seconds to measure background noise.\n"
            "3. Then, speak normally for a few seconds.\n"
            "4. The tool will suggest a threshold based on the measurements."
        )
        ttk.Label(main_frame, text=instructions_text, justify=tk.LEFT, wraplength=400, style='VADDialog.TLabel').pack(pady=(0,15))

        self.status_var = tk.StringVar(value="Status: Idle")
        ttk.Label(main_frame, textvariable=self.status_var, style='VADDialog.TLabel').pack(pady=5)

        self.avg_energy_var = tk.StringVar(value="Avg Energy: N/A")
        ttk.Label(main_frame, textvariable=self.avg_energy_var, style='VADDialog.TLabel').pack(pady=2)
        self.peak_energy_var = tk.StringVar(value="Peak Energy: N/A")
        ttk.Label(main_frame, textvariable=self.peak_energy_var, style='VADDialog.TLabel').pack(pady=2)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(pady=10)
        
        self.calibration_duration_seconds = 5 # Default, can be made configurable

        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.pack(pady=10)

        self.start_button = ttk.Button(button_frame, text="Start Calibration", command=self._start_calibration_process, style='VADDialog.TButton')
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="Cancel Calibration", command=self._cancel_calibration_process, state=tk.DISABLED, style='VADDialog.TButton')
        self.cancel_button.pack(side=tk.LEFT, padx=5)

        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        self.result_frame = ttk.Frame(main_frame, style='TFrame')
        self.result_frame.pack(pady=5)
        
        self.recommended_var = tk.StringVar(value=f"Recommended: N/A (Current: {self.initial_threshold})")
        ttk.Label(self.result_frame, textvariable=self.recommended_var, style='VADDialog.TLabel').pack(side=tk.LEFT, padx=5)
        
        self.apply_button = ttk.Button(self.result_frame, text="Apply", command=self._on_apply, state=tk.DISABLED, style='VADDialog.TButton')
        self.apply_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(main_frame, text="Close", command=self._on_close, style='VADDialog.TButton').pack(pady=10)


    def _start_calibration_process(self):
        self.is_calibrating = True
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.apply_button.config(state=tk.DISABLED)
        self.status_var.set("Status: Calibrating... (Be silent, then speak)")
        self.progress_var.set(0)
        self.avg_energy_var.set("Avg Energy: Measuring...")
        self.peak_energy_var.set("Peak Energy: Measuring...")
        self.recommended_var.set(f"Recommended: N/A (Current: {self.initial_threshold})")

        # Call the actual start method in AudioService
        self.start_calibration_cb(
            self.calibration_duration_seconds,
            self._update_calibration_ui,  # Callback for live updates
            self._calibration_finished_ui # Callback for when done
        )
        self._update_progress()

    def _update_progress(self):
        if not self.is_calibrating:
            return
        
        current_progress = self.progress_var.get()
        max_progress = self.calibration_duration_seconds * 10 # (0.1s steps)
        
        if current_progress < max_progress:
            self.progress_var.set(current_progress + 1)
            self.after(100, self._update_progress) # Update every 100ms
        else: # Should be stopped by _calibration_finished_ui ideally
            if self.is_calibrating: # If somehow still calibrating
                 self.status_var.set("Status: Finishing up...")


    def _cancel_calibration_process(self):
        if self.is_calibrating:
            self.cancel_calibration_cb() # This should trigger AudioService to stop its loop
            # The _calibration_finished_ui will be called by AudioService eventually
        self._reset_ui_after_calibration("Cancelled")


    def _update_calibration_ui(self, avg_energy: float, peak_energy: float, is_done: bool, status_message: str = ""):
        """Callback from AudioService with live VAD data."""
        if not self.is_calibrating and not is_done:
            return

        self.avg_energy_var.set(f"Avg Energy: {avg_energy:.2f}")
        self.peak_energy_var.set(f"Peak Energy: {peak_energy:.2f}")
        if status_message:
            self.status_var.set(f"Status: {status_message}")
        if is_done:
            self.status_var.set("Status: Processing results...")


    def _calibration_finished_ui(self, recommended_threshold: int):
        """Callback from AudioService when calibration is fully processed."""
        self.is_calibrating = False # Ensure flag is reset
        self.recommended_threshold = recommended_threshold
        self.recommended_var.set(f"Recommended: {self.recommended_threshold} (Current: {self.initial_threshold})")
        self.apply_button.config(state=tk.NORMAL)
        self._reset_ui_after_calibration("Completed")


    def _reset_ui_after_calibration(self, status_message_suffix: str):
        self.is_calibrating = False
        self.start_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.status_var.set(f"Status: {status_message_suffix}")
        self.progress_var.set(self.calibration_duration_seconds * 10) # Fill progress bar


    def _on_apply(self):
        if self.recommended_threshold is not None:
            # The dialog itself doesn't change the setting; it returns the value.
            # The caller (ConfigWindow) will handle applying it.
            self.destroy() # Close dialog, recommended_threshold is available to caller
        else:
            messagebox.showinfo("No Recommendation", "Calibration did not yield a recommendation.", parent=self)

    def _on_close(self):
        if self.is_calibrating:
            if messagebox.askyesno("Cancel Calibration?", "Calibration is in progress. Do you want to cancel it and close?", parent=self):
                self._cancel_calibration_process() # This will eventually lead to closing
            else:
                return # Don't close
        self.recommended_threshold = None # Ensure no value is returned if closed without applying
        self.destroy()

    @staticmethod
    def show(parent, start_cb, cancel_cb, initial_thresh, theme_manager) -> Optional[int]:
        dialog = VADCalibrationDialog(parent, start_cb, cancel_cb, initial_thresh, theme_manager)
        return dialog.recommended_threshold # Returns None if closed/cancelled, or the value if applied
