import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Callable, Optional
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
from settings_manager import AppSettings # For type hinting
from theme_manager import ThemeManager


class ScratchpadWindow(tk.Toplevel):
    def __init__(self, parent, settings: AppSettings, theme_manager: ThemeManager):
        super().__init__(parent)
        self.parent_app = parent # MainApp instance
        self.settings = settings # Live AppSettings reference
        self.theme_manager = theme_manager

        self.title("WhisperR Scratchpad")
        self.geometry("550x600") # Increased default size
        # self.transient(parent) # Not transient, can exist independently
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray) # Use new method

        self.text_widget: Optional[tk.Text] = None
        self.append_mode_var = tk.BooleanVar(value=self.settings.scratchpad_append_mode)

        self._apply_theme()
        self._create_widgets()
        
        # Load initial content if any (e.g., from a saved file if feature added later)

    def _apply_theme(self):
        # Assumes parent_app (MainApp) has settings_manager
        current_theme = self.settings.ui_theme
        colors = self.theme_manager.get_current_colors(self, current_theme) # Pass self for Toplevel styling
        self.configure(bg=colors["bg"])
        
        # Specific styling for this dialog's ttk widgets
        style = ttk.Style(self)
        style.configure('Scratchpad.TButton', background=colors["button_bg"], foreground=colors["button_fg"])
        style.map('Scratchpad.TButton',
            background=[('active', colors["select_bg"])],
            foreground=[('active', colors["select_fg"])])
        style.configure('Scratchpad.TCheckbutton', background=colors["bg"], foreground=colors["fg"])


    def _create_widgets(self):
        top_frame = ttk.Frame(self, padding=(10,10,10,0), style='TFrame')
        top_frame.pack(fill=tk.BOTH, expand=True)

        self.text_widget = tk.Text(top_frame, wrap=tk.WORD, undo=True)
        # Theming for tk.Text:
        current_theme = self.settings.ui_theme
        colors = self.theme_manager.get_current_colors(self, current_theme)
        self.text_widget.config(
            background=colors["text_bg"], foreground=colors["text_fg"],
            insertbackground=colors["fg"], # Cursor color
            selectbackground=colors["select_bg"], selectforeground=colors["select_fg"]
        )

        scrollbar_y = ttk.Scrollbar(top_frame, orient="vertical", command=self.text_widget.yview)
        self.text_widget.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bottom controls
        bottom_frame = ttk.Frame(self, padding=(10,5,10,10), style='TFrame')
        bottom_frame.pack(fill=tk.X)

        ttk.Button(bottom_frame, text="Import...", command=self._import_to_scratchpad, style='Scratchpad.TButton').pack(side=tk.LEFT, padx=(0,2))
        ttk.Button(bottom_frame, text="Export As...", command=self._export_from_scratchpad, style='Scratchpad.TButton').pack(side=tk.LEFT, padx=(2,2))
        ttk.Button(bottom_frame, text="Clear", command=self._clear_scratchpad, style='Scratchpad.TButton').pack(side=tk.LEFT, padx=(2,10))
        
        self.append_mode_check = ttk.Checkbutton(bottom_frame, text="Append Mode",
                                                 variable=self.append_mode_var,
                                                 command=self._toggle_append_mode, style='Scratchpad.TCheckbutton')
        self.append_mode_check.pack(side=tk.RIGHT)

    def show(self):
        self.deiconify()
        self.lift()
        self.focus_force()

    def hide_to_tray(self):
        """Hides the window. Called by WM_DELETE_WINDOW or minimize."""
        self.withdraw()

    def iconify(self):
        """Override minimize button to hide to tray."""
        self.hide_to_tray()

    def is_visible(self) -> bool:
        """Checks if the window is currently visible."""
        return self.winfo_exists() and self.winfo_viewable()

    def add_text(self, new_text: str):
        if not self.winfo_exists() or not self.text_widget:
            return
        try:
            if self.append_mode_var.get():
                # Get raw content to check if there's any actual text (non-whitespace)
                raw_current_content = self.text_widget.get("1.0", tk.END)
                has_existing_text = bool(raw_current_content.strip())

                separator = "" 

                if self.settings.clear_text_output: # "Clean Metadata" IS checked
                    if has_existing_text:
                        separator = " " # Add a single space if appending and cleaning metadata
                else: # "Clean Metadata" is NOT checked
                    if has_existing_text:
                        separator = "\n\n---\n\n"
                
                self.text_widget.insert(tk.END, separator + new_text)
                self.text_widget.see(tk.END) # Scroll to the end
            else:
                self.text_widget.delete("1.0", tk.END)
                self.text_widget.insert("1.0", new_text)
                self.text_widget.see("1.0") # Scroll to the top
        except Exception as e:
            log_error(f"Error updating scratchpad text: {e}")


    def _toggle_append_mode(self):
        self.settings.scratchpad_append_mode = self.append_mode_var.get()
        # MainApp should save settings if this is a persistent setting.
        # For now, it's just a session setting for the scratchpad.
        # If MainApp's settings save logic is triggered on config window close,
        # this won't be saved unless explicitly handled.
        log_extended(f"Scratchpad append mode: {self.settings.scratchpad_append_mode}")


    def _clear_scratchpad(self):
        if self.text_widget:
            if messagebox.askyesno("Clear Scratchpad", "Are you sure you want to clear all text from the scratchpad?", parent=self):
                self.text_widget.delete("1.0", tk.END)

    def _import_to_scratchpad(self):
        if not self.text_widget: return
        filepath = filedialog.askopenfilename(
            title="Import to Scratchpad",
            filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")),
            parent=self
        )
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.add_text(content) # Uses append mode logic
                log_essential(f"Imported content from {filepath} to scratchpad.")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import file: {e}", parent=self)
                log_error(f"Scratchpad import error: {e}")

    def _export_from_scratchpad(self):
        if not self.text_widget: return
        content = self.text_widget.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("Export Empty", "Scratchpad is empty, nothing to export.", parent=self)
            return

        filepath = filedialog.asksaveasfilename(
            title="Export Scratchpad As",
            filetypes=(("Text files", "*.txt"), ("Markdown files", "*.md"), ("All files", "*.*")),
            defaultextension=".txt",
            parent=self
        )
        if filepath:
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content)
                log_essential(f"Exported scratchpad content to {filepath}.")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export file: {e}", parent=self)
                log_error(f"Scratchpad export error: {e}")
