import tkinter as tk
from tkinter import ttk
from typing import Optional

class ConfigSection(ttk.LabelFrame):
    def __init__(self, parent, text: str, theme_manager_ref=None, **kwargs):
        # Access theme for styling the labelframe itself if possible
        # This assumes parent (e.g. a tab in config window) is already themed.
        # ttk.LabelFrame should pick up theme from its parent style.
        super().__init__(parent, text=text, style="Config.TLabelframe", **kwargs)
        
        # If direct theme colors are needed:
        # if theme_manager_ref and hasattr(parent, 'settings_manager'): # parent is usually ConfigWindow
        #     current_theme = parent.settings_manager.settings.ui_theme
        #     colors = theme_manager_ref.get_current_colors(parent, current_theme)
        #     self.configure(background=colors["bg"], foreground=colors["fg"])
        #     # For the label part of the LabelFrame:
        #     style = ttk.Style(parent)
        #     style.configure("Config.TLabelframe.Label", background=colors["bg"], foreground=colors["fg"])


        self.pack(fill=tk.X, padx=5, pady=(0 if not parent.winfo_children() else 10, 10))
        
        self.inner_frame = ttk.Frame(self, padding=(10, 5, 10, 10), style='TFrame')
        self.inner_frame.pack(fill=tk.X, expand=True)

    def get_inner_frame(self) -> ttk.Frame:
        return self.inner_frame

def create_browse_row(parent: ttk.Frame, label_text: str, entry_var: tk.StringVar, browse_command: callable, entry_width: int = 30):
    """Creates a Label, Entry, and Browse Button row."""
    row_frame = ttk.Frame(parent, style='TFrame')
    row_frame.pack(fill=tk.X, pady=(2,0))
    
    ttk.Label(row_frame, text=label_text, style='TLabel').pack(side=tk.LEFT, padx=(0,5), anchor=tk.W)
    
    entry = ttk.Entry(row_frame, textvariable=entry_var, width=entry_width)
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
    
    ttk.Button(row_frame, text="Browse...", command=browse_command, style='TButton').pack(side=tk.LEFT)
    return entry # Return entry for potential direct manipulation