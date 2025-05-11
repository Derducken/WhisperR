import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk # type: ignore
from pathlib import Path
from typing import Optional
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
from constants import (
    ALT_INDICATOR_POSITIONS, DEFAULT_ALT_INDICATOR_POSITION,
    DEFAULT_ALT_INDICATOR_SIZE, DEFAULT_ALT_INDICATOR_OFFSET,
    MIN_ALT_INDICATOR_SIZE, MAX_ALT_INDICATOR_SIZE,
    MIN_ALT_INDICATOR_OFFSET, MAX_ALT_INDICATOR_OFFSET,
    COLOR_STATUS_IDLE_NOT_RECORDING # Default color for bg if icon is transparent
)

class AltStatusIndicator:
    def __init__(self, root_tk_instance: tk.Tk, base_path: Path, theme_manager_ref):
        self.root = root_tk_instance
        self.base_path = base_path / "status_icons" # Path to status_icons folder
        self.theme_manager = theme_manager_ref
        self.indicator_window: Optional[tk.Toplevel] = None
        self.icon_label: Optional[ttk.Label] = None
        self.current_icon_path: Optional[Path] = None
        self.icon_cache: dict[str, ImageTk.PhotoImage] = {} # Cache PhotoImage objects

        self.enabled = False
        self.position = DEFAULT_ALT_INDICATOR_POSITION
        self.size = DEFAULT_ALT_INDICATOR_SIZE
        self.offset = DEFAULT_ALT_INDICATOR_OFFSET
        self.current_bg_color = COLOR_STATUS_IDLE_NOT_RECORDING # Fallback if icon has transparency

        self.icon_map = {
            "idle": self.base_path / "idle.png",
            "recording_on": self.base_path / "recording_on.png", # VAD speaking or continuous
            "recording_vad_wait": self.base_path / "recording_off.png", # VAD waiting
            "transcribing": self.base_path / "transcribing.png",
            "rec_and_transcribing": self.base_path / "rec_and_transcribing.png" # Optional: specific icon
        }
        # Ensure base_path for icons exists
        self.base_path.mkdir(parents=True, exist_ok=True)
        # Check for icons and log missing ones (users need to provide these)
        for name, path in self.icon_map.items():
            if not path.exists():
                log_extended(f"Alternative status icon missing: {path}. Please create it.")


    def configure(self, enabled: bool, position: str, size: int, offset: int):
        self.enabled = enabled
        self.position = position if position in ALT_INDICATOR_POSITIONS else DEFAULT_ALT_INDICATOR_POSITION
        self.size = max(MIN_ALT_INDICATOR_SIZE, min(size, MAX_ALT_INDICATOR_SIZE))
        self.offset = max(MIN_ALT_INDICATOR_OFFSET, min(offset, MAX_ALT_INDICATOR_OFFSET))

        if self.enabled:
            self.create_or_update_indicator()
            # Set an initial icon, e.g., idle
            self.update_icon_by_state("idle")
        else:
            self.destroy_indicator()

    def destroy_indicator(self):
        if self.indicator_window and self.indicator_window.winfo_exists():
            try:
                self.indicator_window.destroy()
            except tk.TclError:
                pass
        self.indicator_window = None
        self.icon_label = None
        self.icon_cache.clear() # Clear image cache
        log_extended("Alternative status indicator destroyed.")

    def _get_icon_image(self, icon_name: str) -> Optional[ImageTk.PhotoImage]:
        icon_path = self.icon_map.get(icon_name)
        if not icon_path or not icon_path.exists():
            log_extended(f"Icon '{icon_name}' not found at {icon_path}")
            return None

        if str(icon_path) in self.icon_cache:
            return self.icon_cache[str(icon_path)]

        try:
            img = Image.open(icon_path)
            img = img.resize((self.size, self.size), Image.Resampling.LANCZOS)
            photo_img = ImageTk.PhotoImage(img)
            self.icon_cache[str(icon_path)] = photo_img # Cache it
            return photo_img
        except Exception as e:
            log_error(f"Error loading or resizing icon {icon_path}: {e}")
            return None

    def create_or_update_indicator(self):
        if not self.enabled:
            self.destroy_indicator()
            return

        self.destroy_indicator() # Destroy existing first
        log_debug("Creating alternative status indicator...")

        self.indicator_window = tk.Toplevel(self.root)
        self.indicator_window.overrideredirect(True)
        self.indicator_window.attributes("-topmost", True)
        # Make window background transparent (works on some OSes, not all for Toplevel)
        # For Windows, -transparentcolor might work if set to a specific color.
        # For macOS/Linux, proper compositing manager needed.
        # A common trick is to set a specific color and make it transparent.
        # For simplicity, we'll set the label's background, and if the icon has transparency, it'll show.
        # self.indicator_window.attributes("-transparentcolor", "white") # Example

        # Use themed background for the label, in case icon has alpha
        # This ensures the label itself blends if the icon isn't fully opaque
        colors = self.theme_manager.get_current_colors(self.root, self.theme_manager.current_theme_name)
        self.current_bg_color = colors.get("bg", COLOR_STATUS_IDLE_NOT_RECORDING) # Fallback
        
        # For Windows, try to use -transparentcolor for true alpha blending of PNGs
        # Use pure white as the key color, assuming icons are designed with this in mind for transparency.
        transparent_key_color = "#FFFFFF" # Pure white

        self.indicator_window.config(bg=transparent_key_color) # Set Toplevel BG to the key color
        try:
            # This is the key for Windows transparency with overrideredirect
            self.indicator_window.attributes("-transparentcolor", transparent_key_color)
            log_debug(f"Set -transparentcolor to {transparent_key_color} for alt status indicator.")
        except tk.TclError as e_trans:
            log_extended(f"Setting -transparentcolor failed for alt status indicator: {e_trans}")
            # Fallback for other OS or if -transparentcolor fails:
            try:
                self.indicator_window.wm_attributes("-transparent", True)
                # On some systems, setting a system-defined transparent bg helps
                self.indicator_window.config(bg='systemTransparent') 
                log_debug("Set wm_attributes -transparent to True and bg to systemTransparent.")
            except tk.TclError as e_wm:
                log_extended(f"Setting wm_attributes -transparent failed: {e_wm}")
                # If all else fails, use the theme background
                self.indicator_window.config(bg=self.current_bg_color)


        self.icon_label = ttk.Label(self.indicator_window, background=transparent_key_color)
        self.icon_label.pack(fill=tk.BOTH, expand=True)

        self._set_geometry()
        log_debug(f"Alternative status indicator created: Pos={self.position}, Size={self.size}, Offset={self.offset}")

    def _set_geometry(self):
        if not self.indicator_window or not self.indicator_window.winfo_exists():
            return

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        win_width = self.size
        win_height = self.size

        x, y = 0, 0
        if self.position == "Top-Left":
            x = self.offset
            y = self.offset
        elif self.position == "Top-Right":
            x = screen_width - win_width - self.offset
            y = self.offset
        elif self.position == "Bottom-Left":
            x = self.offset
            y = screen_height - win_height - self.offset
        elif self.position == "Bottom-Right": # Default
            x = screen_width - win_width - self.offset
            y = screen_height - win_height - self.offset
        
        self.indicator_window.geometry(f"{win_width}x{win_height}+{x}+{y}")


    def update_icon_by_state(self, state_key: str):
        if not self.enabled or not self.icon_label or not self.indicator_window or not self.indicator_window.winfo_exists():
            return

        photo_img = self._get_icon_image(state_key)
        if photo_img:
            self.icon_label.configure(image=photo_img)
            # Keep a reference to prevent garbage collection if not cached elsewhere correctly
            self.icon_label.image = photo_img 
        else:
            # Fallback: show text or clear image if icon load fails
            self.icon_label.configure(image=None, text=state_key[:3].upper()) 
            log_warning(f"Failed to load icon for state: {state_key}. Displaying text.")

        # If window transparency is tricky, update label BG to match current color status
        # For example, if icon is red for recording, label bg could also be red.
        # This is a fallback if true window transparency isn't working.
        # self.icon_label.config(background=self.current_status_color_hex_from_app)

    def update_theme(self):
        """Called when the main application theme changes."""
        if self.enabled and self.indicator_window and self.indicator_window.winfo_exists():
            current_theme_name = self.theme_manager.current_theme_name
            colors = self.theme_manager.get_current_colors(self.root, current_theme_name)
            self.current_bg_color = colors.get("bg", COLOR_STATUS_IDLE_NOT_RECORDING)
            
            self.indicator_window.config(bg=self.current_bg_color)
            if self.icon_label:
                self.icon_label.config(background=self.current_bg_color)
            log_extended("Alternative status indicator theme updated.")
