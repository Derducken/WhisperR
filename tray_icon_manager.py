import tkinter as tk
import pystray # type: ignore
from PIL import Image, ImageTk # type: ignore
import os
import sys
from pathlib import Path
from typing import Callable, Optional
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
from constants import APP_ICON_NAME

class TrayIconManager:
    def __init__(self,
                 app_name: str,
                 root_window: tk.Tk,
                 show_window_action: Callable,
                 toggle_recording_action: Callable,
                 quit_action: Callable,
                 base_path: Path):
        self.app_name = app_name
        self.root = root_window
        self.show_window_action = show_window_action
        self.toggle_recording_action = toggle_recording_action
        self.quit_action = quit_action
        self.base_path = base_path # For finding icon.png

        self.tray_icon: Optional[pystray.Icon] = None
        self.icon_image: Optional[Image.Image] = None
        self.default_icon_path: Optional[Path] = None


    def _load_icon(self) -> Optional[Image.Image]:
        # Try to find icon.png in various places
        # 1. Next to the script/executable (via self.base_path)
        # 2. In sys._MEIPASS for PyInstaller
        
        potential_paths = []
        if hasattr(sys, '_MEIPASS'):
            potential_paths.append(Path(sys._MEIPASS) / APP_ICON_NAME)
        
        potential_paths.append(self.base_path / APP_ICON_NAME) # From script/exe dir
        potential_paths.append(Path(os.path.dirname(os.path.abspath(__file__))) / APP_ICON_NAME) # Relative to this file
        potential_paths.append(Path(".") / APP_ICON_NAME) # Current working directory

        for icon_path_obj in potential_paths:
            icon_path_norm = icon_path_obj.resolve()
            if icon_path_norm.exists() and icon_path_norm.is_file():
                try:
                    self.default_icon_path = icon_path_norm
                    log_extended(f"Tray icon found at: {self.default_icon_path}")
                    return Image.open(self.default_icon_path)
                except Exception as e:
                    log_error(f"Error loading tray icon image from {icon_path_norm}: {e}")
        
        log_error(f"{APP_ICON_NAME} not found in expected locations. Using fallback gray icon.")
        return Image.new('RGB', (64, 64), color='gray') # Fallback

    def _on_show_window(self, icon, item):
        self.root.after(0, self.show_window_action)

    def _on_toggle_recording(self, icon, item):
        self.root.after(0, self.toggle_recording_action)

    def _on_quit(self, icon, item):
        self.root.after(0, self.quit_action) # quit_action should handle its own threading if needed

    def setup_tray_icon(self):
        """This method should be called in a separate thread."""
        if self.tray_icon and self.tray_icon.visible:
            log_extended("Tray icon already running.")
            return

        self.icon_image = self._load_icon()
        if not self.icon_image:
            log_error("Failed to load or create any icon image for tray.")
            return

        menu = pystray.Menu(
            pystray.MenuItem(self.app_name, None, enabled=False), # Title, non-clickable
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Show Window", self._on_show_window, default=True),
            pystray.MenuItem("Toggle Recording", self._on_toggle_recording),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit " + self.app_name, self._on_quit)
        )

        self.tray_icon = pystray.Icon(self.app_name.lower().replace(" ", "_"), self.icon_image, self.app_name, menu)
        
        try:
            log_essential("Starting tray icon...")
            self.tray_icon.run() # This is a blocking call
        except SystemExit:
             log_essential("Tray icon exited via SystemExit (likely on app quit).")
        except Exception as e:
            log_error(f"Error running tray icon: {e}", exc_info=True)
            self.tray_icon = None # Ensure it's None if run failed
        finally:
            log_extended("Tray icon run() method has finished.")


    def stop_tray_icon(self):
        if self.tray_icon and hasattr(self.tray_icon, 'stop') and self.tray_icon.visible:
            try:
                log_extended("Stopping tray icon...")
                self.tray_icon.stop()
            except Exception as e:
                log_error(f"Error stopping tray icon: {e}")
        self.tray_icon = None # Clear reference

    def notify(self, title: str, message: str):
        if self.tray_icon and self.tray_icon.HAS_NOTIFICATION:
            try:
                self.tray_icon.notify(message, title)
                log_extended(f"Tray notification: {title} - {message}")
            except Exception as e:
                log_error(f"Failed to send tray notification: {e}")
        else:
            log_extended("Tray icon not available or does not support notifications for this system.")
            # Fallback to messagebox if root window is available and not withdrawn (or handle differently)
            # if self.root and self.root.winfo_exists() and self.root.state() == 'normal':
            #    messagebox.showinfo(title, message, parent=self.root)