import tkinter as tk
import traceback
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 
from constants import (
    DEFAULT_STATUS_BAR_POSITION, DEFAULT_STATUS_BAR_SIZE, STATUS_BAR_POSITIONS,
    COLOR_STATUS_IDLE_NOT_RECORDING
)

# Conditional import for Windows specific features
try:
    import win32gui
    import win32con
    import win32api
    WINDOWS_FEATURES_AVAILABLE = True
except ImportError:
    WINDOWS_FEATURES_AVAILABLE = False
    log_extended("Windows specific libraries (pywin32) not found. Screen edge status bar disabled.")


class StatusBarManager:
    def __init__(self, root_tk_instance: tk.Tk, theme_manager_ref): # Pass ThemeManager instance
        self.root = root_tk_instance
        self.theme_manager = theme_manager_ref # Store reference
        self.status_bar_window: Optional[tk.Toplevel] = None
        self.status_bar_frame: Optional[tk.Frame] = None

        self.enabled = False
        self.position = DEFAULT_STATUS_BAR_POSITION
        self.size = DEFAULT_STATUS_BAR_SIZE
        self.current_bar_color_hex = COLOR_STATUS_IDLE_NOT_RECORDING # Initial default

        if not WINDOWS_FEATURES_AVAILABLE:
            self.enabled = False # Force disable if win32 libs are missing

    def configure(self, enabled: bool, position: str, size: int):
        self.enabled = enabled if WINDOWS_FEATURES_AVAILABLE else False
        self.position = position if position in STATUS_BAR_POSITIONS else DEFAULT_STATUS_BAR_POSITION
        self.size = max(1, min(size, 100)) # Clamp size

        if self.enabled:
            self.create_or_update_status_bar()
        else:
            self.destroy_status_bar()

    def destroy_status_bar(self):
        if self.status_bar_window and self.status_bar_window.winfo_exists():
            try:
                self.status_bar_window.destroy()
            except tk.TclError: # Might already be destroyed
                pass
        self.status_bar_window = None
        self.status_bar_frame = None
        log_extended("Windows edge status bar destroyed.")

    def _get_primary_monitor_info(self):
        if not WINDOWS_FEATURES_AVAILABLE: return None
        try:
            monitors = win32api.EnumDisplayMonitors()
            # Primary monitor usually has MONITORINFOF_PRIMARY flag
            for hMonitor, _, _ in monitors:
                info = win32api.GetMonitorInfo(hMonitor)
                if info and info.get('Flags') == win32con.MONITORINFOF_PRIMARY:
                    rc_monitor = info.get('Monitor')
                    if rc_monitor:
                        left, top, right, bottom = rc_monitor
                        width = right - left
                        height = bottom - top
                        log_debug(f"Primary Monitor: L={left} T={top} W={width} H={height}")
                        return {"left": left, "top": top, "width": width, "height": height}
            
            # Fallback if primary flag not found, use first monitor or system metrics
            log_extended("Primary monitor flag not found, using first available or SM_CXSCREEN.")
            if monitors:
                 info = win32api.GetMonitorInfo(monitors[0][0])
                 rc_monitor = info.get('Monitor')
                 if rc_monitor:
                    left, top, right, bottom = rc_monitor
                    return {"left": left, "top": top, "width": right - left, "height": bottom - top}

        except Exception as e:
            log_error(f"Error getting primary monitor info: {e}\n{traceback.format_exc()}")
        
        # Last fallback to system metrics if EnumDisplayMonitors fails or gives no primary
        try:
            return {
                "left": 0, "top": 0,
                "width": win32api.GetSystemMetrics(win32con.SM_CXSCREEN),
                "height": win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            }
        except Exception as e_sysmets:
            log_error(f"Error getting system metrics for screen size: {e_sysmets}")
            return None


    def create_or_update_status_bar(self):
        if not self.enabled or not WINDOWS_FEATURES_AVAILABLE:
            self.destroy_status_bar()
            return

        self.destroy_status_bar() # Destroy existing first
        log_debug("Attempting Opaque Click-Through status bar creation...")

        try:
            primary_info = self._get_primary_monitor_info()
            if not primary_info:
                log_error("Failed to get monitor info for status bar. Aborting creation.")
                return

            screen_left = primary_info["left"]
            screen_top = primary_info["top"]
            screen_width = primary_info["width"]
            screen_height = primary_info["height"]
            
            log_debug(f"Screen metrics for bar: L={screen_left} T={screen_top} W={screen_width} H={screen_height}")

            self.status_bar_window = tk.Toplevel(self.root)
            self.status_bar_window.overrideredirect(True)
            self.status_bar_window.attributes("-topmost", True)
            try:
                self.status_bar_window.attributes("-toolwindow", True) # Makes it not appear in Alt-Tab
            except tk.TclError:
                log_extended("Could not set -toolwindow attribute for status bar.")

            bar_x, bar_y, bar_w, bar_h = 0, 0, 0, 0
            if self.position == "Top":
                bar_x, bar_y, bar_w, bar_h = screen_left, screen_top, screen_width, self.size
            elif self.position == "Bottom":
                bar_x, bar_y, bar_w, bar_h = screen_left, screen_top + screen_height - self.size, screen_width, self.size
            elif self.position == "Left":
                bar_x, bar_y, bar_w, bar_h = screen_left, screen_top, self.size, screen_height
            elif self.position == "Right":
                bar_x, bar_y, bar_w, bar_h = screen_left + screen_width - self.size, screen_top, self.size, screen_height
            
            geom = f"{bar_w}x{bar_h}+{bar_x}+{bar_y}"
            log_debug(f"Status bar calculated geometry: {geom}")
            self.status_bar_window.geometry(geom)
            
            # Frame color will be set by update_bar_color
            self.status_bar_frame = tk.Frame(self.status_bar_window, bg=self.current_bar_color_hex)
            self.status_bar_frame.pack(fill=tk.BOTH, expand=True)

            hwnd = self.status_bar_window.winfo_id()
            log_debug(f"Status bar HWND: {hwnd}")
            
            # Set window styles for click-through
            current_ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            new_ex_style = current_ex_style | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_ex_style)
            
            # Set transparency: 0 for key color (fully transparent), 255 for alpha (fully opaque)
            # LWA_ALPHA makes it opaque but respects WS_EX_TRANSPARENT for click-through
            win32gui.SetLayeredWindowAttributes(hwnd, 0, 255, win32con.LWA_ALPHA)
            
            # Force update window style
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                 win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED)
            
            log_debug(f"Status bar created/updated: Pos={self.position}, Size={self.size}")
            self.update_bar_color(self.current_bar_color_hex) # Apply initial color

        except Exception as e:
            log_error(f"Error creating Windows edge status bar: {e}\n{traceback.format_exc()}")
            self.destroy_status_bar() # Clean up if creation failed

    def update_bar_color(self, new_color_hex: str):
        self.current_bar_color_hex = new_color_hex
        if self.status_bar_frame and self.status_bar_frame.winfo_exists() and self.enabled:
            try:
                self.status_bar_frame.config(background=new_color_hex)
            except tk.TclError: # Window might be destroying
                pass
            except Exception as e:
                log_error(f"Error updating status bar color: {e}")