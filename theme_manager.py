import tkinter as tk
from tkinter import ttk
from constants import Theme
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning
import platform
import ctypes
from ctypes import wintypes

class ThemeManager:
    def __init__(self):
        self.current_theme_name = "Dark"  # Default theme
        self.themes = {
            Theme.LIGHT.value: {
                "bg": "#F0F0F0",
                "fg": "#000000",
                "text_bg": "#FFFFFF",
                "text_fg": "#000000",
                "select_bg": "#0078D7",
                "select_fg": "#FFFFFF",
                "button_bg": "#E1E1E1",
                "button_fg": "#000000",
                "disabled_fg": "#A0A0A0",
                "treeheading_bg": "#E1E1E1",
                "treeheading_fg": "#000000",
                "oddrow_bg": "#F0F0F0",
                "evenrow_bg": "#E0E0E0",
                "ttk_theme": "clam"
            },
            Theme.DARK.value: {
                "bg": "#2B2B2B",
                "fg": "#D3D3D3",
                "text_bg": "#3C3F41",
                "text_fg": "#A9B7C6",
                "select_bg": "#007ACC",
                "select_fg": "#FFFFFF",
                "button_bg": "#4A4A4A",
                "button_fg": "#D3D3D3",
                "disabled_fg": "#6A6A6A",
                "treeheading_bg": "#3C3F41",
                "treeheading_fg": "#D3D3D3",
                "oddrow_bg": "#2B2B2B",
                "evenrow_bg": "#313335",
                "ttk_theme": "clam"
            }
        }
        self.themes[Theme.LIGHT.value]["TNotebook.Tab"] = {
            "padding": [5, 2],
            "background": [('selected', '#FFFFFF'), ('!selected', '#E1E1E1')],
            "foreground": [('selected', '#000000'), ('!selected', '#333333')]
        }
        self.themes[Theme.DARK.value]["TNotebook.Tab"] = {
            "padding": [5, 2],
            "background": [('selected', '#3C3F41'), ('!selected', '#2B2B2B')],
            "foreground": [('selected', '#D3D3D3'), ('!selected', '#A0A0A0')]
        }


    def get_current_colors(self, root: tk.Tk, theme_name: str) -> dict:
        actual_theme_name = theme_name
        if theme_name == Theme.SYSTEM.value:
            try:
                system_is_dark = root.tk.call("tk::darkmode") 
                actual_theme_name = Theme.DARK.value if system_is_dark else Theme.LIGHT.value
                # Use the imported helper function correctly
                log_extended(f"System theme detected as: {'Dark' if system_is_dark else 'Light'}")
            except tk.TclError:
                # Use the imported helper function correctly
                log_extended("tk::darkmode not available, defaulting System theme to Light.")
                actual_theme_name = Theme.LIGHT.value
        
        return self.themes.get(actual_theme_name, self.themes[Theme.LIGHT.value])


    def apply_theme(self, root: tk.Tk, theme_name: str):
        colors = self.get_current_colors(root, theme_name)
        ttk_theme_base = colors.get("ttk_theme", "clam")

        style = ttk.Style(root)
        try:
            available_themes = style.theme_names()
            if ttk_theme_base not in available_themes:
                for t in ['clam', 'vista', 'xpnative', 'default']:
                    if t in available_themes:
                        ttk_theme_base = t
                        break
            style.theme_use(ttk_theme_base)
            # Use the imported helper function correctly
            log_extended(f"Using ttk base theme: {ttk_theme_base}")
        except tk.TclError as e:
            # Use the imported helper function correctly
            log_error(f"Failed to set ttk base theme {ttk_theme_base}: {e}")
            try: style.theme_use('default')
            except: pass

        root.configure(bg=colors["bg"])

        style.configure('.', background=colors["bg"], foreground=colors["fg"],
                        fieldbackground=colors["text_bg"],
                        selectbackground=colors["select_bg"], selectforeground=colors["select_fg"])
        
        style.map('.', foreground=[('disabled', colors["disabled_fg"])])

        style.configure('TLabel', background=colors["bg"], foreground=colors["fg"])
        style.configure('TButton', background=colors["button_bg"], foreground=colors["button_fg"])
        style.map('TButton',
                  background=[('active', colors["select_bg"]), ('disabled', colors["bg"])],
                  foreground=[('active', colors["select_fg"]), ('disabled', colors["disabled_fg"])])
        
        style.configure('TCheckbutton', background=colors["bg"], foreground=colors["fg"],
                        indicatorcolor=colors["text_bg"])
        style.map('TCheckbutton',
                  background=[('active', colors["bg"])],
                  indicatorcolor=[('selected', colors["select_bg"]), ('!selected', colors["text_bg"])])

        style.configure('TRadiobutton', background=colors["bg"], foreground=colors["fg"],
                        indicatorcolor=colors["text_bg"])
        style.map('TRadiobutton',
                  background=[('active', colors["bg"])],
                  indicatorcolor=[('selected', colors["select_bg"])])

        style.configure('TEntry', fieldbackground=colors["text_bg"], foreground=colors["text_fg"],
                        insertcolor=colors["fg"])
        style.map('TEntry',
                  foreground=[('disabled', colors["disabled_fg"])],
                  fieldbackground=[('disabled', colors["bg"])])

        style.configure('TCombobox', fieldbackground=colors["text_bg"], foreground=colors["text_fg"],
                        selectbackground=colors["select_bg"], selectforeground=colors["select_fg"],
                        arrowcolor=colors["fg"])
        style.map('TCombobox',
                  fieldbackground=[('readonly', colors["text_bg"]), ('disabled', colors["bg"])],
                  foreground=[('readonly', colors["text_fg"]), ('disabled', colors["disabled_fg"])],
                  selectbackground=[('readonly', colors["select_bg"])],
                  selectforeground=[('readonly', colors["select_fg"])])
        
        root.option_add("*TCombobox*Listbox*Background", colors["text_bg"])
        root.option_add("*TCombobox*Listbox*Foreground", colors["text_fg"])
        root.option_add("*TCombobox*Listbox*selectBackground", colors["select_bg"])
        root.option_add("*TCombobox*Listbox*selectForeground", colors["select_fg"])

        style.configure('TFrame', background=colors["bg"])
        style.configure('TLabelframe', background=colors["bg"], foreground=colors["fg"], bordercolor=colors["fg"])
        style.configure('TLabelframe.Label', background=colors["bg"], foreground=colors["fg"])

        notebook_tab_style = colors.get("TNotebook.Tab", {})
        if notebook_tab_style:
             style.configure('TNotebook.Tab', 
                            padding=notebook_tab_style.get("padding", [5,2]))
             style.map('TNotebook.Tab',
                       background=notebook_tab_style.get("background", []),
                       foreground=notebook_tab_style.get("foreground", []))
        style.configure('TNotebook', background=colors["bg"])

        style.configure('TScrollbar', troughcolor=colors["bg"], background=colors["button_bg"],
                        arrowcolor=colors["fg"], bordercolor=colors["bg"])
        style.map('TScrollbar', background=[('active', colors["select_bg"])])

        style.configure("Treeview",
                        background=colors["text_bg"],
                        fieldbackground=colors["text_bg"],
                        foreground=colors["text_fg"])
        style.map("Treeview",
                  background=[('selected', colors["select_bg"])],
                  foreground=[('selected', colors["select_fg"])])
        style.configure("Treeview.Heading",
                        background=colors["treeheading_bg"],
                        foreground=colors["treeheading_fg"],
                        relief=tk.FLAT)
        style.map("Treeview.Heading",
                  background=[('active', colors["select_bg"])]) 
        
        style.configure('Config.TLabelframe', padding=5, borderwidth=1, relief=tk.SOLID,
                        background=colors["bg"], bordercolor=colors["disabled_fg"])
        style.map("Config.TLabelframe", bordercolor=[('active', colors["fg"])])
        style.configure("Config.TLabelframe.Label", padding=(10, 5), background=colors["bg"], foreground=colors["fg"])

        self.update_tk_widget_colors(root, colors)
        
        # Attempt to style the main menubar and its dropdowns
        # Reverted option_add for menus as it caused issues. Sticking to direct configuration.
        root.update_idletasks() # Process pending Tkinter tasks
        try:
            menubar_path = root.cget("menu")
            if menubar_path:
                log_debug(f"Found menubar path for {root}: {menubar_path}")
                menubar_widget = root.nametowidget(menubar_path)
                if isinstance(menubar_widget, tk.Menu):
                    log_debug(f"Applying theme to menubar widget: {menubar_widget}")
                    self._apply_theme_to_menu(menubar_widget, colors)
                else:
                    log_warning(f"Widget at menubar path {menubar_path} is not a tk.Menu: {type(menubar_widget)}")
            else:
                log_debug(f"No menubar path found for {root}.")
        except tk.TclError as e:
            log_warning(f"Could not access or style menubar for {root}: {e}")
        root.update_idletasks() # Process again after styling

        # Attempt to set dark title bar on Windows
        if platform.system() == "Windows":
            try:
                hwnd = root.winfo_id()
                # Determine if dark theme should be applied to title bar
                should_apply_dark_title_bar = False
                if theme_name == Theme.DARK.value:
                    should_apply_dark_title_bar = True
                elif theme_name == Theme.SYSTEM.value:
                    try:
                        system_is_dark = root.tk.call("tk::darkmode") # Returns 1 for dark, 0 for light
                        should_apply_dark_title_bar = bool(system_is_dark)
                        log_debug(f"System theme for title bar: {'Dark' if system_is_dark else 'Light'}")
                    except tk.TclError:
                        log_warning("tk::darkmode call failed for System theme title bar. Defaulting to light title bar.")
                        should_apply_dark_title_bar = False # Default to light if check fails
                
                log_debug(f"Attempting to set Windows title bar dark mode: {should_apply_dark_title_bar} for HWND {hwnd}")
                self._set_windows_dark_title_bar(hwnd, should_apply_dark_title_bar)
            except Exception as e:
                log_error(f"Failed to set Windows dark title bar: {e}", exc_info=True)

        self.current_theme_name = theme_name  # Track current theme
        log_essential(f"Theme '{theme_name}' applied.")

    def _set_windows_dark_title_bar(self, hwnd: int, enable_dark: bool):
        """
        Attempts to set the dark mode for the window title bar on Windows.
        Uses DWMWA_USE_IMMERSIVE_DARK_MODE (value 20 for Win 10 19041+/Win 11, 19 for older Win 10).
        """
        if not platform.system() == "Windows":
            return
        
        try:
            # For Windows 10 build 18985 (version 2004) and later, and Windows 11
            # DWMWA_USE_IMMERSIVE_DARK_MODE:
            # Value 20: Windows 10 20H1 (build 19041) and later, Windows 11.
            # Value 19: Windows 10 19H1 (build 18362) up to 19H2 (build 18363).
            
            attributes_to_try = [20, 19] # Try 20 first, then 19 as a fallback
            success = False
            value = wintypes.DWORD(1 if enable_dark else 0)
            dwmapi = ctypes.WinDLL("dwmapi")
            hwnd_obj = wintypes.HWND(hwnd)
            value_ptr = ctypes.byref(value)
            value_size = ctypes.sizeof(value)

            for attr_val in attributes_to_try:
                log_debug(f"Attempting DwmSetWindowAttribute with attribute {attr_val} for HWND {hwnd}, dark_mode: {enable_dark}")
                attr_obj = wintypes.DWORD(attr_val)
                result = dwmapi.DwmSetWindowAttribute(hwnd_obj, attr_obj, value_ptr, value_size)
                
                if result == 0: # S_OK
                    log_debug(f"DwmSetWindowAttribute(attr={attr_val}) successful for dark title bar ({enable_dark}) on HWND {hwnd}.")
                    success = True
                    break # Stop if successful
                else:
                    # Loglevel changed to debug for non-critical fallback attempt failures
                    log_debug(f"DwmSetWindowAttribute(attr={attr_val}) for dark title bar ({enable_dark}) on HWND {hwnd} returned error code: {result}. Error: {ctypes.WinError(result)}")
            
            if not success:
                log_warning(f"Failed to set dark title bar for HWND {hwnd} using available DWM attributes (20, 19).")

        except (AttributeError, OSError, Exception) as e: # Catch broader exceptions
            log_error(f"Could not set dark title bar via DWM for HWND {hwnd}: {e}", exc_info=True)

    def _apply_theme_to_menu(self, menu_widget: tk.Menu, colors: dict):
        """Recursively applies theme colors to a menu and its submenus."""
        menu_path_name = "UnknownMenu"
        try:
            menu_path_name = menu_widget.winfo_pathname(menu_widget.winfo_id())
        except Exception: pass
        log_debug(f"Attempting to style menu (direct configure): {menu_path_name} with bg: {colors.get('bg', 'N/A')}")

        try:
            menu_widget.configure(
                tearoff=0, # Ensure tearoff is set first
                background=colors["bg"],
                foreground=colors["fg"],
                activebackground=colors["select_bg"],
                activeforeground=colors["select_fg"],
                disabledforeground=colors["disabled_fg"],
                relief=tk.FLAT,
                bd=0,  # Explicitly set borderwidth to 0
                activeborderwidth=0, # Explicitly set active borderwidth to 0
                selectcolor=colors["fg"] # Color for checkbutton/radiobutton indicators
            )
            log_debug(f"Directly configured menu: {menu_path_name}")
        except tk.TclError as e:
            log_warning(f"Error directly configuring menu {menu_path_name}: {e}")
            return # If basic configuration fails, stop for this menu

        last_index = menu_widget.index(tk.END)
        if last_index is not None:
            log_debug(f"Iterating through {last_index + 1} items in menu {menu_path_name}")
            for i in range(last_index + 1):
                try:
                    item_type = menu_widget.type(i)
                    log_debug(f"Menu {menu_path_name} item {i}: type={item_type}")
                    if item_type == "cascade":
                        submenu_path = menu_widget.entrycget(i, "menu")
                        if submenu_path:
                            log_debug(f"Cascade item {i} in {menu_path_name} has submenu path: {submenu_path}")
                            submenu_widget = menu_widget.nametowidget(submenu_path)
                            if isinstance(submenu_widget, tk.Menu):
                                log_debug(f"Recursively styling submenu: {submenu_path}")
                                self._apply_theme_to_menu(submenu_widget, colors)
                            else:
                                log_warning(f"Submenu widget at {submenu_path} is not a tk.Menu: {type(submenu_widget)}")
                        else:
                             log_debug(f"Cascade item {i} in {menu_path_name} has no submenu path.")
                except tk.TclError as e_item:
                    log_warning(f"Error processing menu item at index {i} in menu {menu_path_name}: {e_item}")
                    continue
        else:
            log_debug(f"Menu {menu_path_name} has no items (index END is None).")
                    
    def update_tk_widget_colors(self, parent_widget: tk.Misc, colors: dict):
        widget_map = {
            tk.Text: {"bg": colors["text_bg"], "fg": colors["text_fg"],
                      "insertbackground": colors["fg"], 
                      "selectbackground": colors["select_bg"],
                      "selectforeground": colors["select_fg"]},
            tk.Listbox: {"bg": colors["text_bg"], "fg": colors["text_fg"],
                         "selectbackground": colors["select_bg"],
                         "selectforeground": colors["select_fg"]},
            tk.Frame: {"bg": colors["bg"]}, 
            tk.Label: {"bg": colors["bg"], "fg": colors["fg"]}, 
            tk.Canvas: {"bg": colors["bg"]},
        }

        for child in parent_widget.winfo_children():
            widget_type = type(child)
            if widget_type in widget_map:
                try:
                    child.configure(**widget_map[widget_type])
                except tk.TclError as e:
                    pass 
            
            if isinstance(child, (tk.Frame, ttk.Frame, tk.Toplevel, ttk.Notebook)):
                self.update_tk_widget_colors(child, colors)
