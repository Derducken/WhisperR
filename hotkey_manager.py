import tkinter as tk
from tkinter import messagebox, ttk
import keyboard # type: ignore
from typing import Callable, Dict, Optional
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug, log_warning 

class HotkeyRecorderDialog(tk.Toplevel):
    def __init__(self, parent, title="Record Hotkey"):
        super().__init__(parent)
        self.transient(parent)
        self.title(title)
        self.geometry("300x150")
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel) # Handle window close X button

        self.hotkey_string = None
        self.recording_active = False

        self.parent_app = parent # To access theme manager if needed

        # Apply theme if parent has theme_manager
        if hasattr(self.parent_app, 'theme_manager') and hasattr(self.parent_app, 'settings_manager'):
            current_theme = self.parent_app.settings_manager.settings.ui_theme
            colors = self.parent_app.theme_manager.get_current_colors(self, current_theme)
            self.configure(bg=colors["bg"])
            
            ttk.Label(self, text="Press the desired hotkey combination.\nEsc to cancel.",
                      background=colors["bg"], foreground=colors["fg"], justify=tk.CENTER
                     ).pack(pady=20, padx=10, expand=True, fill=tk.BOTH)
            
            button_frame = ttk.Frame(self, style='TFrame') # Ensure TFrame picks up theme
            button_frame.pack(pady=10)

            self.ok_button = ttk.Button(button_frame, text="OK", command=self._on_ok, state=tk.DISABLED)
            self.ok_button.pack(side=tk.LEFT, padx=5)
            self.cancel_button = ttk.Button(button_frame, text="Cancel", command=self._on_cancel)
            self.cancel_button.pack(side=tk.LEFT, padx=5)
        else: # Fallback basic styling
            ttk.Label(self, text="Press the desired hotkey combination.\nEsc to cancel.", justify=tk.CENTER).pack(pady=20, padx=10, expand=True, fill=tk.BOTH)
            self.ok_button = ttk.Button(self, text="OK", command=self._on_ok, state=tk.DISABLED)
            self.ok_button.pack(pady=5)
            self.cancel_button = ttk.Button(self, text="Cancel", command=self._on_cancel)
            self.cancel_button.pack(pady=5)


        self.status_var = tk.StringVar(value="Recording...")
        ttk.Label(self, textvariable=self.status_var,
                  background=colors.get("bg") if 'colors' in locals() else None,
                  foreground=colors.get("fg") if 'colors' in locals() else None 
                 ).pack(pady=5)

        self.bind("<Escape>", lambda e: self._on_cancel())
        self._start_recording()

        self.lift()
        self.focus_force()
        self.wait_window()

    def _start_recording(self):
        self.recording_active = True
        try:
            # Using keyboard.read_hotkey() blocks, so we need to handle it carefully
            # or use keyboard.hook with a callback to detect key presses.
            # For simplicity in a dialog, keyboard.read_hotkey in a thread
            # might be an option, but can be tricky with Tkinter.
            # A simpler approach: use keyboard.on_press and build the string.

            # For now, let's try a non-blocking approach with keyboard.hook
            # This is more complex than read_hotkey but better for UI.
            # keyboard.read_hotkey() is blocking, making it unsuitable for direct use in Tkinter event loop.
            # Instead, we can use a simpler text entry for now or advanced hook.
            # For this iteration, we'll rely on user typing or a future improvement for live capture.
            # The prompt tells user to press keys, but we don't actively capture them in THIS dialog version
            # yet. This dialog becomes more of a placeholder until live capture is added.
            # OR, we could use a very short `keyboard.read_hotkey(suppress=True)` in a non-blocking way,
            # but that's hard.

            # Let's simulate: The actual recording will happen in HotkeyManager or config window
            # This dialog will just show the result or prompt for it.
            # For now, this dialog is more of a placeholder.
            # Actual recording should happen where the "Record" button is pressed.
            # This is a bit of a change from the original thought:
            # The dialog itself doesn't record; it's called *after* a recording attempt.
            
            # Re-thinking: the dialog *should* be the one capturing.
            # We'll use a short-lived hook.
            self.recorded_keys = []
            self.hook_id = keyboard.hook(self._key_event_callback, suppress=True)
            self.status_var.set("Press keys... (Esc to cancel, Enter to accept)")

        except Exception as e:
            log_error(f"Failed to start hotkey recording: {e}")
            self.status_var.set(f"Error: {e}")
            self.recording_active = False

    def _key_event_callback(self, event: keyboard.KeyboardEvent):
        if not self.recording_active:
            return

        if event.event_type == keyboard.KEY_DOWN:
            key_name = event.name
            if key_name == 'esc':
                self._on_cancel()
                return
            if key_name == 'enter': # Use Enter to confirm the sequence
                if self.recorded_keys:
                    self._on_ok()
                return

            if key_name not in ['enter', 'esc']: # don't add meta keys for finishing
                if key_name not in self.recorded_keys: # Add key only if it's not already in the list
                    self.recorded_keys.append(key_name)
            
            # Normalize hotkey string
            # Use a set for modifiers found, to add them in a canonical order later.
            # Canonical forms for add_hotkey: 'ctrl', 'alt', 'shift', 'win'.
            pressed_keys_unique = list(dict.fromkeys(self.recorded_keys)) # Deduplicate while preserving order

            found_modifiers_set = set()
            action_part_keys = [] 

            for key in pressed_keys_unique:
                # Check against canonical names and sided variants
                if key in ('left ctrl', 'right ctrl', 'ctrl'):
                    found_modifiers_set.add('ctrl')
                elif key in ('left alt', 'right alt', 'alt', 'alt gr'): # alt gr is a common variant
                    found_modifiers_set.add('alt')
                elif key in ('left shift', 'right shift', 'shift'):
                    found_modifiers_set.add('shift')
                elif key in ('left windows', 'right windows', 'win', 'cmd', 'left command', 'right command', 'super'): # 'super' is linux term for win/cmd
                    found_modifiers_set.add('win')
                # Check if the key is any known modifier type using keyboard.all_modifiers
                # This helps catch if a key like 'control' (if event.name ever returns that) is pressed
                elif key in keyboard.all_modifiers: 
                    # Try to map it to a canonical one if it's a variant not explicitly handled above
                    if 'ctrl' in key: found_modifiers_set.add('ctrl')
                    elif 'alt' in key: found_modifiers_set.add('alt')
                    elif 'shift' in key: found_modifiers_set.add('shift')
                    elif 'win' in key or 'cmd' in key or 'super' in key: found_modifiers_set.add('win')
                    # If it's a modifier but not mappable here, it might be ignored or become an action key.
                else: # Not a recognized modifier, so it's an action key
                    action_part_keys.append(key)

            # Construct the final list in canonical order: ctrl, alt, shift, win, then action key
            final_hotkey_parts = []
            if 'ctrl' in found_modifiers_set: final_hotkey_parts.append('ctrl')
            if 'alt' in found_modifiers_set: final_hotkey_parts.append('alt')
            if 'shift' in found_modifiers_set: final_hotkey_parts.append('shift')
            if 'win' in found_modifiers_set: final_hotkey_parts.append('win')
            
            if action_part_keys:
                # If multiple action keys were recorded (e.g., user rolled fingers A then S quickly),
                # take the last one. This is arbitrary but a common choice for hotkey recorders.
                final_hotkey_parts.append(action_part_keys[-1])
            
            self.hotkey_string = "+".join(final_hotkey_parts)
            self.status_var.set(f"Recorded: {self.hotkey_string}")
            if self.hotkey_string:
                self.ok_button.config(state=tk.NORMAL)


    def _stop_recording(self):
        if self.recording_active:
            try:
                if hasattr(self, 'hook_id') and self.hook_id is not None:
                    keyboard.unhook(self.hook_id)
            except Exception as e:
                 log_extended(f"Error unhooking keyboard: {e}")
            self.recording_active = False

    def _on_ok(self):
        self._stop_recording()
        if not self.hotkey_string: # Should not happen if button is enabled
            messagebox.showwarning("No Hotkey", "No hotkey was recorded.", parent=self)
            self.hotkey_string = None # Ensure it's None if dialog is cancelled early
            # self._start_recording() # Optionally restart
            return
        self.destroy()

    def _on_cancel(self):
        self._stop_recording()
        self.hotkey_string = None
        self.destroy()

    @staticmethod
    def record(parent) -> Optional[str]:
        dialog = HotkeyRecorderDialog(parent)
        return dialog.hotkey_string


class HotkeyManager:
    def __init__(self, root_tk_instance: tk.Tk):
        self.root = root_tk_instance # For dialog parent and theme access
        self._registered_hotkeys: Dict[str, Callable] = {}
        self.hotkey_toggle_record_cb: Optional[Callable] = None
        self.hotkey_show_window_cb: Optional[Callable] = None
        self.current_toggle_hk_str: str = ""
        self.current_show_hk_str: str = ""

    def set_callbacks(self, toggle_record_cb: Callable, show_window_cb: Callable):
        self.hotkey_toggle_record_cb = toggle_record_cb
        self.hotkey_show_window_cb = show_window_cb

    def update_hotkeys(self, toggle_record_str: str, show_window_str: str) -> bool:
        """
        Unregisters old hotkeys and registers new ones.
        Returns True if successful, False otherwise.
        """
        log_extended("Updating hotkeys...")
        self._unregister_all_hotkeys()

        self.current_toggle_hk_str = toggle_record_str.strip().lower()
        self.current_show_hk_str = show_window_str.strip().lower()
        
        registration_errors = []

        if self.current_toggle_hk_str and self.hotkey_toggle_record_cb:
            if not self._register_hotkey(self.current_toggle_hk_str, self.hotkey_toggle_record_cb):
                registration_errors.append(f"Toggle Record Hotkey: '{self.current_toggle_hk_str}'")
        
        if self.current_show_hk_str and self.hotkey_show_window_cb:
            if not self._register_hotkey(self.current_show_hk_str, self.hotkey_show_window_cb):
                registration_errors.append(f"Show Window Hotkey: '{self.current_show_hk_str}'")

        if registration_errors:
            error_msg_details = "\n- ".join(registration_errors)
            log_error(f"Failed to register one or more hotkeys:\n- {error_msg_details}")
            # Show error only if the config window is not active, to avoid double popups
            # This check needs to be done by the caller (e.g. main_app)
            return False
        
        log_essential(f"Hotkeys updated. Toggle: '{self.current_toggle_hk_str}', Show: '{self.current_show_hk_str}'")
        return True

    def _register_hotkey(self, key_string: str, callback: Callable) -> bool:
        if not key_string or not callable(callback):
            return False
        try:
            # The `keyboard` library's add_hotkey might try to acquire OS-level hooks.
            # It's possible for this to fail if permissions are insufficient or
            # if there are conflicts with other low-level keyboard hooks.
            keyboard.add_hotkey(key_string, callback, suppress=False) # suppress=False means event is not blocked
            self._registered_hotkeys[key_string] = callback
            log_essential(f"Registered hotkey: {key_string}")
            return True
        except Exception as e:
            # Distinguish between bad syntax (ValueError from keyboard.parse_hotkey)
            # and system-level registration failure (e.g. ImportError if backend fails, or other OS error)
            if isinstance(e, ValueError): # Likely bad hotkey syntax
                 log_error(f"Invalid hotkey syntax for '{key_string}': {e}")
            else:
                 log_error(f"Failed to register hotkey '{key_string}' at system level: {e}")
            return False

    def _unregister_all_hotkeys(self):
        # Iterate over a copy of keys since remove_hotkey might modify the underlying registration
        for key_string in list(self._registered_hotkeys.keys()):
            try:
                keyboard.remove_hotkey(key_string)
                log_extended(f"Unregistered hotkey: {key_string}")
            except Exception as e: # Can be KeyError if already removed, or other errors
                log_extended(f"Error unregistering hotkey '{key_string}' (may already be removed or invalid): {e}")
        self._registered_hotkeys = {}
        # keyboard.unhook_all() # More aggressive, clears all keyboard hooks by this process

    def cleanup(self):
        log_extended("Cleaning up hotkey manager...")
        self._unregister_all_hotkeys()
        # keyboard.unhook_all() # Call this to be absolutely sure all hooks by this app are gone.

    @staticmethod
    def record_new_hotkey(parent_tk_window) -> Optional[str]:
        """Opens the dialog to record a new hotkey."""
        # Pass the main app window or config window as parent
        return HotkeyRecorderDialog.record(parent_tk_window)
