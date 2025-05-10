import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Callable, Optional, Tuple
from app_logger import get_logger, log_error # Ensure log_error is imported
from settings_manager import CommandEntry
from theme_manager import ThemeManager

class CommandEditorWindow(tk.Toplevel):
    def __init__(self, tk_parent, app_instance, # Changed 'parent' to 'tk_parent', added 'app_instance'
                 current_commands: List[CommandEntry],
                 save_callback: Callable[[List[CommandEntry]], None],
                 theme_manager: ThemeManager):
        super().__init__(tk_parent)
        self.app_instance = app_instance # Store the WhisperRApp instance
        self.theme_manager = theme_manager
        # Use app_instance to get settings_manager for theme
        self.current_theme_colors = theme_manager.get_current_colors(tk_parent, self.app_instance.settings_manager.settings.ui_theme)

        self.title("Configure Voice Commands")
        self.geometry("750x500")
        self.transient(tk_parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close_button)

        self.initial_commands = [CommandEntry(voice=c.voice, action=c.action) for c in current_commands]
        self.save_commands_callback = save_callback

        self._apply_theme()
        self._create_widgets()
        self._populate_commands(current_commands)

    def _apply_theme(self):
        self.configure(bg=self.current_theme_colors["bg"])
        style = ttk.Style(self)
        style.configure('CmdEdit.Treeview.Heading', font=('Helvetica', 10, 'bold'),
                        background=self.current_theme_colors["treeheading_bg"],
                        foreground=self.current_theme_colors["treeheading_fg"])
        style.configure('CmdEdit.Treeview',
                        background=self.current_theme_colors["text_bg"],
                        fieldbackground=self.current_theme_colors["text_bg"],
                        foreground=self.current_theme_colors["text_fg"])
        style.map('CmdEdit.Treeview',
                  background=[('selected', self.current_theme_colors["select_bg"])],
                  foreground=[('selected', self.current_theme_colors["select_fg"])])
        style.configure('CmdEdit.TButton', background=self.current_theme_colors["button_bg"], foreground=self.current_theme_colors["button_fg"])
        style.map('CmdEdit.TButton',
            background=[('active', self.current_theme_colors["select_bg"])],
            foreground=[('active', self.current_theme_colors["select_fg"])])

    def _create_widgets(self):
        info_label = ttk.Label(self,
                               text="Define voice triggers and corresponding actions.\nUse ' FF ' (with spaces) as a wildcard for text to be inserted into the action.",
                               justify=tk.LEFT, wraplength=700, style='TLabel',
                               background=self.current_theme_colors["bg"], # Ensure label bg matches
                               foreground=self.current_theme_colors["fg"])
        info_label.pack(pady=(10, 5), padx=10, anchor=tk.W)

        tree_frame = ttk.Frame(self, style='TFrame', padding=(10,0,10,5))
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=("voice", "action"), show="headings", style='CmdEdit.Treeview')
        self.tree.heading("voice", text="Voice Trigger")
        self.tree.heading("action", text="Action / Command")
        self.tree.column("voice", width=300, stretch=tk.YES)
        self.tree.column("action", width=400, stretch=tk.YES)
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self._on_double_click_cell)

        tree_button_frame = ttk.Frame(self, style='TFrame', padding=(10,5))
        tree_button_frame.pack(fill=tk.X)
        ttk.Button(tree_button_frame, text="Add Command", command=self._add_new_command_row, style='CmdEdit.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(tree_button_frame, text="Remove Selected", command=self._remove_selected_command, style='CmdEdit.TButton').pack(side=tk.LEFT, padx=5)

        bottom_button_frame = ttk.Frame(self, style='TFrame', padding=(10,10))
        bottom_button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(bottom_button_frame, text="Save Changes", command=self._save_and_close, style='CmdEdit.TButton').pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_button_frame, text="Cancel", command=self._on_close_button, style='CmdEdit.TButton').pack(side=tk.RIGHT, padx=5)

    def _populate_commands(self, commands: List[CommandEntry]):
        for item_id in self.tree.get_children(): self.tree.delete(item_id) # Clear existing
        for cmd in commands: self.tree.insert("", tk.END, values=(cmd.voice, cmd.action))

    def _add_new_command_row(self):
        item_id = self.tree.insert("", tk.END, values=("New Voice Trigger", "New Action"))
        self.tree.selection_set(item_id); self.tree.focus(item_id)
        self._edit_selected_command(item_id)

    def _remove_selected_command(self):
        selected_items = self.tree.selection()
        if not selected_items: messagebox.showinfo("No Selection", "Please select a command to remove.", parent=self); return
        if messagebox.askyesno("Confirm Removal", f"Remove {len(selected_items)} selected command(s)?", parent=self):
            for item_id in selected_items: self.tree.delete(item_id)

    def _on_double_click_cell(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id: self._edit_selected_command(item_id)

    def _edit_selected_command(self, item_id=None):
        if item_id is None:
            selected_items = self.tree.selection()
            if not selected_items: messagebox.showinfo("No Selection", "Please select a command to edit.", parent=self); return
            item_id = selected_items[0]
        current_values = self.tree.item(item_id, "values")
        edit_dialog = CommandEditDialog(self, current_values[0], current_values[1], self.theme_manager, self.current_theme_colors)
        new_voice, new_action = edit_dialog.get_result()
        if new_voice is not None and new_action is not None:
            self.tree.item(item_id, values=(new_voice, new_action))

    def _get_commands_from_tree(self) -> List[CommandEntry]:
        commands = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, "values")
            voice, action = (values[0] if len(values) > 0 else "").strip(), (values[1] if len(values) > 1 else "").strip()
            if voice and action: commands.append(CommandEntry(voice=voice, action=action))
        return commands

    def _has_changes(self) -> bool:
        current_tree_commands = self._get_commands_from_tree()
        if len(current_tree_commands) != len(self.initial_commands): return True
        set_initial = {(c.voice, c.action) for c in self.initial_commands}
        set_current = {(c.voice, c.action) for c in current_tree_commands}
        return set_initial != set_current

    def _save_and_close(self):
        self.save_commands_callback(self._get_commands_from_tree()); self.destroy()

    def _on_close_button(self):
        if self._has_changes():
            if messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Save them before closing?", parent=self):
                self._save_and_close()
            else: self.destroy()
        else: self.destroy()

class CommandEditDialog(tk.Toplevel):
    def __init__(self, tk_parent, initial_voice: str, initial_action: str, theme_manager, colors): # tk_parent
        super().__init__(tk_parent)
        self.transient(tk_parent)
        self.title("Edit Command")
        self.geometry("500x200")
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.colors = colors
        self.configure(bg=self.colors["bg"])

        self.result_voice: Optional[str] = None
        self.result_action: Optional[str] = None

        main_frame = ttk.Frame(self, padding=15, style='TFrame')
        main_frame.pack(expand=True, fill=tk.BOTH)
        ttk.Label(main_frame, text="Voice Trigger:", style='TLabel', background=self.colors["bg"], foreground=self.colors["fg"]).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.voice_entry_var = tk.StringVar(value=initial_voice)
        voice_entry = ttk.Entry(main_frame, textvariable=self.voice_entry_var, width=50)
        voice_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        ttk.Label(main_frame, text="Action:", style='TLabel', background=self.colors["bg"], foreground=self.colors["fg"]).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.action_entry_var = tk.StringVar(value=initial_action)
        action_entry = ttk.Entry(main_frame, textvariable=self.action_entry_var, width=50)
        action_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        main_frame.columnconfigure(1, weight=1)
        button_frame = ttk.Frame(main_frame, style='TFrame')
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="OK", command=self._on_ok, style='CmdEdit.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=self._on_cancel, style='CmdEdit.TButton').pack(side=tk.LEFT, padx=10)
        voice_entry.focus_set()
        self.wait_window()

    def _on_ok(self):
        self.result_voice = self.voice_entry_var.get().strip()
        self.result_action = self.action_entry_var.get().strip()
        if not self.result_voice or not self.result_action:
            messagebox.showwarning("Missing Info", "Both Voice Trigger and Action must be filled.", parent=self)
            self.result_voice = None; self.result_action = None; return
        self.destroy()

    def _on_cancel(self):
        self.result_voice = None; self.result_action = None; self.destroy()

    def get_result(self) -> Tuple[Optional[str], Optional[str]]:
        return self.result_voice, self.result_action