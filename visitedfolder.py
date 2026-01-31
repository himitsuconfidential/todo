import JLParser
import os
class obj:
    input_file=os.path.join(os.environ["APPDATA"], "Microsoft\\Windows\\Recent\\AutomaticDestinations\\")+'f01b4d95cf55d32a.automaticDestinations-ms'
    input_dir=None
    output_format=None
    output_file=None
    appids_file=None
    delimiter=None
    pretty=True
    quiet=False

import sys
import winshell
import datetime
import re
import tkinter as tk
from tkinter import ttk, messagebox
import json
def get_recent_folders():
    recent_path = JLParser.JL(obj, r"JLParser_AppID.csv")
    folder_entries = tuple(recent_path.folders_with_time.items())
    
    seen = {}
    for folder, ts in folder_entries:
        if folder not in seen or ts > seen[folder]:
            seen[folder] = ts

    sorted_folders = sorted(seen.items(), key=lambda x: x[1], reverse=True)
    return sorted_folders

class RecentFoldersApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Recently Visited Folders")

        self.folders = get_recent_folders()
        self.filtered = self.folders

        # Instruction text
        tk.Label(root,
                 text="Type to filter. Use ↑/↓ to navigate. Press <Enter> to open. Press <Shift+Enter> to open parent folder. Press <Esc> to quit",
                 font=("Arial", 10),
                 anchor="w").pack(fill="x", padx=10, pady=(10,5))

        # Filter input
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self.update_list)
        tk.Label(root, text="Filter:", font=("Arial", 11)).pack(anchor="w", padx=10)
        self.entry = tk.Entry(root, textvariable=self.filter_var, font=("Arial", 11))
        self.entry.pack(fill="x", padx=10, pady=(0,5))
        self.entry.focus_set()

        # Regex checkbox
        self.regex_mode = tk.BooleanVar(value=False)
        self.regex_check = tk.Checkbutton(root, text="Advanced: filter by regex",
                                          variable=self.regex_mode,
                                          command=self.update_list,
                                          font=("Arial", 10),
                                          underline=20)
        self.regex_check.pack(anchor="w", padx=10, pady=(0,10))

        # Centered bold header
        tk.Label(root,
                 text="Recently visited folders",
                 font=("Arial", 14, "bold"),
                 anchor="center").pack(fill="x", pady=(0,5))

        # Frame for Treeview + scrollbar
        frame = tk.Frame(root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Treeview with two columns
        self.tree = ttk.Treeview(frame, columns=("Folder", "Time"), show="headings")
        self.tree.heading("Folder", text="Folder")
        self.tree.heading("Time", text="Last access time")
        self.tree.column("Folder", anchor="w", width=600)
        self.tree.column("Time", anchor="center", width=150)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Buttons frame
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=10, pady=(5,10))

        self.btn_open = tk.Button(btn_frame, text="Open", font=("Arial", 11), command=self.open_selected, underline=0)
        self.btn_open.pack(side="left", padx=5)

        self.btn_open_parent = tk.Button(btn_frame, text="Open Parent Folder", font=("Arial", 11), command=self.open_parent, underline=12)
        self.btn_open_parent.pack(side="left", padx=5)

        self.btn_exit = tk.Button(btn_frame, text="Exit", font=("Arial", 11), command=self.root.quit, underline=1)
        self.btn_exit.pack(side="right", padx=5)

        # Bind keys and double-click
        self.root.bind("<Up>", self.move_up)
        self.root.bind("<Down>", self.move_down)
        self.root.bind("<Prior>", self.page_up)     # PageUp
        self.root.bind("<Next>", self.page_down)   # PageDown
        self.root.bind("<Return>", self.open_selected)
        self.root.bind("<Shift-Return>", self.open_parent)
        self.root.bind("<Escape>", lambda e: self.root.quit())
        self.tree.bind("<Double-1>", self.open_selected)

        # Alt accelerators
        self.root.bind("<Alt-o>", lambda e: self.open_selected())
        self.root.bind("<Alt-f>", lambda e: self.open_parent())
        self.root.bind("<Alt-r>", self.toggle_regex)
        self.root.bind("<Alt-x>", lambda e: self.root.quit())
        
        # Mouse wheel scrolling
        self.tree.bind("<MouseWheel>", self.on_mousewheel)
        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_list()

    def update_list(self, *args):
        filter_text = self.filter_var.get().strip()

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.filtered = []
        for folder, ts in self.folders:
            match = True
            if filter_text:
                if self.regex_mode.get():
                    try:
                        regex = re.compile(filter_text, re.IGNORECASE)
                        match = bool(regex.search(folder))
                    except re.error:
                        match = False
                else:
                    match = filter_text.lower() in folder.lower()

            if match:
                self.filtered.append((folder, ts))

        if not self.filtered:
            self.tree.insert("", "end", values=("No folders match the filter.", ""))
            return

        for folder, ts in self.filtered:
            dt_str = datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
            self.tree.insert("", "end", values=(folder, dt_str))

        # Select first row by default
        if self.filtered:
            first_item = self.tree.get_children()[0]
            self.tree.selection_set(first_item)
            self.tree.focus(first_item)
            self.tree.see(first_item)  # ensure visible

    def toggle_regex(self, event=None): 
        """Toggle regex mode when Alt+R is pressed.""" 
        self.regex_mode.set(not self.regex_mode.get()) 
        self.update_list()
    def move_up(self, event=None):
        sel = self.tree.selection()
        if sel:
            index = self.tree.index(sel[0])
            if index > 0:
                prev_item = self.tree.get_children()[index - 1]
                self.tree.selection_set(prev_item)
                self.tree.focus(prev_item)
                self.tree.see(prev_item)  # auto-scroll

    def move_down(self, event=None):
        sel = self.tree.selection()
        if sel:
            index = self.tree.index(sel[0])
            children = self.tree.get_children()
            if index < len(children) - 1:
                next_item = children[index + 1]
                self.tree.selection_set(next_item)
                self.tree.focus(next_item)
                self.tree.see(next_item)  # auto-scroll

    def page_up(self, event=None):
        sel = self.tree.selection()
        if sel:
            index = self.tree.index(sel[0])
            children = self.tree.get_children()
            visible = int(self.tree.winfo_height() / 20)  # rough estimate rows per page
            new_index = max(0, index - visible)
            target_item = children[new_index]
            self.tree.selection_set(target_item)
            self.tree.focus(target_item)
            self.tree.see(target_item)

    def page_down(self, event=None):
        sel = self.tree.selection()
        if sel:
            index = self.tree.index(sel[0])
            children = self.tree.get_children()
            visible = int(self.tree.winfo_height() / 20)  # rough estimate rows per page
            new_index = min(len(children) - 1, index + visible)
            target_item = children[new_index]
            self.tree.selection_set(target_item)
            self.tree.focus(target_item)
            self.tree.see(target_item)

    def open_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        values = self.tree.item(item, "values")
        folder = values[0]
        if not folder or folder == "No folders match the filter.":
            return
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open {folder}:\n{e}")

    def on_mousewheel(self, event):
        # Scroll Treeview with mouse wheel
        self.tree.yview_scroll(int(-1*(event.delta/120)), "units")

    def open_parent(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        values = self.tree.item(item, "values")
        folder = values[0]
        if not folder or folder == "No folders match the filter.":
            return
        parent = os.path.dirname(folder)
        if not parent:
            return
        try:
            os.startfile(parent)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open parent folder:\n")

    def save_config(self):
        config = {
            "name_filter": self.filter_var.get().strip(),
            "regex_mode": self.regex_mode.get()
        }
        try:
            with open("recentfolder.config", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)
        except Exception:
            pass  # silent fail - don't disturb user

    def load_config(self):
        if not os.path.exists("recentfolder.config"):
            return
        try:
            with open("recentfolder.config", "r", encoding="utf-8") as f:
                config = json.load(f)

            if "name_filter" in config:
                self.filter_var.set(config["name_filter"])
            if "regex_mode" in config:
                self.regex_mode.set(bool(config["regex_mode"]))

            # Trigger update after loading values
            self.update_list()
        except Exception:
            pass  # silent fail
    def on_closing(self):
        self.save_config()
        self.root.destroy()

def main():
    if sys.platform != "win32":
        print("This script only works on Windows.")
        return

    root = tk.Tk()
    root.state("zoomed")   # maximize window on startup
    app = RecentFoldersApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
