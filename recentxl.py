import os
import sys
import winshell
import subprocess
import msvcrt  # Windows-only for key handling
import re
import datetime
def clear_screen():
    os.system("cls")

def get_recent_excel_files():
    recent_path = os.path.join(os.environ["APPDATA"], r"Microsoft\\Windows\\Recent")
    excel_exts = (".xls", ".xlsx", ".xlsm", ".xlsb")
    file_entries = []
    for item in os.listdir(recent_path):
        if item.endswith(".lnk"):
            shortcut_path = os.path.join(recent_path, item)
            try:
                shortcut = winshell.shortcut(shortcut_path)
                target = shortcut.path
                if target and target.lower().endswith(excel_exts) and os.path.isfile(target):
                    # Use shortcut file's modified time as "recent access time"
                    access_time = os.path.getmtime(shortcut_path)
                    file_entries.append((target, access_time))
            except Exception:
                continue

    # Deduplicate: keep the most recent timestamp for each file
    seen = {}
    for fpath, ts in file_entries:
        if fpath not in seen or ts > seen[fpath]:
            seen[fpath] = ts

    # Sort by timestamp descending (most recent first)
    sorted_files = sorted(seen.items(), key=lambda x: x[1], reverse=True)

    # Return just the file paths in sorted order
    return sorted_files

# ANSI escape code for color text
RED = "\033[31m"
GREEN = "\033[92m"
RESET = "\033[0m"

def show_menu(files_with_time, filter_text, selected_index):
    clear_screen()
    print("=== Recently opened Excel files ===")
    print("Type to filter. Use ↑/↓ to navigate. Press <Enter> to open. Press <Esc> to quit.\n")

    try:
        regex = re.compile(filter_text, re.IGNORECASE) if filter_text else None
    except re.error:
        regex = None

    filtered = []
    for file, ts in files_with_time:
        if regex is None or regex.search(file):
            filtered.append((file, ts))

    if not filtered:
        print("No files match the filter.\n")
    else:
        for i, (file, ts) in enumerate(filtered):
            prefix = "-> " if i == selected_index else "   "
            # Format timestamp
            dt_str = datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
            # Extract file name only (last part of path)
            file_name = os.path.basename(file)
            # Highlight file name in green and time in red
            highlighted = file.replace(file_name, f"{GREEN}{file_name}{RESET}")
            print(f"{prefix}{i+1}. {highlighted} {RED}[{dt_str}]{RESET}")

    print(f"Filter Input: {filter_text}")
    return filtered

def open_file(path):
    try:
        subprocess.Popen(f'explorer "{path}"')
    except Exception as e:
        print(f"Failed to open {path}: {e}")

def main():
    files = get_recent_excel_files()
    if not files:
        print("No recent Excel files found.")
        return

    filter_text = ""
    selected_index = 0

    while True:
        filtered = show_menu(files, filter_text, selected_index)

        ch = msvcrt.getwch()
        if ch == "\x1b":  # ESC key
            return
        elif ch == "\r":  # Enter
            if filtered:
                open_file(filtered[selected_index])
        elif ch == "\x08":  # Backspace
            if filter_text:
                filter_text = filter_text[:-1]
                selected_index = 0
        elif ch in ("\xe0", "\000"):  # Arrow keys
            arrow = msvcrt.getwch()
            if filtered:
                if arrow == "H":  # Up
                    selected_index = max(0, selected_index - 1)
                elif arrow == "P":  # Down
                    selected_index = min(len(filtered) - 1, selected_index + 1)
        else:
            # normal character input
            filter_text += ch
            selected_index = 0  # reset selection to top after typing

if __name__ == "__main__":
    if sys.platform != "win32":
        print("This script only works on Windows.")
    else:
        main()
