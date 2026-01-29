import os
import sys
import winshell
import subprocess
import msvcrt  # Windows-only for key handling
import re

def clear_screen():
    os.system("cls")

def get_recent_excel_files():
    recent_path = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Recent")
    excel_exts = (".xls", ".xlsx", ".xlsm", ".xlsb")
    files = []
    for item in os.listdir(recent_path):
        if item.endswith(".lnk"):
            shortcut_path = os.path.join(recent_path, item)
            try:
                shortcut = winshell.shortcut(shortcut_path)
                target = shortcut.path
                if target and target.lower().endswith(excel_exts) and os.path.isfile(target):
                    files.append(target)
            except Exception:
                continue
    return list(dict.fromkeys(files))  # remove duplicates, preserve order

def show_menu(files, filter_text, selected_index):
    clear_screen()
    print("=== Recently opened Excel files ===")
    print("Type to filter. Use ↑/↓ to navigate. Press <Enter> to open. Press <Esc> to quit.\n")
    print(f"Filter Input: {filter_text}\n")

    try:
        regex = re.compile(filter_text, re.IGNORECASE) if filter_text else None
    except re.error:
        regex = None

    filtered = []
    for f in files:
        if regex is None or regex.search(f):
            filtered.append(f)

    if not filtered:
        print("No Excel files match the filter.\n")
    else:
        for i, f in enumerate(filtered):
            prefix = "-> " if i == selected_index else "   "
            print(f"{prefix}{i+1}. {f}")

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
