import os
import subprocess
import sys
import psutil
import win32gui
import win32con
import pyperclip
import win32process
import json

from xl_pq_handler import XLPowerQueryHandler


# Default names
INDEX_FILENAME = "index.json"
LOCK_FILE = os.path.join(os.path.dirname(__file__), "ui.lock")


def insert_pq(name: str, root: str) -> str:
    """
    Inserts a single PQ (and its dependencies) into the active workbook.
    """
    try:
        handler = XLPowerQueryHandler(root)
        result = handler.insert_pq_into_active_excel(name)
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def build_index(root: str) -> str:
    """
    Rebuilds the index.
    """
    try:
        handler = XLPowerQueryHandler(root)
        handler.build_index()
        return json.dumps({"status": "ok", "message": "Index rebuilt successfully."})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def parse_pq_file(path):
    handler = XLPowerQueryHandler(".", index_file=INDEX_FILENAME)
    return handler.parse_pq_file(path)


def read_index(root):
    handler = XLPowerQueryHandler(root=root, index_file=INDEX_FILENAME)
    return handler.read_index()


def copy_pq_function(name, root):
    print(name, root)
    handler = XLPowerQueryHandler(root=root, index_file=INDEX_FILENAME)
    try:
        pq_data = handler.get_pq_by_name(name)
        if pq_data:
            pyperclip.copy(
                f'{pq_data.get("name", "")}\n{pq_data.get("description", "")}\n{pq_data.get("body", "")}')
        else:
            pyperclip.copy("")
        return json.dumps({"status": "ok", "message": "Copied to clipboard."})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def open_pq_function_selector(root_path: str):
    """
    Launches the Power Query Function Selector UI (ui.py) in a separate
    process so Excel remains usable while the CTk window is open.
    - If the UI is already open, brings it to front.
    """
    ui_path = os.path.join(os.path.dirname(__file__), "ui.py")

    # Check if UI is already running
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read().strip())

            if psutil.pid_exists(pid):
                try:
                    hwnd = None
                    # Bring it to front

                    def enum_handler(h, ctx):
                        nonlocal hwnd
                        _, found_pid = win32process.GetWindowThreadProcessId(h)
                        if found_pid == pid:
                            hwnd = h

                    win32gui.EnumWindows(enum_handler, None)
                    if hwnd:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(hwnd)
                        return None  # Exit if window was found and brought to front
                except Exception:
                    pass
        except Exception:
            # If lock file is stale or unreadable, just continue to launch
            pass

    # --- FIX: Build a command-line string to handle spaces ---
    # 1. Create the argument list
    cmd_list = [sys.executable, ui_path, root_path]

    # 2. Convert list to a correctly quoted command-line string
    cmd_string = subprocess.list2cmdline(cmd_list)

    # 3. Pass the string to Popen (this is robust on Windows)
    proc = subprocess.Popen(
        cmd_string,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    # --- End Fix ---

    with open(LOCK_FILE, "w") as f:
        f.write(str(proc.pid))

    return None
