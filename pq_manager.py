import cmd
import os
import subprocess
import sys
import psutil
import win32gui
import win32con
import pyperclip
import win32process
import json

from xl_pq_handler import PQManager

PACKAGE_NAME = "xl_pq_handler"


def insert_pq(name: str, root: str) -> str:
    """
    Inserts a single PQ (and its dependencies) into the active workbook.
    """
    try:
        handler = PQManager(root)
        result = handler.insert_into_excel([name])
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def build_index(root: str) -> str:
    """
    Rebuilds the index.
    """
    try:
        handler = PQManager(root)
        handler.build_index()
        return json.dumps({"status": "ok", "message": "Index rebuilt successfully."})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def copy_pq_function(name, root):
    print(name, root)
    handler = PQManager(root)
    try:
        pq_data = handler.get_script(name)
        if pq_data:
            pyperclip.copy(
                f'{pq_data.meta.name}\n{pq_data.meta.description}\n{pq_data.body}')
        else:
            pyperclip.copy("")
        return json.dumps({"status": "ok", "message": "Copied to clipboard."})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})


def open_pq_function_selector(root_path: str):
    """
    Launches the new xl_pq_handler UI window.
    - If the UI is already open, brings it to front.
    """
    lock_file = os.path.join(root_path, "ui.lock")

    if os.path.exists(root_path):

        # Check if UI is already running
        if os.path.exists(lock_file):
            try:
                with open(lock_file, "r") as f:
                    pid = int(f.read().strip())

                if psutil.pid_exists(pid):
                    try:
                        hwnd = None
                        # Bring it to front

                        def enum_handler(h, ctx):
                            nonlocal hwnd
                            _, found_pid = win32process.GetWindowThreadProcessId(
                                h)
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
        cmd_list = [
            sys.executable,
            "-m", PACKAGE_NAME,
            root_path
        ]

        cmd_string = subprocess.list2cmdline(cmd_list)

        # 3. Pass the string to Popen (this is robust on Windows)
        proc = subprocess.Popen(
            cmd_string,
            creationflags=subprocess.CREATE_NO_WINDOW,
            close_fds=True
        )
        # --- End Fix ---

        with open(lock_file, "w") as f:
            f.write(str(proc.pid))

        return None

    return None
