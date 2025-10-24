import os
import subprocess
import sys
import psutil
import win32gui
import win32con
import win32process

from xl_pq_handler import XLPowerQueryHandler


# Default names
INDEX_FILENAME = "index.csv"


def parse_pq_file(path):
    handler = XLPowerQueryHandler(".", index_name=INDEX_FILENAME)
    return handler.parse_pq_file(path)


def read_index(root):
    handler = XLPowerQueryHandler(root=root, index_name=INDEX_FILENAME)
    return handler.read_index()


def build_index(root):
    handler = XLPowerQueryHandler(root=root, index_name=INDEX_FILENAME)
    return handler.build_index()


def insert_pq(name, root):
    handler = XLPowerQueryHandler(root=root, index_name=INDEX_FILENAME)
    return handler.insert_pq_into_active_excel(name)


def copy_pq_function(name, root):
    handler = XLPowerQueryHandler(root=root, index_name=INDEX_FILENAME)
    return handler.copy_pq_function(name)


LOCK_FILE = os.path.join(os.path.dirname(__file__), "ui.lock")


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
                        return None
                except Exception:
                    pass
        except Exception:
            pass

    proc = subprocess.Popen(
        [sys.executable, ui_path, root_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    with open(LOCK_FILE, "w") as f:
        f.write(str(proc.pid))

    return None
