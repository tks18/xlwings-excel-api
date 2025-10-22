# pq_manager.py
import os
import csv
import json
import threading
import pyperclip
import xlwings as xw

from pq_manager.helpers import parse_pq_file, build_index, read_index, INDEX_FILENAME
from pq_manager.ui import PQManagerUI


def insert_pq(name, root):
    """
    Insert the PQ (by name) into ActiveWorkbook.Queries.Add(Name, Formula)
    """
    root = os.path.abspath(root)
    index = read_index(root)
    match = next((x for x in index if x["name"] == name), None)
    if not match:
        raise FileNotFoundError(f"{name} not found in index at {root}")
    pq_path = match["path"]
    if not os.path.exists(pq_path):
        raise FileNotFoundError(pq_path)
    parsed = parse_pq_file(pq_path)
    m_code = parsed["body"]
    # access excel COM
    app = xw.apps.active
    if app is None:
        raise RuntimeError(
            "No active Excel instance (xlwings couldn't find Excel)")
    excel = app.api
    active_wb = excel.ActiveWorkbook

    # remove existing query(s) with same name (if any)
    try:
        queries = active_wb.Queries
        # iterate backward
        i = queries.Count
        while i >= 1:
            try:
                q = queries.Item(i)
                if q.Name == name:
                    q.Delete()
                i -= 1
            except Exception:
                i -= 1
    except Exception:
        # some Excel versions might not expose .Queries; let the next Add attempt throw a helpful error
        pass
    # add query
    try:
        active_wb.Queries.Add(Name=name, Formula=m_code,
                              Description=parsed["description"])
    except Exception as e:
        raise RuntimeError(f"Failed to add Query in Excel: {e}")
    return {"status": "ok", "name": name}


def copy_pq_function(name, root):
    """
    Copy the PQ Function into Clipboard, 
    Makes it easy to paste in power bi power query editor 
    since it doesnt give any direct interface to insert
    """
    root = os.path.abspath(root)
    index = read_index(root)
    match = next((x for x in index if x["name"] == name), None)
    if not match:
        raise FileNotFoundError(f"{name} not found in index at {root}")
    pq_path = match["path"]
    if not os.path.exists(pq_path):
        raise FileNotFoundError(pq_path)
    parsed = parse_pq_file(pq_path)
    m_code = parsed["body"]
    name = parsed["name"]

    func_to_copy = f"""
    // {name}
    {m_code}
    """

    # Copy query
    try:
        pyperclip.copy(func_to_copy.strip())
    except Exception as e:
        raise RuntimeError(f"Failed to add Query in Excel: {e}")
    return {"status": "ok", "name": name}


# Global variable to track running threads
_ui_thread = None


def open_pq_function_selector(root_path: str):
    """
    Enhanced Power Query Function Selector:
      - Multi-select categories via a dropdown menu (Menubutton with checkbuttons).
      - Multi-select functions in the Treeview and insert ALL selected at once.
      - Sort by Category (toggle asc/desc).
      - Dark mode UI.
    """
    def start_ui():
        PQManagerUI(root_path)
    threading.Thread(target=start_ui).start()
