# pq_manager.py
import os
import csv
import json
import yaml
import pyperclip
import xlwings as xw

# Default names (you can override by passing root arg)
INDEX_FILENAME = "index.csv"


def _safe_str(x):
    if x is None:
        return ""
    return str(x).replace('\r', ' ').replace('\n', ' ').strip()


def parse_pq_file(path):
    with open(path, "r", encoding="utf-8") as fh:
        text = fh.read()
    fm = {}
    body = text
    if text.lstrip().startswith("---"):
        parts = text.split("---", 2)
        if len(parts) >= 3:
            try:
                fm = yaml.safe_load(parts[1]) or {}
            except Exception:
                fm = {}
            body = parts[2].lstrip("\n")
    name = _safe_str(fm.get("name") or os.path.splitext(
        os.path.basename(path))[0])
    category = _safe_str(fm.get("category") or "Uncategorized")
    tags = fm.get("tags") or []
    description = _safe_str(fm.get("description") or "")
    version = _safe_str(fm.get("version") or "")
    return {
        "name": name,
        "category": category,
        "tags": tags,
        "description": description.replace("\n", " ").replace("\r", " "),
        "version": version,
        "path": os.path.abspath(path),
        "body": body
    }


def build_index(root):
    """
    Walk `root` and write index.csv into that folder.
    Columns: name,category,tags,description,version,path
    """
    root = os.path.abspath(root)
    rows = []
    for dirpath, _, files in os.walk(root):
        for fn in files:
            if fn.lower().endswith(".pq"):
                p = os.path.join(dirpath, fn)
                try:
                    parsed = parse_pq_file(p)
                    rows.append(parsed)
                except Exception as e:
                    # skip problematic file but print for debug
                    print(f"Failed to parse {p}: {e}")
    # sort
    rows = sorted(rows, key=lambda r: (
        r["category"].lower(), r["name"].lower()))
    index_path = os.path.join(root, INDEX_FILENAME)
    with open(index_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
        # header
        writer.writerow(["name", "category", "tags",
                        "description", "version", "path"])
        for r in rows:
            # write tags as JSON string (safe within quoted CSV)
            writer.writerow([r["name"], r["category"], json.dumps(
                r["tags"], ensure_ascii=False), r["description"], r["version"], r["path"]])
    return index_path


def read_index(root):
    index_path = os.path.join(root, INDEX_FILENAME)
    if not os.path.exists(index_path):
        return []
    out = []
    with open(index_path, "r", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for r in reader:
            # parse tags back to list if possible
            tags_field = r.get("tags", "")
            try:
                tags = json.loads(tags_field)
            except Exception:
                tags = [tags_field] if tags_field else []
            out.append({
                "name": r.get("name", ""),
                "category": r.get("category", ""),
                "tags": tags,
                "description": r.get("description", ""),
                "version": r.get("version", ""),
                "path": r.get("path", "")
            })
    return out


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
