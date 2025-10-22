import yaml
import os
import csv
import json

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
