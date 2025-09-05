from lxml import html
import os
import hashlib

# -------------------------
# Cache setup
# -------------------------
import os
import hashlib
import pandas as pd
import requests

CACHE_DIR = r"C:\Tools\Automation Scripts\shan_xlwings_project\_df_cache"
MAX_CACHE_SIZE_MB = 50

if not os.path.exists(CACHE_DIR):
    os.makedirs(CACHE_DIR)

# -------------------------
# Helper: get cache path
# -------------------------


def _get_cache_path(source_name):
    fname = f"{source_name}.pkl"
    return os.path.join(CACHE_DIR, fname)

# -------------------------
# Cache HTML by source_name
# -------------------------


def cache_html(source_name, html_content):
    """Save HTML content as a file named by source_name."""
    fname = f"{source_name}.html"
    fpath = os.path.join(CACHE_DIR, fname)
    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(html_content)
    cleanup_cache()

# -------------------------
# Load HTML by source_name
# -------------------------


def load_html(source_name):
    """Load HTML content from file by source_name."""
    fpath = os.path.join(CACHE_DIR, f"{source_name}.html")
    if os.path.exists(fpath):
        with open(fpath, 'r', encoding='utf-8') as f:
            return f.read()
    return None


# -------------------------
# Cleanup cache directory
# -------------------------


def cleanup_cache():
    """Delete oldest files if cache exceeds max size."""
    total_size = sum(os.path.getsize(os.path.join(CACHE_DIR, f))
                     for f in os.listdir(CACHE_DIR))
    total_mb = total_size / (1024*1024)
    if total_mb <= MAX_CACHE_SIZE_MB:
        return
    files = sorted([os.path.join(CACHE_DIR, f) for f in os.listdir(CACHE_DIR)],
                   key=os.path.getctime)
    while total_mb > MAX_CACHE_SIZE_MB and files:
        os.remove(files[0])
        files.pop(0)
        total_size = sum(os.path.getsize(f) for f in files)
        total_mb = total_size / (1024*1024)


def extract_text_xpath(html_content, xpath_expr):
    """
    Extract first matching text using XPath.
    """
    tree = html.fromstring(html_content)
    result = tree.xpath(xpath_expr)
    if not result:
        return ""
    # If element, get text; if string, return directly
    if isinstance(result[0], html.HtmlElement):
        return result[0].text_content().strip()
    return str(result[0]).strip()


def extract_list_xpath(html_content, xpath_expr):
    """
    Extract multiple elements or attributes using XPath.
    Returns a list of strings.
    """
    tree = html.fromstring(html_content)
    results = tree.xpath(xpath_expr)
    extracted = []
    for r in results:
        if isinstance(r, html.HtmlElement):
            extracted.append(r.text_content().strip())
        else:
            extracted.append(str(r).strip())
    return extracted
