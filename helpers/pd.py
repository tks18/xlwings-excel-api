import inspect
import shutil
import pandas as pd
import ast
import os
from collections import OrderedDict
from typing import Any

# -------------------------
# Global registry + cache
# -------------------------
DF_REGISTRY = OrderedDict()
CACHE_DIR = r"C:\Tools\Automation Scripts\shan_xlwings_project\_df_cache"
CACHE_MAX_SIZE = 200 * 1024 * 1024  # 200 MB
os.makedirs(CACHE_DIR, exist_ok=True)
MEMORY_THRESHOLD = 50_000_000  # 50 MB
LRU_MAX_ITEMS = 3


def get_cache_path(df_name):
    return os.path.join(CACHE_DIR, f"{df_name}.parquet")


def memory_check_and_lru():
    while len(DF_REGISTRY) > LRU_MAX_ITEMS:
        old_name, old_df = DF_REGISTRY.popitem(last=False)
        path = get_cache_path(old_name)
        old_df.to_parquet(path)


def auto_load(df_name):
    if df_name in DF_REGISTRY:
        DF_REGISTRY.move_to_end(df_name)
        return DF_REGISTRY[df_name]
    path = get_cache_path(df_name)
    if os.path.exists(path):
        df = pd.read_parquet(path)
        DF_REGISTRY[df_name] = df
        memory_check_and_lru()
        return df
    raise ValueError(f"DF '{df_name}' not found")


def auto_cache(df_name, df):
    path = get_cache_path(df_name)
    df.to_parquet(path)
    DF_REGISTRY[df_name] = df
    memory_check_and_lru()


def get_dir_size(path):
    """Return directory size in bytes."""
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.isfile(fp):
                total += os.path.getsize(fp)
    return total


def check_cache_dir():
    """Delete cache dir if it exceeds max size."""
    size = get_dir_size(CACHE_DIR)
    if size > CACHE_MAX_SIZE:
        print(
            f"[Cache Cleanup] Cache size {size/1e6:.2f} MB exceeded limit. Clearing...")
        shutil.rmtree(CACHE_DIR)
        os.makedirs(CACHE_DIR, exist_ok=True)


def try_literal_eval(s: str):
    """Try ast.literal_eval, with a fallback to fix Excel-style TRUE/FALSE."""
    if not isinstance(s, str):
        return s
    s_strip = s.strip()
    try:
        return ast.literal_eval(s_strip)
    except Exception:
        # repair common Excel tokens and reattempt
        fixed = (
            s_strip
            .replace("TRUE", "True")
            .replace("True", "True")
            .replace("FALSE", "False")
            .replace("False", "False")
            .replace("NULL", "None")
            .replace("null", "None")
        )
        try:
            return ast.literal_eval(fixed)
        except Exception:
            return s  # fallback to raw string


def normalize_scalar(v: Any):
    """Normalize single scalar values (strings -> bool/int/float/None if possible)."""
    # Already correct types
    if v is None or isinstance(v, (bool, int, float)):
        return v

    # If it's a string, attempt to interpret it
    if isinstance(v, str):
        t = v.strip()
        low = t.lower()
        # booleans
        if low in ("true", "false"):
            return low == "true"
        # none/null
        if low in ("none", "null", "na", "nan", ""):
            return None
        # numeric attempts (int then float)
        try:
            if "." in t:
                return float(t)
            return int(t)
        except Exception:
            # lastly try literal_eval for e.g. "'abc'" or numeric with extra
            le = try_literal_eval(t)
            if not isinstance(le, str) or le != t:
                return le
            return t  # keep as string

    # lists/tuples/dicts are handled above, but keep fallback
    return v


def normalize(obj: Any):
    """
    Recursively normalize:
      - dict -> {k: normalize(v)}
      - list -> [ normalize_scalar(elem) ] (and flatten 2D single-col lists)
      - tuple -> tuple(...)
      - str -> try try_literal_eval then normalize result
    """
    # dict
    if isinstance(obj, dict):
        return {k: normalize(v) for k, v in obj.items()}

    # tuple
    if isinstance(obj, tuple):
        return tuple(normalize(x) for x in obj)

    # list (may be nested because xlwings passes ranges as 2D lists)
    if isinstance(obj, list):
        # flatten 2D single-column: [[a],[b],[c]] -> [a,b,c]
        if all(isinstance(i, list) and len(i) == 1 for i in obj):
            flat = [i[0] for i in obj]
            return [normalize_scalar(x) if not isinstance(x, (list, dict, tuple)) else normalize(x) for x in flat]
        # flatten single-row 2D like [[a,b,c]] -> [a,b,c]
        if len(obj) == 1 and isinstance(obj[0], list):
            return normalize(obj[0])
        # otherwise normalize each element
        return [normalize(x) for x in obj]

    # string: try to literal_eval (could yield dict/list/tuple/bool/num)
    if isinstance(obj, str):
        le = try_literal_eval(obj)
        # If literal_eval returned something complex, normalize it
        if not isinstance(le, str):
            return normalize(le)
        # else try scalar normalization (numbers / true/false)
        return normalize_scalar(obj)

    # scalar types (bool, int, float, None)
    return normalize_scalar(obj)

# ------------ main parser ------------


def parse_kwargs(kwargs_input) -> dict:
    """
    Parse and normalize kwargs passed from Excel UDFs.

    Accepts:
      - dict -> normalizes and returns it
      - string (Python-literal dict) -> ast.literal_eval + normalize
      - list/tuple etc (unlikely for top-level) -> if dict returned, convert; else {}
    Returns:
      - dict safe to use as **kwargs (empty dict on failure)
    """
    if kwargs_input is None:
        return {}

    # If the caller already passed a dict (e.g., via xl_dict), normalize it
    if isinstance(kwargs_input, dict):
        return normalize(kwargs_input)

    # If it's bytes or other, try converting to str first
    if not isinstance(kwargs_input, (str, list, tuple, dict)):
        # fallback: try literal_eval on str(kwargs_input)
        try:
            parsed = try_literal_eval(str(kwargs_input))
            if isinstance(parsed, dict):
                return normalize(parsed)
            return {}
        except Exception:
            return {}

    # If it's a string, try parsing it as a Python-literal dict first
    if isinstance(kwargs_input, str):
        parsed = try_literal_eval(kwargs_input)
        if isinstance(parsed, dict):
            return normalize(parsed)
        # if literal_eval returned a different type, attempt to normalize that, but we need a dict
        if isinstance(parsed, (list, tuple)):
            # maybe user passed a list of pairs? convert pairs to dict if possible
            try:
                return {k: normalize(v) for k, v in parsed}
            except Exception:
                return {}
        # else cannot convert to dict
        return {}

    # If it's a list/tuple/dict object from Excel (xl_* helpers), normalize
    if isinstance(kwargs_input, (list, tuple)):
        # try to interpret as list of pairs: ["k1","v1","k2","v2"] or [[k1,v1],[k2,v2]]
        # flatten first
        # case: [[k1,v1],[k2,v2]]
        if all(isinstance(x, (list, tuple)) and len(x) == 2 for x in kwargs_input):
            try:
                return {str(x[0]): normalize(x[1]) for x in kwargs_input}
            except Exception:
                pass
        # case: flat list ["k1","v1","k2","v2"]
        if len(kwargs_input) % 2 == 0:
            try:
                items = list(kwargs_input)
                d = {}
                for i in range(0, len(items), 2):
                    d[str(items[i])] = normalize(items[i+1])
                return d
            except Exception:
                return {}
        return {}

    # fallback
    return {}


def df_wrapper(df: pd.DataFrame, method: str, params: dict):
    try:
        if not hasattr(df, method):
            return f"DF error: Unsupported method '{method}'"
        func = getattr(df, method)
        sig = inspect.signature(func)

        valid_args = {k: v for k, v in params.items() if k in sig.parameters}
        result = func(**valid_args)

        return result if isinstance(result, (pd.DataFrame, pd.Series)) \
            else pd.DataFrame([[result]], columns=[method])
    except Exception as e:
        return f"{method} error: {e}"
