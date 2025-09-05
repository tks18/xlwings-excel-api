import xlwings as xw
import pandas as pd
import os

from helpers.pd import _auto_load, _get_cache_path,  _auto_cache, DF_REGISTRY

# ---------- PERSISTENCE APIS ----------


@xw.func
@xw.arg('df', pd.DataFrame, index=False)
def DF_LOAD(df_name: str, df):
    """Load Excel range into memory and cache"""
    _auto_cache(df_name, df)
    return f"{df_name} loaded ({df.shape[0]} rows, {df.shape[1]} cols)"


@xw.func
@xw.ret(index=False)
def DF_GET(name: str):
    """
    Retrieve a previously persisted DataFrame by name.
    Example:
        =DF_GET("Employees_30")
    """
    try:
        return _auto_load(name)
    except Exception as e:
        return f"DF_GET error: {e}"


@xw.func
def DF_EXISTS(name: str):
    """
    Return TRUE/FALSE if a persisted DF is available (memory or disk).
    """
    try:
        _ = _auto_load(name)
        return True
    except Exception:
        return False


@xw.func
def DF_LIST():
    """
    List names currently in memory (vertical list).
    """
    try:
        return [[k] for k in DF_REGISTRY.keys()]
    except Exception as e:
        return f"DF_LIST error: {e}"


@xw.func
def DF_UNLOAD(df_name: str):
    DF_REGISTRY.pop(df_name, None)
    path = _get_cache_path(df_name)
    if os.path.exists(path):
        os.remove(path)
    return f"{df_name} unloaded"
