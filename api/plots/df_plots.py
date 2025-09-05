import xlwings as xw
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

from helpers.pd import _auto_load, parse_kwargs, _normalize
from helpers.plot import _plot_wrapper


import xlwings as xw
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd


@xw.func
def DF_PLOT(src_name: str, kind: str, plot_name: str,  kwargs_in="{}"):
    """
    kind      : str, type of plot ("line", "bar", "box", "violin", "hist", "count", "scatter", "reg", "heatmap", "pairplot")
    src_name  : str, registered DataFrame name (loaded via _auto_load)
    kwargs_in : str, Excel-friendly kwargs dictionary, e.g. '{"x": "col1", "y": "col2", "hue": "col3"}'
    """
    try:
        # Load DataFrame
        df = _auto_load(src_name)

        # Parse kwargs from Excel string
        params = parse_kwargs(kwargs_in)

        return _plot_wrapper(kind, df, plot_name, params)

    except Exception as e:
        return f"PLOT error: {e}"


@xw.func
def SNS_PLOT(df: pd.DataFrame, kind: str, plot_name: str,  kwargs_in="{}"):
    """
    kind      : str, type of plot ("line", "bar", "box", "violin", "hist", "count", "scatter", "reg", "heatmap", "pairplot")
    src_name  : str, registered DataFrame name (loaded via _auto_load)
    kwargs_in : str, Excel-friendly kwargs dictionary, e.g. '{"x": "col1", "y": "col2", "hue": "col3"}'
    """
    try:
        # Parse kwargs from Excel string
        params = parse_kwargs(kwargs_in)

        return _plot_wrapper(kind, df, plot_name, params)

    except Exception as e:
        return f"PLOT error: {e}"
