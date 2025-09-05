import inspect
import xlwings as xw
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# -------------------------
# Seaborn plotting UDFs (stateless)
# -------------------------


def insert_figure(fig, name="Figure"):
    try:
        caller = xw.Book.caller()
        sht = caller.sheets.active
        sht.pictures.add(fig, name=name, update=True,
                         left=300, top=50, scale=1)
    except Exception as e:
        raise RuntimeError(f"Insert figure error: {e}")

# helper wrapper for all plots


def plot_wrapper(kind: str, df: pd.DataFrame, plot_name: str, params: dict):
    try:
        # Get seaborn function dynamically
        if not hasattr(sns, kind):
            return f"PLOT error: Unsupported plot type '{kind}'"
        func = getattr(sns, kind)

        # Introspect function signature
        sig = inspect.signature(func)
        params_copy = params.copy()

        # If function supports "data" argument → pass df
        if "data" in sig.parameters:
            result = func(data=df, **params_copy)
        else:
            result = func(df, **params_copy)

        # Extract figure
        if hasattr(result, "figure"):   # FacetGrid, JointGrid, etc.
            fig = result.figure
        else:                           # Axes or direct plotting
            fig = plt.gcf()

        # Insert back to Excel
        insert_figure(fig, name=plot_name)
        plt.close(fig)
        return f"{kind.capitalize()} done"

    except Exception as e:
        return f"{plot_name} error: {e}"
