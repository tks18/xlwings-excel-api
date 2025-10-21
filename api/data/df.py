import pandas as pd
import xlwings as xw
import io
from typing import Union, List, Mapping, cast, Callable

from helpers.pd import parse_kwargs


@xw.func
def DF_STD_HEAD(df: pd.DataFrame, kwargs_in="{}"):
    """Return head of DF"""
    try:
        params = parse_kwargs(kwargs_in)
        return df.head(**params)
    except Exception as e:
        return f"DF_STD_HEAD error: {e}"


@xw.func
def DF_STD_TAIL(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.tail(**params)
    except Exception as e:
        return f"DF_STD_TAIL error: {e}"


@xw.func
def DF_STD_INFO(df: pd.DataFrame, as_table=True, kwargs_in=None):
    try:
        params = parse_kwargs(kwargs_in)

        # Keep only valid kwargs for DataFrame.info
        valid_keys = {"verbose", "max_cols", "memory_usage", "show_counts"}
        info_kwargs = {k: params[k] for k in params.keys() & valid_keys}

        # Capture info() output
        buf = io.StringIO()
        df.info(buf=buf, **info_kwargs)
        text = buf.getvalue()

        if as_table:
            # One line per row for cleaner display in Excel
            return [[line] for line in text.splitlines()]

        return text
    except Exception as e:
        return f"DF_STD_INFO error: {e}"


@xw.func
def DF_STD_DESCRIBE(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.describe(**params)
    except Exception as e:
        return f"DF_STD_DESCRIBE error: {e}"


@xw.func
def DF_STD_GROUPBY(df: pd.DataFrame, by, cols=None, funcs=None):
    """
    Group by columns and aggregate specific columns with given functions.

    Parameters
    ----------
    src_name : str
        Name of the source DataFrame.
    by : str | list[str]
        Column(s) to group by.
    cols : str | list[str] | None
        Column(s) to aggregate. If None -> all numeric columns.
    funcs : str | list[str] | None
        Aggregation function(s). Examples: "sum", "mean", "max", ["sum","count"]

    Returns
    -------
    DataFrame grouped and aggregated.
    """
    try:

        # Normalize grouping columns
        by_cols = [by] if isinstance(by, str) else list(x for x in by if x)

        # Normalize columns to aggregate
        if cols is None or cols == "":
            agg_cols = df.select_dtypes("number").columns.tolist()
        elif isinstance(cols, str):
            agg_cols = [cols]
        else:
            agg_cols = [c for c in cols if c]

        # Normalize aggregation functions
        if funcs is None or funcs == "":
            agg_funcs = ["sum"]
        elif isinstance(funcs, str):
            agg_funcs = [funcs]
        else:
            agg_funcs = [f for f in funcs if f]

        # Build aggregation mapping
        agg_dict = cast(Mapping[str, Union[str, Callable, List[Union[str, Callable]]]], {
                        col: agg_funcs for col in agg_cols})

        result = df.groupby(by_cols).agg(agg_dict).reset_index(drop=True)

        # Flatten MultiIndex if multiple agg funcs
        if isinstance(result.columns, pd.MultiIndex):
            result.columns = ["_".join([c for c in tup if c])
                              for tup in result.columns]

        return result

    except Exception as e:
        return f"DF_STD_GROUPBY error: {e}"


@xw.func
def DF_STD_SORT(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.sort_values(**params)
    except Exception as e:
        return f"DF_STD_SORT error: {e}"


@xw.func
def DF_STD_QUERY(df: pd.DataFrame, expr: str):
    try:
        return df.query(expr)
    except Exception as e:
        return f"DF_STD_QUERY error: {e}"


@xw.func
def DF_STD_PIVOT(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.pivot_table(**params)
    except Exception as e:
        return f"DF_STD_PIVOT error: {e}"


@xw.func
def DF_STD_DROP(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.drop(**params)
    except Exception as e:
        return f"DF_STD_DROP error: {e}"


@xw.func
def DF_STD_FILLNA(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.fillna(**params)
    except Exception as e:
        return f"DF_STD_FILLNA error: {e}"


@xw.func
def DF_STD_RENAME(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.rename(**params)
    except Exception as e:
        return f"DF_STD_RENAME error: {e}"


@xw.func
def DF_STD_ASSIGN(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.assign(**params)
    except Exception as e:
        return f"DF_STD_ASSIGN error: {e}"


@xw.func
@xw.ret(index=False)
def DF_STD_RESET_INDEX(df: pd.DataFrame, kwargs_in="{}"):
    try:
        params = parse_kwargs(kwargs_in)
        return df.reset_index(**params)
    except Exception as e:
        return f"DF_STD_ASSIGN error: {e}"


@xw.func
def DF_STD_VALUE_COUNTS(df: pd.DataFrame, kwargs_in="{}"):
    """
    Return counts of unique values in a column.
    Example: =DF_STD_VALUE_COUNTS("sales_df","Region")
    """
    try:
        params = parse_kwargs(kwargs_in)
        return df.value_counts(**params).reset_index(drop=True)
    except Exception as e:
        return f"DF_STD_VALUE_COUNTS error: {e}"


@xw.func
def DF_STD_STATS(df: pd.DataFrame, mode: str, kwargs_in="{}"):
    """Compute covariance matrix."""
    try:

        if not hasattr(df, mode):
            return f"DF_STD_STATS error: Unsupported type '{mode}'"

        func = getattr(df, mode)

        params = parse_kwargs(kwargs_in)
        with_index = False
        if "with_index" in params:
            with_index = params.pop("with_index")

        result = func(**params)

        if with_index:
            return result.reset_index()
        else:
            return result
    except Exception as e:
        return f"DF_STATS error: {e}"


@xw.func
def DF_STD_TO_DATETIME(df: pd.DataFrame, cols, fmt=None):
    """
    Convert specific columns in a DataFrame to datetime.
    - cols: string or list of column names
    - fmt: optional strftime/strptime format (e.g. "%Y-%m-%d")
    - If fmt is None, will try Excel serial dates first, then pandas auto parse
    """
    try:

        # normalize columns
        col_list = [cols] if isinstance(
            cols, str) else list(c for c in cols if c)

        for col in col_list:
            if col not in df.columns:
                continue

            if fmt:  # user provided a format
                df[col] = pd.to_datetime(df[col], format=fmt, errors="coerce")
            else:
                # Try Excel serial conversion first
                if pd.api.types.is_numeric_dtype(df[col]):
                    df[col] = pd.to_datetime(
                        df[col], origin="1899-12-30", unit="D", errors="coerce")
                else:
                    df[col] = pd.to_datetime(df[col], errors="coerce")

        return df

    except Exception as e:
        return f"DF_STD_TO_DATETIME error: {e}"
