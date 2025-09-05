import xlwings as xw

# -------------------------
# Helper UDF's
# 1. Build a Python list/array from Excel
# -------------------------


@xw.func
def XL_ARRAY(*args):
    """
    Build a Python list from Excel values.
    Handles flattened Excel ranges.

    Example:
        =xl_array(1,2,3) -> [1,2,3]
        =xl_array(A1:A3) -> [val1,val2,val3]
    """
    res = []
    for a in args:
        # Flatten single-column Excel ranges
        if isinstance(a, list) and all(isinstance(i, list) and len(i) == 1 for i in a):
            res.extend([i[0] for i in a])
        else:
            res.append(a)
    return res

# -------------------------
# 2. Build a Python tuple from Excel
# -------------------------


@xw.func
def XL_TUPLE(*args):
    """
    Build a Python tuple from Excel values.

    Example:
        =xl_tuple(1,2,3) -> (1,2,3)
        =xl_tuple(A1:A3) -> (val1,val2,val3)
    """
    res = []
    for a in args:
        if isinstance(a, list) and all(isinstance(i, list) and len(i) == 1 for i in a):
            res.extend([i[0] for i in a])
        else:
            res.append(a)
    return tuple(res)

# -------------------------
# 3. Build a Python dict from Excel
# -------------------------


@xw.func
def XL_DICT(*args):
    """
    Build a Python dict from Excel key-value pairs.

    Example:
        =xl_dict("alpha",0.5,"s",20) -> {'alpha':0.5,'s':20}
    """
    d = {}
    for i in range(0, len(args), 2):
        if i+1 < len(args):
            key = args[i]
            val = args[i+1]
            # Flatten single-cell lists from Excel
            if isinstance(val, list) and len(val) == 1:
                val = val[0]
            d[key] = val
    return d

# -------------------------
# 4. Build kwargs dict from Excel
# -------------------------


@xw.func
def XL_PARAMS(*args):
    """
    Build a Python-literal dict string from Excel key-value arguments.

    Usage in Excel:
      =BUILD_KWARGS(
         "hue","species",
         "vars", xl_tuple("col1","col2"),
         "plot_kws", xl_dict("alpha",0.5,"s",20),
         "ascending", xl_array(TRUE,FALSE)
      )

    Returns a string like:
      "{'hue': 'species', 'vars': ('col1','col2'), 'plot_kws': {'alpha': 0.5, 's': 20}, 'ascending': [True, False]}"
    """
    # Build raw kwargs dict from positional args
    kwargs = {}
    for i in range(0, len(args), 2):
        if i + 1 >= len(args):
            break
        key = str(args[i])
        val = args[i + 1]

        # Flatten 2D single-column/single-row Excel ranges (xlwings passes ranges as nested lists)
        if isinstance(val, list) and all(isinstance(r, list) and len(r) == 1 for r in val):
            val = [r[0] for r in val]
            # if single element, keep it as list for explicit list semantics (caller choice)
        kwargs[key] = val

    # Serializer: convert Python objects -> Python-literal string pieces
    def to_literal(v):
        # dict
        if isinstance(v, dict):
            items = []
            for k, vv in v.items():
                # dict keys from xl_dict are likely strings; use repr to be safe
                items.append(f"{repr(k)}: {to_literal(vv)}")
            return "{" + ", ".join(items) + "}"
        # list
        if isinstance(v, list):
            return "[" + ", ".join(to_literal(x) for x in v) + "]"
        # tuple
        if isinstance(v, tuple):
            items = ", ".join(to_literal(x) for x in v)
            # single-element tuple must include trailing comma
            if len(v) == 1:
                return "(" + items + ",)"
            return "(" + items + ")"
        # string -> add quotes
        if isinstance(v, str):
            return repr(v)
        # None
        if v is None:
            return "None"
        # bool / int / float / other -> use repr
        return repr(v)

    # Build final python-literal dict string
    pairs = []
    for k, v in kwargs.items():
        pairs.append(f"{repr(k)}: {to_literal(v)}")
    return "{" + ", ".join(pairs) + "}"
