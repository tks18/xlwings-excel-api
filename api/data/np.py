import xlwings as xw
import numpy as np

# Helper to convert Excel input to a proper NumPy array


def _to_array(data, dtype=float):
    arr = np.array(data)
    if arr.ndim == 0:
        arr = np.array([arr], dtype=dtype)
    else:
        arr = arr.astype(dtype)
    return arr

# ------------------------
# Random Generators
# ------------------------


@xw.func
def NP_RANDOM():
    return np.random.random()


@xw.func
def NP_RANDOM_INT(low: int, high: int):
    return np.random.randint(low, high)


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_ARRAY(low: float, high: float, size: int):
    return np.random.uniform(low, high, size).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_INT_ARRAY(low: int, high: int, size: int):
    return np.random.randint(low, high, size).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_NORMAL(mean: float, std: float, size: int):
    return np.random.normal(mean, std, size).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_CHOICE(data, size: int, replace: bool = True):
    arr = _to_array(data)
    return np.random.choice(arr, size=size, replace=replace).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_SHUFFLE(data):
    arr = _to_array(data)
    np.random.shuffle(arr)
    return arr.tolist()


@xw.func
def NP_RANDOM_SEED(seed: int):
    np.random.seed(seed)
    return f"Random seed set to {seed}"

# ------------------------
# Probability Distributions
# ------------------------


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_BINOMIAL(n: int, p: float, size: int):
    return np.random.binomial(n, p, size).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_POISSON(lam: float, size: int):
    return np.random.poisson(lam, size).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RANDOM_EXPONENTIAL(sc: float, size: int):
    return np.random.exponential(sc, size).tolist()

# ------------------------
# Array Utilities
# ------------------------


@xw.func
def NP_MEAN(data):
    arr = _to_array(data)
    return np.mean(arr)


@xw.func
def NP_MEDIAN(data):
    arr = _to_array(data)
    return np.median(arr)


@xw.func
def NP_STD(data):
    arr = _to_array(data)
    return np.std(arr)


@xw.func
def NP_SUM(data):
    arr = _to_array(data)
    return np.sum(arr)


@xw.func
def NP_MIN(data):
    arr = _to_array(data)
    return np.min(arr)


@xw.func
def NP_MAX(data):
    arr = _to_array(data)
    return np.max(arr)


@xw.func
def NP_UNIQUE(data):
    arr = _to_array(data)
    return np.unique(arr).tolist()

# ------------------------
# Array Transformations
# ------------------------


@xw.func
@xw.ret(expand='table')
def NP_FLATTEN(data):
    arr = _to_array(data)
    return arr.flatten().tolist()


@xw.func
@xw.ret(expand='table')
def NP_SORT(data):
    arr = _to_array(data)
    return np.sort(arr).tolist()


@xw.func
@xw.ret(expand='table')
def NP_ARGSORT(data):
    arr = _to_array(data)
    return np.argsort(arr).tolist()


@xw.func
@xw.ret(expand='table')
def NP_RESIZE(data, rows: int, cols: int):
    arr = _to_array(data)
    reshaped = np.resize(arr, (rows, cols))
    return reshaped.tolist()

# ------------------------
# Logical Utilities
# ------------------------


@xw.func
@xw.ret(expand='table')
def NP_WHERE(condition_array, value_if_true=1, value_if_false=0):
    arr = _to_array(condition_array)
    return np.where(arr, value_if_true, value_if_false).tolist()


@xw.func
@xw.ret(expand='table')
def NP_ISIN(data, test_elements):
    arr = _to_array(data)
    test = _to_array(test_elements)
    return np.isin(arr, test).tolist()
