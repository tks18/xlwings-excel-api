import xlwings as xw
from rapidfuzz import fuzz, process
import re

# -------------------------
# Helpers
# -------------------------


def _clean_text(s, lower=True, remove_punct=True, strip_spaces=True):
    """Basic text cleaner for better fuzzy matching"""
    if s is None:
        return ""
    s = str(s)
    if lower:
        s = s.lower()
    if remove_punct:
        s = re.sub(r"[^\w\s]", "", s)
    if strip_spaces:
        s = s.strip()
    return s

# -------------------------
# UDF Functions
# -------------------------


@xw.func
def FZ_RATIO(s1, s2):
    """Simple fuzzy ratio"""
    return fuzz.ratio(str(s1), str(s2)) if s1 and s2 else None


@xw.func
def FZ_PARTIAL_RATIO(s1, s2):
    """Partial ratio (substring match)"""
    return fuzz.partial_ratio(str(s1), str(s2)) if s1 and s2 else None


@xw.func
def FZ_TOKEN_SORT_RATIO(s1, s2):
    """Token sort ratio (ignores word order)"""
    return fuzz.token_sort_ratio(str(s1), str(s2)) if s1 and s2 else None

# -------------------------
# Extract functions
# -------------------------


@xw.func
@xw.arg("choices", expand="down")
def FZ_EXTRACT_ONE(query, choices):
    """Best fuzzy match from list"""
    if not query or not choices:
        return None
    match, score, _ = process.extractOne(
        str(query), [str(c) for c in choices if c])
    return match


@xw.func
@xw.arg("choices", expand="down")
def FZ_EXTRACT_SCORE(query, choices):
    """Best fuzzy match + score"""
    if not query or not choices:
        return None
    match, score, _ = process.extractOne(
        str(query), [str(c) for c in choices if c])
    return f"{match} ({score})"


@xw.func
@xw.arg("choices", expand="down")
def FZ_EXTRACT_INDEX(query, choices):
    """Return index (row offset in range) of best fuzzy match"""
    if not query or not choices:
        return None
    result = process.extractOne(str(query), [str(c) for c in choices if c])
    if result:
        match, score, idx = result
        return idx + 1  # 1-based index (Excel style)
    return None

# -------------------------
# Extended utilities
# -------------------------


@xw.func
@xw.arg("choices", expand="down")
def FZ_TOP_N(query, choices, n=3):
    """
    Return top N fuzzy matches.
    Example: =FZ_TOP_N("appl", A1:A10, 3)
    """
    if not query or not choices:
        return None
    matches = process.extract(
        str(query), [str(c) for c in choices if c], limit=int(n))
    return ", ".join([f"{m[0]} ({m[1]})" for m in matches])


@xw.func
@xw.arg("choices", expand="down")
def FZ_THRESHOLD(query, choices, threshold=80):
    """
    Return matches above threshold.
    Example: =FZ_THRESHOLD("appl", A1:A10, 85)
    """
    if not query or not choices:
        return None
    matches = process.extract(
        str(query), [str(c) for c in choices if c], limit=None)
    filtered = [f"{m[0]} ({m[1]})" for m in matches if m[1] >= int(threshold)]
    return ", ".join(filtered) if filtered else "No Match"


@xw.func
@xw.arg("choices", expand="down")
def FZ_CLEAN_EXTRACT_ONE(query, choices):
    """
    Fuzzy extract with cleaned text (lower, no punctuation).
    Better for messy names/data.
    """
    if not query or not choices:
        return None
    clean_choices = [_clean_text(c) for c in choices if c]
    query_clean = _clean_text(query)
    match, score, _ = process.extractOne(query_clean, clean_choices)
    return match


@xw.func
@xw.arg("choices", expand="down")
@xw.ret(expand="down")
def FZ_TOP_N_ARRAY(query, choices, n=3):
    """
    Return top N fuzzy matches as a vertical array.
    Example: =FZ_TOP_N_ARRAY("appl", A1:A10, 3)
    Output spills like:
        Apple   90
        Appel   85
        Applle  82
    """
    if not query or not choices:
        return [["No Match", ""]]
    matches = process.extract(
        str(query), [str(c) for c in choices if c], limit=int(n))
    return [[m[0], m[1]] for m in matches]


@xw.func
@xw.arg("choices", expand="down")
@xw.ret(expand="down")
def FZ_THRESHOLD_ARRAY(query, choices, threshold=80):
    """
    Return matches above threshold as a vertical array.
    Example: =FZ_THRESHOLD_ARRAY("appl", A1:A10, 85)
    Output spills like:
        Apple   90
        Applle  87
    """
    if not query or not choices:
        return [["No Match", ""]]
    matches = process.extract(
        str(query), [str(c) for c in choices if c], limit=None)
    filtered = [[m[0], m[1]] for m in matches if m[1] >= int(threshold)]
    return filtered if filtered else [["No Match", ""]]
