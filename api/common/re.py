import re
import xlwings as xw
from typing import List

# -----------------------
# RE Module UDF Functions
# -----------------------

# ---------------- Basic Functions ----------------


@xw.func
def RE_MATCH(text: str, pattern: str, flags: str = "") -> bool:
    """Returns TRUE if pattern fully matches text"""
    if text is None or pattern is None:
        return False
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    return bool(re.fullmatch(pattern, text, flags=re_flags))


@xw.func
def RE_SEARCH(text: str, pattern: str, flags: str = "") -> bool:
    """Returns TRUE if pattern found anywhere in text"""
    if text is None or pattern is None:
        return False
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    return bool(re.search(pattern, text, flags=re_flags))


@xw.func
def RE_FINDALL(text: str, pattern: str, flags: str = "") -> str:
    """Returns all matches as comma-separated string"""
    if text is None or pattern is None:
        return ""
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    matches = re.findall(pattern, text, flags=re_flags)
    return ", ".join(matches) if matches else ""


@xw.func
def RE_SPLIT(text: str, pattern: str, flags: str = "") -> str:
    """Splits text by pattern and returns comma-separated values"""
    if text is None or pattern is None:
        return text
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    parts = re.split(pattern, text, flags=re_flags)
    return ", ".join(parts)


@xw.func
def RE_SUB(text: str, pattern: str, repl: str, flags: str = "") -> str:
    """Replaces pattern with replacement in text"""
    if text is None or pattern is None:
        return text
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    return re.sub(pattern, repl, text, flags=re_flags)


@xw.func
def RE_SUBN(text: str, pattern: str, repl: str, flags: str = "") -> str:
    """Returns replaced text along with number of replacements: 'text | count'"""
    if text is None or pattern is None:
        return text
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    result, count = re.subn(pattern, repl, text, flags=re_flags)
    return f"{result} | {count}"


@xw.func
def RE_ESCAPE(text: str) -> str:
    """Escapes all regex special characters in text"""
    if text is None:
        return ""
    return re.escape(text)


@xw.func
def RE_FINDALL_MULTILINE(text: str, pattern: str, flags: str = "") -> str:
    """Finds all matches line by line in multiline text"""
    if text is None or pattern is None:
        return ""
    re_flags = re.MULTILINE
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    matches = re.findall(pattern, text, flags=re_flags)
    return ", ".join(matches) if matches else ""


# ---------------- Advanced Functions ----------------
@xw.func
def RE_EXTRACT_BEFORE(text: str, pattern: str, flags: str = "") -> str:
    """Returns everything before the first occurrence of pattern"""
    if text is None or pattern is None:
        return ""
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    match = re.search(pattern, text, flags=re_flags)
    return text[:match.start()] if match else text


@xw.func
def RE_EXTRACT_AFTER(text: str, pattern: str, flags: str = "") -> str:
    """Returns everything after the first occurrence of pattern"""
    if text is None or pattern is None:
        return ""
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    match = re.search(pattern, text, flags=re_flags)
    return text[match.end():] if match else ""


@xw.func
def RE_GROUP(text: str, pattern: str, group_number: int = 0, flags: str = "") -> str:
    """Returns a specific regex group"""
    if text is None or pattern is None:
        return ""
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    match = re.search(pattern, text, flags=re_flags)
    if match:
        try:
            return match.group(group_number)
        except IndexError:
            return ""
    return ""


@xw.func
def RE_FIND_ITER(text: str, pattern: str, max_matches: int = 10, flags: str = "") -> str:
    """Returns first N matches as comma-separated string"""
    if text is None or pattern is None:
        return ""
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    matches = [m.group() for m in re.finditer(pattern, text, flags=re_flags)]
    return ", ".join(matches[:max_matches]) if matches else ""


@xw.func
def RE_COUNT(text: str, pattern: str, flags: str = "") -> int:
    """Counts how many times pattern appears in text"""
    if text is None or pattern is None:
        return 0
    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE
    return len(re.findall(pattern, text, flags=re_flags))


@xw.func
def RE_IS_MATCH(text: str, pattern: str, flags: str = "") -> str:
    """Returns 'Yes' if pattern found, 'No' otherwise"""
    return "Yes" if RE_SEARCH(text, pattern, flags) else "No"


@xw.func
@xw.ret(expand='table')  # ensures Excel spills the output
def RE_EXTRACT_ALL_GROUPS_LIST(text: str, pattern: str, flags: str = "") -> list:
    """
    Extracts all groups for all matches and returns a 2D list.
    Each row = one match, each column = one group.
    """
    if text is None or pattern is None:
        return []

    re_flags = 0
    if 'i' in flags:
        re_flags |= re.IGNORECASE

    matches = re.findall(pattern, text, flags=re_flags)

    result = []
    for m in matches:
        if isinstance(m, tuple):
            result.append(list(m))
        else:
            result.append([m])  # single group match

    return result
