import re
import xlwings as xw


@xw.func
def RE_FINDALL(text, pattern, flags_in="0"):
    """
    Return all matches of pattern in text.

    Args:
        text (str): Input string
        pattern (str): Regex pattern
        flags_in (str/int): Regex flags as int or string
    """
    try:
        # Convert flags from string/int to re flags
        flags = int(flags_in)
        matches = re.findall(pattern, text, flags)
        return matches if matches else ["No match"]
    except Exception as e:
        return f"RE_FINDALL error: {e}"


@xw.func
def RE_SEARCH(text, pattern, flags_in="0"):
    """
    Return first match object string of pattern in text.
    """
    try:
        flags = int(flags_in)
        m = re.search(pattern, text, flags)
        return m.group(0) if m else "No match"
    except Exception as e:
        return f"RE_SEARCH error: {e}"


@xw.func
def RE_SPLIT(text, pattern, maxsplit=0, flags_in="0"):
    """
    Split text using regex pattern
    """
    try:
        flags = int(flags_in)
        return re.split(pattern, text, maxsplit=maxsplit, flags=flags)
    except Exception as e:
        return f"RE_SPLIT error: {e}"
