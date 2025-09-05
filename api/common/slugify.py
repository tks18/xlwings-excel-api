# slugify_udfs.py
import xlwings as xw
from slugify import slugify

# Default separator
DEFAULT_SEPARATOR = "_"

# -------------------------
# Basic Slugify UDFs
# -------------------------


@xw.func
@xw.arg('text', doc="Text to convert into a slug")
def SLUG_BASIC(text):
    """Convert text to a simple URL-friendly slug"""
    return slugify(text, separator=DEFAULT_SEPARATOR)


@xw.func
@xw.arg('text', doc="Text to convert")
@xw.arg('separator', doc="Separator to use in slug (default is _)")
def SLUG_SEPARATOR(text, separator=DEFAULT_SEPARATOR):
    """Convert text to slug with a custom separator"""
    return slugify(text, separator=separator)


@xw.func
@xw.arg('text', doc="Text to convert")
@xw.arg('lowercase', doc="True/False to convert slug to lowercase")
def SLUG_CASE(text, lowercase=True):
    """Convert text to slug with optional lowercase"""
    return slugify(text, lowercase=lowercase, separator=DEFAULT_SEPARATOR)


@xw.func
@xw.arg('text', doc="Text to convert")
def SLUG_CLEAN(text):
    """Convert text to slug keeping only letters, numbers, and separator"""
    return slugify(text, regex_pattern=r'[^a-zA-Z0-9\s_]', lowercase=True, separator=DEFAULT_SEPARATOR)


# -------------------------
# Advanced Slugify UDFs
# -------------------------

@xw.func
@xw.arg('text', doc="Text to convert")
@xw.arg('max_length', doc="Maximum length of the slug")
def SLUG_TRUNCATE(text, max_length=50):
    """Convert text to slug and truncate to max_length"""
    return slugify(text, separator=DEFAULT_SEPARATOR)[:int(max_length)]


@xw.func
@xw.arg('text', doc="Text to convert")
@xw.arg('words_to_remove', doc="Comma-separated words to remove before slugifying")
def SLUG_REMOVE_WORDS(text, words_to_remove=""):
    """Remove specified words before slugifying"""
    words = [w.strip() for w in words_to_remove.split(",") if w.strip()]
    for word in words:
        text = text.replace(word, "")
    return slugify(text, separator=DEFAULT_SEPARATOR)


@xw.func
@xw.arg('text', doc="Text to convert")
@xw.arg('prefix', doc="Text to prepend to slug")
@xw.arg('suffix', doc="Text to append to slug")
def SLUG_PREFIX_SUFFIX(text, prefix="", suffix=""):
    """Add prefix and/or suffix to slug"""
    slug = slugify(text, separator=DEFAULT_SEPARATOR)
    return f"{prefix}{slug}{suffix}"


@xw.func
@xw.arg('text', doc="Text to convert")
def SLUG_UNICODE(text):
    """Slugify while keeping Unicode characters"""
    return slugify(text, allow_unicode=True, separator=DEFAULT_SEPARATOR)


@xw.func
@xw.arg('text', doc="Text to convert")
def SLUG_NO_STOPWORDS(text):
    """Slugify text while removing English stopwords"""
    stopwords = ['a', 'an', 'the', 'and', 'or',
                 'but', 'for', 'on', 'in', 'with', 'to', 'of']
    return slugify(text, separator=DEFAULT_SEPARATOR, stopwords=stopwords)


@xw.func
@xw.arg('text', doc="Text to convert")
def SLUG_ONLY_ASCII(text):
    """Slugify text and remove all non-ASCII characters"""
    return slugify(text, lowercase=True, allow_unicode=False, separator=DEFAULT_SEPARATOR)
