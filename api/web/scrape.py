import re
import requests
from bs4 import BeautifulSoup
import xlwings as xw

from helpers.pd import parse_kwargs
from helpers.web import cache_html, load_html, extract_text_xpath, extract_list_xpath

# Optional Selenium for JS pages
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False


# -------------------------
# 1️⃣ Fetch page with caching
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_FETCH(source_name: str, url: str, kwargs_in=None):
    """
    Fetch HTML content with optional caching.
    """
    kwargs = parse_kwargs(kwargs_in)
    try:
        headers = kwargs.get('headers', {'User-Agent': 'Mozilla/5.0'})
        timeout = kwargs.get('timeout', 10)
        r = requests.get(url, headers=headers, timeout=timeout)
        html = r.text
        cache_html(source_name, html)
        return f"{source_name} cached ({len(html)} chars)"
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 2️⃣ Fetch JS page with Selenium + wait
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_FETCH_JS(source_name: str,  url: str, kwargs_in=None):
    """
    Fetch JS-rendered page using Selenium with optional wait.
    """
    if not SELENIUM_AVAILABLE:
        return "Error: Selenium not installed"
    kwargs = parse_kwargs(kwargs_in)
    try:
        driver_path = kwargs.get('driver_path', 'chromedriver')
        wait_selector = kwargs.get('wait_selector', None)
        wait_time = kwargs.get('wait_time', 10)

        options = Options()
        options.add_argument("--headless")
        driver = webdriver.Chrome(
            service=Service(driver_path), options=options)
        driver.get(url)

        if wait_selector:
            WebDriverWait(driver, wait_time).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, wait_selector))
            )

        html = driver.page_source
        driver.quit()
        cache_html(source_name, html)
        return f"{source_name} cached ({len(html)} chars)"
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 3️⃣ Extract text with regex
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_REGEX_EXTRACT(source_name: str, pattern, kwargs_in=None):
    """
    Extract all matches for a regex pattern.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    flags = kwargs.get('flags', 0)
    try:
        return re.findall(pattern, html_content, flags)
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 4️⃣ Extract table directly to Excel
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_TABLE_TO_SHEET(source_name: str, selector, sheet_name=None, kwargs_in=None):
    """
    Extract HTML table and write directly to Excel sheet.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.select_one(selector)
        if not table:
            return "No table found"

        data = []
        for tr in table.find_all('tr'):
            row = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if row:
                data.append(row)

        wb = xw.Book.caller()
        sht = wb.sheets[sheet_name] if sheet_name else wb.sheets.active
        sht.range("A1").value = data
        return f"Table written to sheet '{sht.name}'"
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 3️⃣ Extract text from selector
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_TEXT(source_name: str, selector, kwargs_in=None):
    """
    Extract first matching text from HTML using CSS selector.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        element = soup.select_one(selector)
        return element.get_text(strip=True) if element else ""
    except Exception as e:
        return f"Error: {e}"


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_XPATH(source_name, xpath_expr, kwargs_in=None):
    """
    Extract text using XPath from cached HTML (source_name).
    """
    html_content = load_html(source_name)
    if not html_content:
        return "Not found"
    try:
        return extract_text_xpath(html_content, xpath_expr)
    except Exception as e:
        return f"Error: {e}"


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_XPATH_LIST(source_name, xpath_expr, kwargs_in=None):
    """
    Extract multiple items using XPath from cached HTML (source_name).
    Returns an Excel array.
    """
    html_content = load_html(source_name)
    if not html_content:
        return "Not found"
    try:
        return extract_list_xpath(html_content, xpath_expr)
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 4️⃣ Extract list of texts
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_LIST(source_name: str, selector, kwargs_in=None):
    """
    Extract multiple text items from HTML using CSS selector.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        elements = soup.select(selector)
        return [el.get_text(strip=True) for el in elements]
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 5️⃣ Extract attribute
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_ATTR(source_name: str, selector, attr_name, kwargs_in=None):
    """
    Extract a specific attribute (href, src, etc.) from HTML.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        element = soup.select_one(selector)
        return element[attr_name] if element and element.has_attr(attr_name) else ""
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 6️⃣ Extract table to 2D array
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXTRACT_TABLE(source_name: str, selector, kwargs_in=None):
    """
    Extract HTML table as Excel-friendly 2D array.
    """
    html_content = load_html(source_name)
    kwargs = parse_kwargs(kwargs_in)
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.select_one(selector)
        if not table:
            return []
        rows = []
        for tr in table.find_all('tr'):
            cells = [td.get_text(strip=True)
                     for td in tr.find_all(['td', 'th'])]
            if cells:
                rows.append(cells)
        return rows
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 8️⃣ Filter links by keyword
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_FILTER_LINKS(source_name: str, keyword, kwargs_in=None):
    """
    Return links that contain a keyword.
    """
    try:
        html_content = load_html(source_name)
        links = WEB_EXTRACT_LIST(html_content, 'a', kwargs_in)
        return [l for l in links if keyword.lower() in l.lower()]
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 9️⃣ Count elements matching selector
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_COUNT(source_name: str, selector, kwargs_in=None):
    """
    Count number of elements matching a CSS selector.
    """
    try:
        html_content = load_html(source_name)
        soup = BeautifulSoup(html_content, 'html.parser')
        return len(soup.select(selector))
    except Exception as e:
        return f"Error: {e}"


# -------------------------
# 🔟 Check if element exists
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_EXISTS(source_name: str, selector, kwargs_in=None):
    """
    Check if element exists in HTML.
    """
    try:
        html_content = load_html(source_name)
        soup = BeautifulSoup(html_content, 'html.parser')
        return bool(soup.select_one(selector))
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 1️⃣1️⃣ Extract meta content
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_META_CONTENT(source_name: str, meta_name, kwargs_in=None):
    """
    Extract <meta> tag content by name attribute.
    """
    try:
        html_content = load_html(source_name)

        soup = BeautifulSoup(html_content, 'html.parser')
        tag = soup.find('meta', attrs={'name': meta_name})
        return tag['content'] if tag and tag.has_attr('content') else ""
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 1️⃣2️⃣ Clean HTML to plain text
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_CLEAN_TEXT(source_name: str, kwargs_in=None):
    """
    Remove all HTML tags and return plain text.
    """
    try:
        html_content = load_html(source_name)
        soup = BeautifulSoup(html_content, 'html.parser')
        return soup.get_text(separator=" ", strip=True)
    except Exception as e:
        return f"Error: {e}"

# -------------------------
# 1️⃣3️⃣ Get element attribute list
# -------------------------


@xw.func
@xw.arg('kwargs_in', doc='Optional kwargs as dict')
def WEB_ATTR_LIST(source_name: str, selector, attr_name, kwargs_in=None):
    """
    Extract attribute values for all matching elements.
    """
    try:
        html_content = load_html(source_name)
        soup = BeautifulSoup(html_content, 'html.parser')
        return [el[attr_name] for el in soup.select(selector) if el.has_attr(attr_name)]
    except Exception as e:
        return f"Error: {e}"
