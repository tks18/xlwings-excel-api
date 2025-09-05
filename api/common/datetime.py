import xlwings as xw
from datetime import datetime, date, timedelta
from calendar import monthrange
import xlwings.utils as xw_utils

# -------------------------------
# Helper Functions
# -------------------------------


def to_datetime(val):
    """Convert Excel serial date or string to datetime"""
    if isinstance(val, (int, float)):
        return xw_utils.xlserial_to_datetime(val)
    elif isinstance(val, str):
        return datetime.fromisoformat(val)
    elif isinstance(val, datetime):
        return val
    elif isinstance(val, date):
        return datetime.combine(val, datetime.min.time())
    else:
        raise ValueError(f"Cannot convert {val} to datetime")


def to_excel_serial(dt):
    """Convert datetime/date to Excel serial number"""
    return xw_utils.datetime_to_xlserial(dt)

# -------------------------------
# 1. Current Date & Time
# -------------------------------


@xw.func
def DT_CURRENT_DATETIME():
    """Returns current datetime (Excel serial)"""
    return to_excel_serial(datetime.now())


@xw.func
def DT_CURRENT_DATE():
    """Returns current date (Excel serial)"""
    return to_excel_serial(date.today())


@xw.func
def DT_CURRENT_TIME():
    """Returns current time (Excel serial)"""
    return to_excel_serial(datetime.now())

# -------------------------------
# 2. Date Arithmetic
# -------------------------------


@xw.func
def DT_ADD_DAYS(start_date, days):
    """Add days to a date"""
    dt = to_datetime(start_date)
    return to_excel_serial(dt + timedelta(days=int(days)))


@xw.func
def DT_ADD_WEEKS(start_date, weeks):
    return DT_ADD_DAYS(start_date, int(weeks) * 7)


@xw.func
def DT_ADD_MONTHS(start_date, months):
    dt = to_datetime(start_date)
    month = dt.month - 1 + int(months)
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, monthrange(year, month)[1])
    return to_excel_serial(date(year, month, day))


@xw.func
def DT_ADD_YEARS(start_date, years):
    return DT_ADD_MONTHS(start_date, int(years) * 12)

# -------------------------------
# 3. Date Difference & Checks
# -------------------------------


@xw.func
def DT_DAYS_BETWEEN(start_date, end_date):
    dt1 = to_datetime(start_date)
    dt2 = to_datetime(end_date)
    return (dt2 - dt1).days


@xw.func
def DT_IS_WEEKEND(some_date):
    dt = to_datetime(some_date)
    return dt.weekday() >= 5

# -------------------------------
# 4. Start / End of Period
# -------------------------------


@xw.func
def DT_START_OF_MONTH(d):
    dt = to_datetime(d)
    return to_excel_serial(date(dt.year, dt.month, 1))


@xw.func
def DT_END_OF_MONTH(d):
    dt = to_datetime(d)
    last_day = monthrange(dt.year, dt.month)[1]
    return to_excel_serial(date(dt.year, dt.month, last_day))


@xw.func
def DT_START_OF_WEEK(d):
    dt = to_datetime(d)
    return to_excel_serial((dt - timedelta(days=dt.weekday())).date())


@xw.func
def DT_END_OF_WEEK(d):
    dt = to_datetime(d)
    return to_excel_serial((dt + timedelta(days=6 - dt.weekday())).date())


@xw.func
def DT_START_OF_YEAR(d):
    dt = to_datetime(d)
    return to_excel_serial(date(dt.year, 1, 1))


@xw.func
def DT_END_OF_YEAR(d):
    dt = to_datetime(d)
    return to_excel_serial(date(dt.year, 12, 31))

# -------------------------------
# 5. Age / Leap Year
# -------------------------------


@xw.func
def DT_IS_LEAP_YEAR(year):
    y = int(year)
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


@xw.func
def DT_AGE_FROM_BIRTHDATE(birth_date):
    dt = to_datetime(birth_date)
    today = date.today()
    age = today.year - dt.year - \
        ((today.month, today.day) < (dt.month, dt.day))
    return age

# -------------------------------
# 6. Week/Quarter/Fiscal Utilities
# -------------------------------


@xw.func
def DT_WEEK_NUMBER(d):
    dt = to_datetime(d)
    return dt.isocalendar()[1]


@xw.func
def DT_QUARTER(d):
    dt = to_datetime(d)
    return ((dt.month - 1) // 3) + 1


@xw.func
def DT_IS_BUSINESS_DAY(d):
    dt = to_datetime(d)
    return dt.weekday() < 5

# -------------------------------
# 7. Human-Readable Time Difference
# -------------------------------


@xw.func
def DT_TIME_AGO(d):
    dt = to_datetime(d)
    now = datetime.now()
    diff = now - dt
    seconds = diff.total_seconds()
    if seconds < 60:
        return f"{int(seconds)} sec ago"
    elif seconds < 3600:
        return f"{int(seconds//60)} min ago"
    elif seconds < 86400:
        return f"{int(seconds//3600)} hr ago"
    else:
        return f"{int(seconds//86400)} days ago"

# -------------------------------
# 8. Excel Serial Converters
# -------------------------------


@xw.func
def DT_TO_SERIAL(d):
    dt = to_datetime(d)
    return to_excel_serial(dt)


@xw.func
def DT_FROM_SERIAL(serial):
    return xw_utils.xldate_to_datetime(serial)
