import xlwings as xw
from faker import Faker
import re

from helpers.pd import parse_kwargs

# Initialize default Faker instance
fake = Faker()

# -------------------------
# FAKER UDFs - EXTENDED
# -------------------------


@xw.func
def FAKER_NAME(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.name()


@xw.func
def FAKER_FIRST_NAME(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.first_name_male() if params.get('gender') == 'male' else \
        f.first_name_female() if params.get('gender') == 'female' else f.first_name()


@xw.func
def FAKER_LAST_NAME(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.last_name()


@xw.func
def FAKER_ADDRESS(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    addr = f.address()
    if params.get('include_postcode') is False:
        addr = re.sub(r'\d{5,}', '', addr)
    return addr


@xw.func
def FAKER_CITY(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.city()


@xw.func
def FAKER_STATE(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.state()


@xw.func
def FAKER_COUNTRY(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.country()


@xw.func
def FAKER_POSTCODE(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.postcode()


@xw.func
def FAKER_EMAIL(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.email()


@xw.func
def FAKER_PHONE_NUMBER(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    phone = f.phone_number()
    if params.get('country_code'):
        phone = f"+{params['country_code']} {phone}"
    return phone


@xw.func
def FAKER_COMPANY(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    company = f.company()
    if params.get('suffix'):
        company += f" {params['suffix']}"
    return company


@xw.func
def FAKER_JOB(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.job()


@xw.func
def FAKER_TEXT(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.text(max_nb_chars=int(params.get('max_nb_chars', 200)))


@xw.func
def FAKER_DATE(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    start = params.get('start_date') or '-30y'
    end = params.get('end_date') or 'today'
    return str(f.date_between(start_date=start, end_date=end))


@xw.func
def FAKER_UUID(kwargs_in="{}"):
    return str(fake.uuid4())


@xw.func
def FAKER_COLOR_NAME(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    f = Faker(params.get('locale')) if params.get('locale') else fake
    return f.color_name()


@xw.func
def FAKER_PASSWORD(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    length = int(params.get('length', 12))
    special_chars = bool(params.get('special_chars', True))
    digits = bool(params.get('digits', True))
    upper_case = bool(params.get('upper_case', True))
    lower_case = bool(params.get('lower_case', True))
    return fake.password(length=length, special_chars=special_chars, digits=digits,
                         upper_case=upper_case, lower_case=lower_case)


@xw.func
def FAKER_LATITUDE(kwargs_in="{}"):
    return fake.latitude()


@xw.func
def FAKER_LONGITUDE(kwargs_in="{}"):
    return fake.longitude()


@xw.func
def FAKER_LANGUAGE_CODE(kwargs_in="{}"):
    return fake.language_code()


@xw.func
def FAKER_ISBN13(kwargs_in="{}"):
    return fake.isbn13()


@xw.func
def FAKER_BANK_ACCOUNT(kwargs_in="{}"):
    return fake.bban()


@xw.func
def FAKER_IBAN(kwargs_in="{}"):
    return fake.iban()


@xw.func
def FAKER_CREDIT_CARD_NUMBER(kwargs_in="{}"):
    return fake.credit_card_number()


@xw.func
def FAKER_CREDIT_CARD_EXPIRY(kwargs_in="{}"):
    return fake.credit_card_expire()


@xw.func
def FAKER_CREDIT_CARD_PROVIDER(kwargs_in="{}"):
    return fake.credit_card_provider()


@xw.func
def FAKER_PROFILE(kwargs_in="{}"):
    params = parse_kwargs(kwargs_in)
    if params.get('fields'):
        return str(fake.profile(fields=params['fields']))
    return str(fake.profile())
