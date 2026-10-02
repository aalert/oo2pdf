import datetime as dt

import pytest

from oo2pdf.xlsx.locale import get_locale
from oo2pdf.xlsx.numfmt import format_general, format_value

ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'


@pytest.mark.parametrize("value,code,expected", [
    (1234.5, "General", "1234.5"),
    (1234567.891, "General", "1234567.891"),
    (0.1234, "0.0%", "12.3%"),
    (1234567.891, "#,##0", "1,234,568"),
    (-42.5, ACCOUNTING, " $(42.50)"),
    (0, ACCOUNTING, " $-   "),
    (3.14159, "# ?/?", "3 1/7"),
    (0.000012345, "0.00E+00", "1.23E-05"),
    (12345.678, "0.00E+00", "1.23E+04"),
    (dt.datetime(2024, 3, 15, 14, 30), "dddd, mmmm d, yyyy h:mm AM/PM", "Friday, March 15, 2024 2:30 PM"),
    (dt.date(2024, 1, 1), "yyyy-mm-dd", "2024-01-01"),
    (dt.date(2024, 1, 15), "mm-dd-yy", "1/15/2024"),  # built-in 14 uses the locale's short date
    (1.5, "[h]:mm:ss", "36:00:00"),
    (0.5, "h:mm AM/PM", "12:00 PM"),
    (-45.678, "0.00", "-45.68"),
    (-5, "0;[Red](0)", "(5)"),
    (1234, "000000", "001234"),
    (12345678, "#,##0,", "12,346"),
    (123456789, '#,##0.0,,"M"', "123.5M"),
    ("abc", '@" text"', "abc text"),
    (1.75, "# ?/4", "1 3/4"),
    (123456789012, "General", "1.23457E+11"),
    (0.1 + 0.2, "General", "0.3"),
    (1 / 3, "General", "0.333333333"),
    (True, "General", "TRUE"),
    (44927, "d-mmm-yy", "1-Jan-23"),
    (0, '0.00;-0.00;"zero"', "zero"),
    (5, '[>=10]"big";[<0]"neg";"small"', "small"),
    (60, "yyyy-mm-dd", "1900-02-29"),  # Excel's fictitious leap day
    (-0.0001, "0.00", "0.00"),  # no minus sign when every digit shown is zero
    (0.000012345, "General", "0.000012345"),
])
def test_format_en_us(value, code, expected):
    assert format_value(value, code).plain() == expected


def test_locale_separators_and_names():
    loc = get_locale("es-CL")
    assert format_value(9.99, '"$"#,##0.00', loc=loc).plain() == "$9,99"
    assert format_value(1234567.891, "#,##0", loc=loc).plain() == "1.234.568"
    assert format_value(1234567.891, "General", loc=loc).plain() == "1234567,891"
    assert format_value(dt.date(2024, 1, 15), "mm-dd-yy", loc=loc).plain() == "15-01-2024"
    assert format_value(dt.date(2024, 3, 15), "mmmm", loc=loc).plain() == "marzo"


def test_general_shrinks_to_fit():
    assert format_general(1234567.891, 7) == "1234568"
    assert format_general(123456789012, 8) == "1.23E+11"


def test_fill_and_padding_parts():
    f = format_value(-42.5, ACCOUNTING)
    kinds = [k for k, _ in f.parts]
    assert "fill" in kinds and "pad" in kinds
