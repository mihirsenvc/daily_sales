"""Parse report dates from Daily Sales Register file names."""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from typing import Optional

_MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

_EXT = re.compile(r"\.(xlsx|xls|xlsm|gsheet|csv)$", re.I)


def parse_date_from_name(raw_name: str) -> Optional[str]:
    """Return YYYY-MM-DD parsed from names like 'Daily Sales Register on 18-Aug-2026.xlsx'."""
    name = _EXT.sub("", str(raw_name or ""))

    iso = re.search(r"\b(20\d{2})[-_./](\d{1,2})[-_./](\d{1,2})\b", name)
    if iso:
        return _ymd(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))

    compact = re.search(r"\b(20\d{2})(\d{2})(\d{2})\b", name)
    if compact:
        return _ymd(int(compact.group(1)), int(compact.group(2)), int(compact.group(3)))

    named = re.search(
        r"\b(\d{1,2})[-_ ]([A-Za-z]{3,9})[-_ ](20\d{2})\b",
        name,
    )
    if named:
        month = _MONTHS.get(named.group(2).lower())
        if month:
            return _ymd(int(named.group(3)), month, int(named.group(1)))

    named2 = re.search(
        r"\b([A-Za-z]{3,9})[-_ ](\d{1,2})[,_-]?[ ]?(20\d{2})\b",
        name,
    )
    if named2:
        month = _MONTHS.get(named2.group(1).lower())
        if month:
            return _ymd(int(named2.group(3)), month, int(named2.group(2)))

    dmy = re.search(r"\b(\d{1,2})[-_./](\d{1,2})[-_./](20\d{2})\b", name)
    if dmy:
        a, b, y = int(dmy.group(1)), int(dmy.group(2)), int(dmy.group(3))
        if a > 12 >= b:
            return _ymd(y, b, a)
        if b > 12 >= a:
            return _ymd(y, a, b)
        return _ymd(y, b, a)

    return None


def _ymd(year: int, month: int, day: int) -> Optional[str]:
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return None


def format_display(date_key: str) -> str:
    try:
        return datetime.strptime(date_key, "%Y-%m-%d").strftime("%d %b %Y")
    except ValueError:
        return date_key


def _year(raw: str) -> Optional[int]:
    try:
        year = int(raw)
    except (TypeError, ValueError):
        return None
    if year < 100:
        return 2000 + year if year < 70 else 1900 + year
    return year


def normalize_cell_text(value) -> str:
    """Collapse whitespace and case-fold for brand/team/currency matching."""
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).casefold()


def normalize_cell_date(value) -> str:
    """Turn Invoice Date / Service Period cell values into YYYY-MM-DD, or ''."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).replace("\xa0", " ").strip()
    if not text or text.lower() in ("none", "nat", "null", "nan"):
        return ""

    iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if iso:
        return _ymd(int(iso.group(1)), int(iso.group(2)), int(iso.group(3))) or ""

    ym = re.match(r"^(\d{4})-(\d{1,2})$", text)
    if ym:
        return _ymd(int(ym.group(1)), int(ym.group(2)), 1) or ""

    named = re.match(r"^(\d{1,2})[-/ ]([A-Za-z]{3,9})[-/ ](\d{2,4})$", text)
    if named:
        month = _MONTHS.get(named.group(2).lower())
        year = _year(named.group(3))
        if month and year:
            return _ymd(year, month, int(named.group(1))) or ""

    period = re.match(r"^([A-Za-z]{3,9})[-/ ]+(\d{2,4})$", text)
    if period:
        month = _MONTHS.get(period.group(1).lower())
        year = _year(period.group(2))
        if month and year:
            return _ymd(year, month, 1) or ""

    period2 = re.match(r"^(\d{1,2})[-/](\d{4})$", text)
    if period2:
        month, year = int(period2.group(1)), int(period2.group(2))
        if 1 <= month <= 12:
            return _ymd(year, month, 1) or ""

    dmy = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$", text)
    if dmy:
        a, b, year = int(dmy.group(1)), int(dmy.group(2)), _year(dmy.group(3))
        if year:
            if a > 12 >= b:
                return _ymd(year, b, a) or ""
            if b > 12 >= a:
                return _ymd(year, a, b) or ""
            return _ymd(year, b, a) or ""

    try:
        num = float(text)
    except ValueError:
        num = None
    if num is not None and 20000 <= num <= 80000:
        try:
            return (date(1899, 12, 30) + timedelta(days=int(num))).isoformat()
        except OverflowError:
            return ""
    return ""
