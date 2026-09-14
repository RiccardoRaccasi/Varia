#!/usr/bin/env python3
"""Track a Swiss Fund Data fund's Net Asset Value in an Excel workbook.

The module is self-contained: it depends only on the Python standard library
plus ``openpyxl`` (for the .xlsx file). It works both as a library and as a
command line tool.

Library use::

    import sys; sys.path.insert(0, "/path/to/Varia/nav-tracker")
    import nav_tracker

    point  = nav_tracker.fetch_current_price()      # scrape, no side effects
    result = nav_tracker.update()                   # scrape + append to xlsx
    rows   = nav_tracker.read_history()             # the full series

Command line use::

    python nav_tracker.py import-csv history.csv    # build the workbook
    python nav_tracker.py probe                     # show what the page offers
    python nav_tracker.py fetch                     # scrape only
    python nav_tracker.py update                    # scrape + append
    python nav_tracker.py show -n 10                # tail the series

Exit codes: 0 success, 1 usage/IO error, 2 fetch or parse failure,
3 sanity check rejected the scraped value.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import gzip
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
import zlib
from dataclasses import dataclass, field as _dc_field
from html.parser import HTMLParser
from typing import Iterable, Optional, Sequence

__version__ = "1.0.0"

__all__ = [
    "FUND_URL", "FUND_NAME", "FUND_ISIN", "FUND_CURRENCY", "DEFAULT_WORKBOOK",
    "PricePoint", "HistoryRow", "UpdateResult", "Candidate",
    "NavTrackerError", "FetchError", "ParseError", "SanityError",
    "fetch_current_price", "update", "read_history", "write_history",
    "import_csv", "extract_candidates", "parse_number", "parse_date",
]

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

FUND_URL = "https://www.swissfunddata.ch/sfdpub/en/funds/show/175908"
FUND_NAME = "Stableton Morningstar PitchBook Unicorn 20 AMC All Investors"
FUND_ISIN = "CH1234846777"
FUND_CURRENCY = "USD"

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WORKBOOK = os.path.join(_HERE, "data", "unicorn20_nav_history.xlsx")

SHEET_NAME = "NAV History"
META_SHEET = "Fund"
HEADERS = ["Date", "Net Asset Value", "Currency", "Source", "Retrieved (UTC)"]

#: Reject a scraped value that moves more than this many percent away from the
#: most recent stored value. Guards against a mis-parsed number silently
#: corrupting the series. Override with ``--max-move`` / ``max_move_pct=``.
MAX_MOVE_PCT = 30.0

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36 nav_tracker/%s" % __version__
)
TIMEOUT = 30
RETRIES = 3

LOCAL_TZ = "Europe/Zurich"


class NavTrackerError(Exception):
    """Base class for every error this module raises deliberately."""


class FetchError(NavTrackerError):
    """The page could not be retrieved."""


class ParseError(NavTrackerError):
    """The page was retrieved but no price could be read from it."""


class SanityError(NavTrackerError):
    """A price was read but it failed the plausibility checks."""


# ---------------------------------------------------------------------------
# Value objects
# ---------------------------------------------------------------------------

@dataclass
class PricePoint:
    """One observation of the fund price."""

    date: dt.date
    value: float
    currency: str = FUND_CURRENCY
    label: str = ""
    source: str = "swissfunddata"
    retrieved_at: Optional[dt.datetime] = None
    date_is_assumed: bool = False

    def as_dict(self) -> dict:
        return {
            "date": self.date.isoformat(),
            "value": self.value,
            "currency": self.currency,
            "label": self.label,
            "source": self.source,
            "retrieved_at": (self.retrieved_at or _utcnow()).isoformat(timespec="seconds"),
            "date_is_assumed": self.date_is_assumed,
        }


@dataclass
class HistoryRow:
    """One row of the workbook."""

    date: dt.date
    value: float
    currency: str = FUND_CURRENCY
    source: str = ""
    retrieved_at: str = ""


@dataclass
class UpdateResult:
    """Outcome of :func:`update`."""

    status: str           # added | updated | unchanged | skipped-stale
    point: Optional[PricePoint]
    workbook: str
    rows: int
    previous: Optional[HistoryRow] = None
    message: str = ""

    @property
    def changed(self) -> bool:
        return self.status in ("added", "updated")

    def as_dict(self) -> dict:
        return {
            "status": self.status,
            "changed": self.changed,
            "workbook": self.workbook,
            "rows": self.rows,
            "message": self.message,
            "point": self.point.as_dict() if self.point else None,
            "previous": (
                {"date": self.previous.date.isoformat(), "value": self.previous.value}
                if self.previous else None
            ),
        }


@dataclass
class Candidate:
    """A label/number pair found on the page, with where it came from."""

    field: str
    label: str
    value: float
    date: Optional[dt.date]
    strategy: str
    context: str = ""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def _utcnow() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)


def local_today() -> dt.date:
    """Today in the fund's home timezone, falling back to system local time."""
    try:
        from zoneinfo import ZoneInfo
        return dt.datetime.now(ZoneInfo(LOCAL_TZ)).date()
    except Exception:
        return dt.date.today()


def _squash(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


# No-break and thin spaces separate thousands or pad a currency code; treat
# them as ordinary spaces. Zero-width spaces carry no meaning at all.
_INVISIBLE = {0x00a0: " ", 0x202f: " ", 0x2009: " ", 0x200b: None}

#: A plain space counts as a thousands separator only inside a well-formed
#: group run, so "2026 231.845" is never glued into a single number.
_SPACED_GROUPS = re.compile(r"(?<!\d)\d{1,3}(?: \d{3})+(?:[.,]\d+)?(?!\d)")

#: A number must not begin inside a word: CH1234846777 is an ISIN.
_NUM_TOKEN = re.compile(r"(?<![A-Za-z0-9])(?<!\d[.,])(?:[-+]?\d[\d’'`.,]*\d|[-+]?\d)")


def parse_number(text: Optional[str]) -> Optional[float]:
    """Parse a number written in Swiss, German, French or English notation.

    Handles ``1'234.56``, ``1,234.56``, ``1.234,56``, ``1 234,56`` and plain
    ``231.856``. Returns ``None`` when the text holds no usable number (dates
    such as ``09.09.2026`` are rejected).
    """
    if text is None:
        return None
    cleaned = str(text).translate(_INVISIBLE)
    cleaned = _SPACED_GROUPS.sub(lambda hit: hit.group(0).replace(" ", ""), cleaned)
    match = _NUM_TOKEN.search(cleaned)
    if not match:
        return None
    tok = match.group(0)
    for sep in ("’", "'", "`"):          # Swiss thousands separators
        tok = tok.replace(sep, "")
    if "," in tok and "." in tok:
        # Whichever comes last is the decimal separator.
        cut = max(tok.rfind(","), tok.rfind("."))
        tok = re.sub(r"[.,]", "", tok[:cut]) + "." + tok[cut + 1:]
    elif "," in tok:
        parts = tok.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            tok = parts[0] + "." + parts[1]   # 1,5 -> 1.5
        else:
            tok = "".join(parts)              # 1,234 / 1,234,567 -> thousands
    try:
        return float(tok)
    except ValueError:
        return None


_DATE_TOKEN = re.compile(r"\b(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b")
_DATE_FORMATS = ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d",
                 "%d.%m.%y", "%d/%m/%y")


def parse_date(text: Optional[str]) -> Optional[dt.date]:
    """Pull a day-first or ISO date out of ``text``. Ambiguous US order is
    not attempted — Swiss Fund Data renders dates day-first."""
    if text is None:
        return None
    match = _DATE_TOKEN.search(str(text).translate(_INVISIBLE))
    if not match:
        return None
    raw = match.group(1)
    for fmt in _DATE_FORMATS:
        try:
            parsed = dt.datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
        if 1990 <= parsed.year <= local_today().year + 1:
            return parsed
    return None


# ---------------------------------------------------------------------------
# What to look for on the page
# ---------------------------------------------------------------------------

FIELD_PATTERNS = {
    "nav": [
        r"net\s*asset\s*value", r"\bnav\b", r"net\s*inventory\s*value",
        r"nettoinventarwert", r"inventarwert",
        r"valeur\s+nette\s+d[’']?inventaire",
        r"valore\s+patrimoniale\s+netto",
    ],
    "price": [
        r"current\s+price", r"latest\s+price", r"last\s+price", r"closing\s+price",
        r"chart\s+price", r"\bprice\b", r"aktueller\s+kurs", r"\bkurs\b", r"\bpreis\b",
        r"\bcours\b", r"\bprezzo\b",
    ],
    "issue": [r"issue\s+price", r"ausgabepreis", r"prix\s+d[’']?[ée]mission"],
    "redemption": [r"redemption\s+price", r"r[uü]cknahmepreis", r"prix\s+de\s+rachat"],
}

DEFAULT_FIELD_ORDER = ("nav", "price")

DATE_LABEL = re.compile(
    r"\b(date|datum|as\s+of|valuation\s+date|nav\s+date|price\s+date|stichtag|"
    r"per\b|data)\b", re.I)

# Column headers that hold something other than a price and must never be read
# as one (unit counts, fund size, performance figures).
_NOT_A_PRICE = re.compile(
    r"(volume|units?\s+outstanding|shares?\s+outstanding|fund\s+size|net\s+assets\s+"
    r"under|performance|%|ytd|ter\b|fee|isin|valor|security\s*(no|number)|"
    r"wkn|sedol|cusip|ticker)", re.I)


def _compile(patterns: Iterable[str]) -> list[re.Pattern]:
    return [re.compile(p, re.I) for p in patterns]


# ---------------------------------------------------------------------------
# HTML -> tables and label/value pairs
# ---------------------------------------------------------------------------

class _DocParser(HTMLParser):
    """Collect tables (as lists of rows of cell text) and <dt>/<dd> pairs."""

    _SKIP = {"script", "style", "noscript", "template"}
    _CELL = {"td", "th"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self.pairs: list[tuple[str, str]] = []
        self.text_parts: list[str] = []
        self._skip = 0
        self._tables: list[list[list[str]]] = []
        self._rows: list[list[str]] = []
        self._cell: Optional[list[str]] = None
        self._pending_dt: Optional[str] = None

    # -- structure ---------------------------------------------------------
    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP:
            self._skip += 1
            return
        if self._skip:
            return
        self.text_parts.append(" ")
        if tag == "table":
            table: list[list[str]] = []
            self.tables.append(table)
            self._tables.append(table)
        elif tag == "tr":
            row: list[str] = []
            if self._tables:
                self._tables[-1].append(row)
            self._rows.append(row)
        elif tag in self._CELL or tag in ("dt", "dd"):
            self._close_cell(tag)
            self._cell = []
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_startendtag(self, tag, attrs):
        if tag == "br" and self._cell is not None and not self._skip:
            self._cell.append(" ")

    def handle_endtag(self, tag):
        if tag in self._SKIP:
            self._skip = max(0, self._skip - 1)
            return
        if self._skip:
            return
        self.text_parts.append(" ")
        if tag in self._CELL or tag in ("dt", "dd"):
            self._close_cell(tag)
        elif tag == "tr":
            self._close_cell(None)
            if self._rows:
                self._rows.pop()
        elif tag == "table":
            self._close_cell(None)
            self._rows.clear()
            if self._tables:
                self._tables.pop()

    def _close_cell(self, tag: Optional[str]) -> None:
        if self._cell is None:
            return
        text = _squash("".join(self._cell))
        self._cell = None
        if tag in ("dt",):
            self._pending_dt = text
        elif tag in ("dd",):
            self.pairs.append((self._pending_dt or "", text))
            self._pending_dt = None
        elif self._rows:
            self._rows[-1].append(text)

    # -- text --------------------------------------------------------------
    def handle_data(self, data):
        if self._skip:
            return
        if self._cell is not None:
            self._cell.append(data)
        self.text_parts.append(data)

    @property
    def text(self) -> str:
        return _squash("".join(self.text_parts))


def _parse_document(html: str) -> _DocParser:
    parser = _DocParser()
    try:
        parser.feed(html)
        parser.close()
    except Exception:
        # Malformed markup: keep whatever was collected before the failure.
        pass
    return parser


# ---------------------------------------------------------------------------
# Extraction strategies
# ---------------------------------------------------------------------------

def _matches(text: str, matchers: Sequence[re.Pattern]) -> bool:
    return any(m.search(text) for m in matchers)


def _row_date(cells: Sequence[str], skip: int = -1) -> Optional[dt.date]:
    for i, cell in enumerate(cells):
        if i == skip:
            continue
        found = parse_date(cell)
        if found:
            return found
    return None


def _column_candidates(table, matchers, field_name) -> list[Candidate]:
    """Price tables: the label sits in a header row, values in the rows below."""
    out: list[Candidate] = []
    for r, row in enumerate(table):
        for c, cell in enumerate(row):
            if not cell or _NOT_A_PRICE.search(cell) or not _matches(cell, matchers):
                continue
            # A label immediately followed by a number is a label/value row,
            # which the row strategy handles; skip it here.
            if c + 1 < len(row) and parse_number(row[c + 1]) is not None:
                continue
            date_cols = [i for i, head in enumerate(row) if DATE_LABEL.search(head)]
            best: Optional[tuple] = None
            for below in table[r + 1:]:
                if c >= len(below):
                    continue
                value = parse_number(below[c])
                if value is None:
                    continue
                when = None
                for i in date_cols:
                    if i < len(below):
                        when = parse_date(below[i])
                        if when:
                            break
                if when is None:
                    when = _row_date(below, skip=c)
                key = when or dt.date.min
                if best is None or key > best[0]:
                    best = (key, value, when, below)
            if best is not None:
                out.append(Candidate(field_name, cell, best[1], best[2],
                                     "table-column", " | ".join(best[3])[:200]))
    return out


def _row_candidates(rows, matchers, field_name, strategy="table-row") -> list[Candidate]:
    """Label/value layouts: ``Net asset value | USD | 231.856``."""
    out: list[Candidate] = []
    for row in rows:
        for c, cell in enumerate(row[:-1]):
            if not cell or _NOT_A_PRICE.search(cell) or not _matches(cell, matchers):
                continue
            for offset, nxt in enumerate(row[c + 1:], start=c + 1):
                value = parse_number(nxt)
                if value is None:
                    continue
                out.append(Candidate(field_name, cell, value,
                                     _row_date(row, skip=offset), strategy,
                                     " | ".join(row)[:200]))
                break
            break
    return out


def _text_candidates(text, matchers, field_name) -> list[Candidate]:
    """Last resort: ``<label> ... <number>`` anywhere in the visible text."""
    out: list[Candidate] = []
    for matcher in matchers:
        for hit in matcher.finditer(text):
            tail = text[hit.end():hit.end() + 60]
            if _NOT_A_PRICE.search(tail[:20]):
                continue
            value = parse_number(tail)
            if value is None:
                continue
            window = text[max(0, hit.start() - 60):hit.end() + 60]
            out.append(Candidate(field_name, _squash(hit.group(0)), value,
                                 parse_date(tail) or parse_date(window),
                                 "text", _squash(window)[:200]))
    return out


def _page_date(doc: _DocParser) -> Optional[dt.date]:
    """A date the page states for its prices, e.g. ``Date: 09.09.2026``."""
    pairs = list(doc.pairs)
    for table in doc.tables:
        for row in table:
            for c, cell in enumerate(row[:-1]):
                pairs.append((cell, row[c + 1]))
    for label, value in pairs:
        if label and DATE_LABEL.search(label):
            found = parse_date(value)
            if found:
                return found
    hit = re.search(
        r"(?:as\s+of|valuation\s+date|nav\s+date|price\s+date|date|stichtag|datum)\b"
        r"\D{0,20}(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})",
        doc.text, re.I)
    return parse_date(hit.group(1)) if hit else None


_STRATEGY_RANK = {"table-column": 0, "table-row": 1, "definition-list": 2, "text": 3}


def extract_candidates(html: str,
                       field_order: Sequence[str] = DEFAULT_FIELD_ORDER,
                       label_regex: Optional[str] = None) -> list[Candidate]:
    """Every price-looking label/value pair on the page, best first."""
    doc = _parse_document(html)
    fallback_date = _page_date(doc)

    if label_regex:
        plan = [("custom", _compile([label_regex]))]
    else:
        plan = []
        for name in field_order:
            if name not in FIELD_PATTERNS:
                raise ValueError("unknown field %r (known: %s)"
                                 % (name, ", ".join(sorted(FIELD_PATTERNS))))
            plan.append((name, _compile(FIELD_PATTERNS[name])))

    found: list[Candidate] = []
    for rank, (name, matchers) in enumerate(plan):
        collected: list[Candidate] = []
        for table in doc.tables:
            collected += _column_candidates(table, matchers, name)
            collected += _row_candidates(table, matchers, name)
        collected += _row_candidates([[a, b] for a, b in doc.pairs], matchers, name,
                                     strategy="definition-list")
        collected += _text_candidates(doc.text, matchers, name)
        for cand in collected:
            if cand.date is None:
                cand.date = fallback_date
            found.append((rank, _STRATEGY_RANK.get(cand.strategy, 9), cand))

    found.sort(key=lambda item: (item[0], item[1]))
    ordered, seen = [], set()
    for _, _, cand in found:
        key = (cand.label.lower(), round(cand.value, 6), cand.date)
        if key in seen:
            continue
        seen.add(key)
        ordered.append(cand)
    return ordered


def _page_currency(html: str, default: str = FUND_CURRENCY) -> str:
    hit = re.search(r"\b(USD|CHF|EUR|GBP|JPY)\b", html)
    return hit.group(1) if hit else default


# ---------------------------------------------------------------------------
# Fetching
# ---------------------------------------------------------------------------

def fetch_url(url: str, timeout: int = TIMEOUT, retries: int = RETRIES) -> tuple[str, str]:
    """GET ``url`` and return ``(text, content_type)``. Retries transient errors."""
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en,de;q=0.8",
        "Accept-Encoding": "gzip, deflate",
    }
    last: Optional[Exception] = None
    for attempt in range(retries):
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read()
                encoding = (response.headers.get("Content-Encoding") or "").lower()
                content_type = response.headers.get("Content-Type") or ""
            if encoding == "gzip":
                raw = gzip.decompress(raw)
            elif encoding == "deflate":
                try:
                    raw = zlib.decompress(raw)
                except zlib.error:
                    raw = zlib.decompress(raw, -zlib.MAX_WBITS)
            charset = "utf-8"
            hit = re.search(r"charset=([\w-]+)", content_type, re.I)
            if hit:
                charset = hit.group(1)
            return raw.decode(charset, errors="replace"), content_type
        except urllib.error.HTTPError as exc:
            last = exc
            if 400 <= exc.code < 500 and exc.code != 429:
                break            # not worth retrying
        except Exception as exc:
            last = exc
        if attempt < retries - 1:
            time.sleep(2 ** attempt)
    raise FetchError("could not fetch %s: %s" % (url, last))


def _looks_like_csv(text: str, content_type: str) -> bool:
    if "csv" in content_type.lower():
        return True
    head = text.lstrip()[:2000]
    return "<" not in head[:200] and "Net Asset Value" in head


def fetch_current_price(url: str = FUND_URL,
                        field_order: Sequence[str] = DEFAULT_FIELD_ORDER,
                        label_regex: Optional[str] = None,
                        html: Optional[str] = None,
                        assume_today: bool = False,
                        timeout: int = TIMEOUT) -> PricePoint:
    """Scrape the fund page and return the current price. No side effects.

    Pass ``html=`` to parse a saved page instead of fetching. If the response
    is the Swiss Fund Data CSV export rather than HTML, its last row is used.
    """
    if html is None:
        html, content_type = fetch_url(url, timeout=timeout)
    else:
        content_type = "text/html"

    if _looks_like_csv(html, content_type):
        series = _parse_csv_payload(html)
        if not series:
            raise ParseError("CSV response from %s held no Net Asset Value rows" % url)
        when, value = series[-1]
        return PricePoint(date=when, value=value, currency=FUND_CURRENCY,
                          label="Net Asset Value (CSV export)", source="swissfunddata",
                          retrieved_at=_utcnow())

    candidates = extract_candidates(html, field_order, label_regex)
    if not candidates:
        raise ParseError(
            "no price found on %s — the page layout may have changed. "
            "Run `probe` to see what the page offers, then pin the right row "
            "with --label-regex." % url)

    best = candidates[0]
    when, assumed = best.date, False
    if when is None:
        if not assume_today:
            raise ParseError(
                "found %s = %s on %s but no price date; the value cannot be "
                "filed against a day. Run `probe` to check, then re-run with "
                "--assume-today to stamp it with today's date."
                % (best.label, best.value, url))
        when, assumed = local_today(), True

    return PricePoint(date=when, value=best.value, currency=_page_currency(html),
                      label=best.label, source="swissfunddata",
                      retrieved_at=_utcnow(), date_is_assumed=assumed)


# ---------------------------------------------------------------------------
# CSV import (the Swiss Fund Data chart export)
# ---------------------------------------------------------------------------

def _parse_csv_payload(text: str) -> list[tuple[dt.date, float]]:
    """Read ``Date`` + ``Net Asset Value`` out of a Swiss Fund Data CSV export.

    The export carries a fund-name title line above the real header, so the
    header row is located by content rather than by position.
    """
    rows = list(csv.reader(io.StringIO(text)))
    header_at, date_col, value_col = None, None, None
    for index, row in enumerate(rows[:20]):
        cells = [_squash(cell).lower() for cell in row]
        if "date" not in cells:
            continue
        for wanted in ("net asset value", "chart price", "closing price"):
            if wanted in cells:
                header_at, date_col, value_col = index, cells.index("date"), cells.index(wanted)
                break
        if header_at is not None:
            break
    if header_at is None:
        raise ParseError("no 'Date' + 'Net Asset Value' header found in the CSV")

    series: list[tuple[dt.date, float]] = []
    for row in rows[header_at + 1:]:
        if len(row) <= max(date_col, value_col):
            continue
        when, value = parse_date(row[date_col]), parse_number(row[value_col])
        if when is None or value is None:
            continue
        series.append((when, value))
    series.sort(key=lambda item: item[0])
    return series


def import_csv(csv_path: str, workbook: str = DEFAULT_WORKBOOK,
               currency: str = FUND_CURRENCY, merge: bool = False) -> int:
    """Build (or top up) the workbook from a downloaded CSV export.

    Returns the number of rows in the resulting workbook. With ``merge=False``
    (the default) the CSV replaces the history; with ``merge=True`` its rows
    are merged into whatever is already stored.
    """
    with open(csv_path, encoding="utf-8-sig", newline="") as handle:
        series = _parse_csv_payload(handle.read())
    if not series:
        raise ParseError("no usable rows in %s" % csv_path)

    stamp = _utcnow().isoformat(timespec="seconds")
    existing = read_history(workbook) if (merge and os.path.exists(workbook)) else []
    by_date = {row.date: row for row in existing}
    for when, value in series:
        by_date[when] = HistoryRow(when, value, currency, "csv-import", stamp)
    rows = sorted(by_date.values(), key=lambda row: row.date)
    write_history(workbook, rows)
    return len(rows)


# ---------------------------------------------------------------------------
# Workbook storage
# ---------------------------------------------------------------------------

def _openpyxl():
    try:
        import openpyxl
    except ImportError as exc:            # pragma: no cover - environment issue
        raise NavTrackerError(
            "openpyxl is required to read and write the .xlsx file — "
            "install it with `python -m pip install openpyxl`") from exc
    return openpyxl


def read_history(workbook: str = DEFAULT_WORKBOOK) -> list[HistoryRow]:
    """Read the stored series, oldest first."""
    if not os.path.exists(workbook):
        raise NavTrackerError(
            "workbook not found: %s — create it first with "
            "`python nav_tracker.py import-csv <export.csv>`" % workbook)
    openpyxl = _openpyxl()
    book = openpyxl.load_workbook(workbook, data_only=True)
    sheet = book[SHEET_NAME] if SHEET_NAME in book.sheetnames else book.worksheets[0]
    rows: list[HistoryRow] = []
    for raw in sheet.iter_rows(min_row=2, values_only=True):
        if raw is None or raw[0] is None:
            continue
        cell = raw[0]
        when = cell.date() if isinstance(cell, dt.datetime) else cell
        if not isinstance(when, dt.date):
            when = parse_date(str(cell))
        value = raw[1] if len(raw) > 1 else None
        if when is None or value is None:
            continue
        rows.append(HistoryRow(
            date=when,
            value=float(value),
            currency=str(raw[2]) if len(raw) > 2 and raw[2] else FUND_CURRENCY,
            source=str(raw[3]) if len(raw) > 3 and raw[3] else "",
            retrieved_at=str(raw[4]) if len(raw) > 4 and raw[4] else "",
        ))
    book.close()
    rows.sort(key=lambda row: row.date)
    return rows


def write_history(workbook: str = DEFAULT_WORKBOOK,
                  rows: Sequence[HistoryRow] = ()) -> None:
    """Rewrite the workbook from ``rows``. The file is replaced atomically."""
    openpyxl = _openpyxl()
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = SHEET_NAME
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        sheet.append([row.date, row.value, row.currency, row.source, row.retrieved_at])
    for line in sheet.iter_rows(min_row=2, max_col=2):
        line[0].number_format = "DD/MM/YYYY"
        line[1].number_format = "0.000"

    sheet.freeze_panes = "A2"
    last = sheet.max_row
    if last > 1:
        sheet.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(HEADERS)), last)
    for index, width in enumerate((12, 18, 10, 26, 22), start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    if last > 2:
        chart = LineChart()
        chart.title = "%s — Net Asset Value" % FUND_NAME
        chart.y_axis.title = FUND_CURRENCY
        chart.x_axis.title = "Date"
        chart.height, chart.width = 9, 24
        chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=last))
        series = chart.series[0]
        series.smooth = False
        series.graphicalProperties.line.width = 14000
        sheet.add_chart(chart, "G2")

    meta = book.create_sheet(META_SHEET)
    newest = rows[-1] if rows else None
    for key, value in (
        ("Fund", FUND_NAME),
        ("ISIN", FUND_ISIN),
        ("Currency", FUND_CURRENCY),
        ("Source", FUND_URL),
        ("Observations", len(rows)),
        ("First date", rows[0].date.isoformat() if rows else ""),
        ("Last date", newest.date.isoformat() if newest else ""),
        ("Last Net Asset Value", newest.value if newest else ""),
        ("Workbook written (UTC)", _utcnow().isoformat(timespec="seconds")),
        ("Maintained by", "nav_tracker.py %s" % __version__),
    ):
        meta.append([key, value])
    for cell in meta["A"]:
        cell.font = Font(bold=True)
    meta.column_dimensions["A"].width = 24
    meta.column_dimensions["B"].width = 62

    folder = os.path.dirname(os.path.abspath(workbook))
    os.makedirs(folder, exist_ok=True)
    temp = os.path.join(folder, ".%s.tmp" % os.path.basename(workbook))
    book.save(temp)
    os.replace(temp, workbook)


# ---------------------------------------------------------------------------
# The routine
# ---------------------------------------------------------------------------

def update(url: str = FUND_URL,
           workbook: str = DEFAULT_WORKBOOK,
           field_order: Sequence[str] = DEFAULT_FIELD_ORDER,
           label_regex: Optional[str] = None,
           html: Optional[str] = None,
           assume_today: bool = False,
           force: bool = False,
           max_move_pct: float = MAX_MOVE_PCT,
           dry_run: bool = False,
           timeout: int = TIMEOUT) -> UpdateResult:
    """Fetch the current price and append it to the workbook.

    Safe to run repeatedly: a date already stored is left alone unless the
    value changed, in which case it is corrected in place. Raises
    :class:`SanityError` if the scraped number looks implausible.
    """
    rows = read_history(workbook)
    newest = rows[-1] if rows else None
    point = fetch_current_price(url, field_order=field_order, label_regex=label_regex,
                                html=html, assume_today=assume_today, timeout=timeout)

    if point.value <= 0:
        raise SanityError("scraped a non-positive price (%s)" % point.value)
    if point.date > local_today() + dt.timedelta(days=1):
        raise SanityError("scraped price is dated %s, which is in the future"
                          % point.date)
    if newest and not force:
        move = abs(point.value - newest.value) / newest.value * 100.0
        if move > max_move_pct:
            raise SanityError(
                "scraped %s = %.4f, a %.1f%% move from the last stored value "
                "%.4f on %s (limit %.1f%%). Check the page with `probe`; "
                "re-run with --force if the move is genuine."
                % (point.label, point.value, move, newest.value,
                   newest.date, max_move_pct))

    # With an assumed date we cannot tell a fresh publication from a stale
    # page, so an unchanged value is treated as "nothing new today".
    if (point.date_is_assumed and newest and point.date > newest.date
            and abs(point.value - newest.value) < 1e-9):
        return UpdateResult("skipped-stale", point, workbook, len(rows), newest,
                            "page still shows %.4f from %s; nothing appended"
                            % (newest.value, newest.date))

    by_date = {row.date: row for row in rows}
    previous = by_date.get(point.date)
    if previous and abs(previous.value - point.value) < 1e-9:
        return UpdateResult("unchanged", point, workbook, len(rows), previous,
                            "%s already stored as %.4f" % (point.date, point.value))

    status = "updated" if previous else "added"
    source = point.source + (" (assumed date)" if point.date_is_assumed else "")
    by_date[point.date] = HistoryRow(
        point.date, point.value, point.currency, source,
        (point.retrieved_at or _utcnow()).isoformat(timespec="seconds"))
    merged = sorted(by_date.values(), key=lambda row: row.date)

    if not dry_run:
        write_history(workbook, merged)
    return UpdateResult(
        status, point, workbook, len(merged), previous,
        "%s %s = %.4f %s%s" % (
            "would record" if dry_run else "recorded", point.date, point.value,
            point.currency, " (replacing %.4f)" % previous.value if previous else ""))


# ---------------------------------------------------------------------------
# Command line interface
# ---------------------------------------------------------------------------

def _env(name: str, fallback: str) -> str:
    return os.environ.get(name) or fallback


def _add_source_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--url", default=_env("NAV_TRACKER_URL", FUND_URL),
                        help="page to scrape (default: %(default)s)")
    parser.add_argument("--html", metavar="PATH",
                        help="parse this saved page instead of fetching")
    parser.add_argument("--field", default=None, choices=sorted(FIELD_PATTERNS),
                        help="which figure to read (default: net asset value, "
                             "then any current price)")
    parser.add_argument("--label-regex", metavar="RX",
                        help="pin the extraction to labels matching this regex")
    parser.add_argument("--timeout", type=int, default=TIMEOUT,
                        help="HTTP timeout in seconds (default: %(default)s)")


def _add_workbook_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--workbook", "-w",
                        default=_env("NAV_TRACKER_WORKBOOK", DEFAULT_WORKBOOK),
                        help="xlsx file to maintain (default: %(default)s)")


def _source_kwargs(args) -> dict:
    html = None
    if getattr(args, "html", None):
        with open(args.html, encoding="utf-8", errors="replace") as handle:
            html = handle.read()
    return {
        "url": args.url,
        "html": html,
        "field_order": (args.field,) if args.field else DEFAULT_FIELD_ORDER,
        "label_regex": args.label_regex,
        "timeout": args.timeout,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="nav_tracker",
        description=__doc__.splitlines()[0],
        formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("import-csv", help="build the workbook from a CSV export")
    p.add_argument("csv", help="Swiss Fund Data CSV export")
    _add_workbook_arg(p)
    p.add_argument("--merge", action="store_true",
                   help="merge into the existing workbook instead of replacing it")
    p.add_argument("--currency", default=FUND_CURRENCY)

    p = sub.add_parser("probe", help="show every price the page offers, best first")
    _add_source_args(p)
    p.add_argument("--save-html", metavar="PATH", help="also write the raw page here")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("fetch", help="scrape the current price, change nothing")
    _add_source_args(p)
    p.add_argument("--assume-today", action="store_true",
                   help="use today's date when the page states none")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("update", help="scrape the current price and append it")
    _add_source_args(p)
    _add_workbook_arg(p)
    p.add_argument("--assume-today", action="store_true",
                   help="use today's date when the page states none")
    p.add_argument("--force", action="store_true",
                   help="accept a value outside the plausibility band")
    p.add_argument("--max-move", type=float, default=MAX_MOVE_PCT, metavar="PCT",
                   help="reject moves larger than this (default: %(default)s%%)")
    p.add_argument("--dry-run", action="store_true", help="do not write the file")
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("show", help="print the tail of the stored series")
    _add_workbook_arg(p)
    p.add_argument("-n", type=int, default=10, help="rows to show (default: %(default)s)")
    p.add_argument("--json", action="store_true")
    return parser


def _cmd_import_csv(args) -> int:
    count = import_csv(args.csv, args.workbook, args.currency, merge=args.merge)
    rows = read_history(args.workbook)
    print("Wrote %d observations to %s" % (count, args.workbook))
    print("Range: %s to %s   Latest: %.4f %s"
          % (rows[0].date, rows[-1].date, rows[-1].value, rows[-1].currency))
    return 0


def _cmd_probe(args) -> int:
    kwargs = _source_kwargs(args)
    html = kwargs["html"]
    if html is None:
        html, _ = fetch_url(args.url, timeout=args.timeout)
    if args.save_html:
        with open(args.save_html, "w", encoding="utf-8") as handle:
            handle.write(html)
    candidates = extract_candidates(html, kwargs["field_order"], kwargs["label_regex"])
    if args.json:
        print(json.dumps([{
            "field": c.field, "label": c.label, "value": c.value,
            "date": c.date.isoformat() if c.date else None,
            "strategy": c.strategy, "context": c.context,
        } for c in candidates], indent=2))
        return 0 if candidates else 2

    print("Source: %s (%d bytes)" % (args.html or args.url, len(html)))
    if not candidates:
        print("\nNo price-looking label/value pair found.")
        print("Save the page with --save-html and check what it actually contains.")
        return 2
    print("\n%d candidate(s), best first — the first one is what `update` uses:\n"
          % len(candidates))
    for index, c in enumerate(candidates, start=1):
        mark = ">>" if index == 1 else "  "
        print("%s %2d. %-34.34s %14.4f   %-10s  [%s/%s]"
              % (mark, index, c.label, c.value,
                 c.date.isoformat() if c.date else "no date", c.field, c.strategy))
        if c.context:
            print("        context: %s" % c.context[:120])
    return 0


def _cmd_fetch(args) -> int:
    point = fetch_current_price(assume_today=args.assume_today, **_source_kwargs(args))
    if args.json:
        print(json.dumps(point.as_dict(), indent=2))
    else:
        print("%s  %.4f %s   (%s%s)"
              % (point.date, point.value, point.currency, point.label,
                 ", date assumed" if point.date_is_assumed else ""))
    return 0


def _cmd_update(args) -> int:
    result = update(workbook=args.workbook, assume_today=args.assume_today,
                    force=args.force, max_move_pct=args.max_move,
                    dry_run=args.dry_run, **_source_kwargs(args))
    if args.json:
        print(json.dumps(result.as_dict(), indent=2))
    else:
        print("%s: %s" % (result.status, result.message))
        print("Workbook: %s (%d observations)" % (result.workbook, result.rows))
    return 0


def _cmd_show(args) -> int:
    rows = read_history(args.workbook)
    tail = rows[-args.n:] if args.n > 0 else rows
    if args.json:
        print(json.dumps([{
            "date": r.date.isoformat(), "value": r.value, "currency": r.currency,
            "source": r.source, "retrieved_at": r.retrieved_at,
        } for r in tail], indent=2))
        return 0
    print("%d observations, %s to %s" % (len(rows), rows[0].date, rows[-1].date))
    print("%-12s %12s  %s" % ("Date", "NAV", "Source"))
    for row in tail:
        print("%-12s %12.3f  %s" % (row.date, row.value, row.source))
    return 0


_COMMANDS = {
    "import-csv": _cmd_import_csv,
    "probe": _cmd_probe,
    "fetch": _cmd_fetch,
    "update": _cmd_update,
    "show": _cmd_show,
}


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return _COMMANDS[args.command](args)
    except SanityError as exc:
        print("sanity check failed: %s" % exc, file=sys.stderr)
        return 3
    except (FetchError, ParseError) as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 2
    except NavTrackerError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 1
    except OSError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
