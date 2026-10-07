#!/usr/bin/env python3
"""
Automation Tool: for each company in a spreadsheet, find its latest annual accounts on
Companies House and pull out the turnover (revenue) figure.

How it reads the accounts, best first:
  1. iXBRL (electronically filed accounts): the turnover figure is tagged, read exactly.
  2. iXBRL tables, when the figure is shown but not tagged.
  3. PDF: text PDFs are read directly; scanned page images are read with OCR and marked
     VERIFY, because OCR can misread digits.

Small and micro companies are allowed to leave the profit and loss account out of what
they file, so for many companies there is no turnover to find. Those rows are flagged
with the reason instead.

Usage:
    python turnover.py clients.xlsx --name-col "Company Name" --number-col "Company No"

Install:
    pip install -r requirements.txt
"""

import argparse
import re
import sys
import time
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime

import numpy as np
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz

BASE = "https://find-and-update.company-information.service.gov.uk"
WEB_DELAY = 1.5           # seconds between requests, be polite to the website
RETRIES = 4               # attempts per request for network errors and server errors
MATCH_THRESHOLD = 90      # name similarity needed to accept a search result
FILING_PAGES = 20         # pages of filing history (25 filings each) to look through
STALE_DAYS = 730          # accounts older than this get a note

TURNOVER_TAGS = {"TurnoverRevenue", "Revenue", "TurnoverGrossOperatingRevenue"}
ACCOUNT_TYPES = {"AA", "AAMD"}  # accounts, amended accounts
STATUSES = ["FOUND", "VERIFY", "NOT DISCLOSED", "CHECK PDF", "NOT FOUND", "ERROR"]

# ---------------------------------------------------------------- results


@dataclass
class Result:
    status: str = "NOT FOUND"       # one of STATUSES
    turnover: float | None = None
    prior_turnover: float | None = None
    currency: str = ""
    period_end: str = ""
    company_name: str = ""
    company_number: str = ""
    company_status: str = ""
    accounts_type: str = ""
    filed_on: str = ""
    method: str = ""
    note: str = ""
    pdf_link: str = ""
    last_accounts: str | None = None    # from the company page: date, '' = none filed, None = unknown
    keywords_found: list | None = None  # None = keywords weren't searched (no document, or none given)
    keyword_counts: dict = field(default_factory=dict)  # keyword -> times it appears
    notes: list = field(default_factory=list)

    @property
    def company_link(self):
        return f"{BASE}/company/{self.company_number}" if self.company_number else ""


class SourceError(Exception):
    """Companies House couldn't be reached or refused the request."""


# ---------------------------------------------------------------- web access


class Web:
    def __init__(self, delay=WEB_DELAY):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0 (Automation Tool)"
        self.delay = delay
        self._last = 0.0

    def get(self, path, params=None):
        """GET a page; None for 404. Retries network errors, server errors and rate limits."""
        url = path if path.startswith("http") else BASE + path
        problem = ""
        for attempt in range(RETRIES):
            wait = self.delay - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.time()
            try:
                r = self.s.get(url, params=params, timeout=60)
            except requests.RequestException as e:
                problem = f"network error ({type(e).__name__})"
                time.sleep(5 * (attempt + 1))
                continue
            if r.status_code == 404:
                return None
            if r.status_code == 429:
                problem = "rate limited by Companies House"
                time.sleep(45 * (attempt + 1))
                continue
            if r.status_code >= 500:
                problem = f"Companies House server error (HTTP {r.status_code})"
                time.sleep(10 * (attempt + 1))
                continue
            if r.status_code >= 400:
                raise SourceError(f"Companies House refused the request (HTTP {r.status_code})")
            return r
        raise SourceError(f"{problem}, tried {RETRIES} times; run this row again later")

    def soup(self, path, params=None):
        r = self.get(path, params)
        return BeautifulSoup(r.text, "html.parser") if r is not None else None


# ---------------------------------------------------------------- finding the company


def normalise(name):
    s = str(name or "").upper().replace("&", " AND ")
    s = re.sub(r"\bP\.\s*L\.\s*C\.?", "PLC", s)              # "P.L.C." as the register writes it
    s = re.sub(r"\bL\.\s*T\.\s*D\.?", "LTD", s)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\bLTD\b", "LIMITED", s)
    s = re.sub(r"\bPLC\b", "PUBLIC LIMITED COMPANY", s)
    s = re.sub(r"\b(PUBLIC LIMITED COMPANY|LIMITED LIABILITY PARTNERSHIP|LIMITED|COMPANY)\b", " ", s)
    s = re.sub(r"^THE\s+", "", s.strip())
    return re.sub(r"\s+", " ", s).strip()


def name_score(a, b):
    na, nb = normalise(a), normalise(b)
    return 100.0 if na == nb else max(fuzz.token_sort_ratio(na, nb), fuzz.ratio(na, nb))


def clean_number(value):
    """Company numbers as people type them, or as Excel mangles them: 445790, 445790.0,
    'sc 5261', ' 00445790 ' -> '00445790', 'SC005261'. '' when there's nothing usable."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()
    if s.isdigit():
        return s.zfill(8)
    m = re.fullmatch(r"([A-Z]{2})(\d{1,6})", s)
    return m.group(1) + m.group(2).zfill(6) if m else s


def read_profile(soup, res):
    """Name, status and accounts facts from a company page. Returns the 'last accounts' date
    text, '' when the company hasn't filed accounts yet, or None when the page doesn't say."""
    h1 = soup.select_one("#company-name") or soup.find("h1")
    res.company_name = h1.get_text(" ", strip=True) if h1 else res.company_name
    st = soup.select_one("#company-status")
    res.company_status = st.get_text(" ", strip=True) if st else ""
    if res.company_status.lower() in ("dissolved", "liquidation", "administration", "receivership",
                                      "insolvency proceedings", "voluntary arrangement", "closed"):
        res.notes.append(f"company status: {res.company_status.lower()}")
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    if re.search(r"Accounts overdue", text, re.I):
        res.notes.append("accounts overdue at Companies House")
    m = re.search(r"Last accounts made up to (\d{1,2} \w+ \d{4})", text)
    if m:
        return m.group(1)
    if re.search(r"First accounts made up to", text):
        return ""
    return None


def find_company(web, name, number, res):
    """Look the company up by number (exact) or name (search). Fills in res; returns the
    company page soup, or None with res.note saying why."""
    number = clean_number(number)
    if number:
        soup = web.soup(f"/company/{number}")
        if not soup:
            res.note = f"company number {number} not on the register"
            return None
        res.company_number = number
        res.last_accounts = read_profile(soup, res)
        if name and res.company_name and name_score(name, res.company_name) < 70:
            res.notes.append(f"the number belongs to {res.company_name}, check it matches '{name}'")
        return soup

    if not name:
        res.note = "no company name or number in this row"
        return None
    search = web.soup("/search/companies", {"q": name})
    hits = []
    for li in (search.select("#results li") if search else []):
        a = li.find("a", href=re.compile(r"^/company/[A-Z0-9]+"))
        if a:
            hits.append((a.get_text(" ", strip=True), a["href"].split("/")[2]))
    if not hits:
        res.note = "company not found on the register"
        return None
    best = max(hits, key=lambda h: name_score(name, h[0]))
    score = name_score(name, best[0])
    if score < MATCH_THRESHOLD:
        res.note = f"no confident match (closest: {best[0]}, {score:.0f}%), add the company number"
        return None
    if any(n != best[1] and name_score(name, t) >= MATCH_THRESHOLD for t, n in hits):
        res.notes.append("several companies with near-identical names, add the company number to be sure")
    res.company_name, res.company_number = best
    soup = web.soup(f"/company/{best[1]}")
    if soup:
        res.last_accounts = read_profile(soup, res)
    return soup or BeautifulSoup("", "html.parser")


# ---------------------------------------------------------------- latest accounts filing


def latest_accounts(web, number, pages=FILING_PAGES):
    """Return (date, description, pdf_url, xhtml_url) of the newest annual accounts filing,
    or None. The website ignores its own 'accounts only' filter, so this pages through the
    whole filing history (newest first) until it meets an accounts filing."""
    for page in range(1, pages + 1):
        soup = web.soup(f"/company/{number}/filing-history", {"page": page})
        rows = soup.select("#fhTable tr")[1:] if soup else []
        for tr in rows:
            tds = tr.find_all("td")
            if len(tds) < 3 or tds[1].get_text(strip=True) not in ACCOUNT_TYPES:
                continue
            description = tds[2].get_text(" ", strip=True).split(" View PDF")[0]
            if re.match(r"\s*(interim|initial)\s+accounts", description, re.I):
                continue                    # not the annual accounts
            links = {a["href"] for a in tr.find_all("a", href=True)}
            pdf = next((l for l in links if "format=pdf" in l), "")
            xhtml = next((l for l in links if "format=xhtml" in l), "")
            return (tds[0].get_text(" ", strip=True), description,
                    BASE + pdf if pdf else "", BASE + xhtml if xhtml else "")
        if not rows or not soup.find("a", href=re.compile(rf"[?&]page={page + 1}\b")):
            return None
    return None


# ---------------------------------------------------------------- reading amounts from lines of text

LABEL = re.compile(r"^\s*(?:group\s*|total\s*)?(?:turnover|revenue)s?\b\s*(?:\((?!\s*\d)[^)]*\))?\s*"
                   r"(?:\(?\s*notes?\b[^)]*\)?)?(.*)$", re.I)   # allows "Revenue (including ...)"
WHOLE = re.compile(r"\d{1,3}(?:,\d{3})+|\d{3,}")         # 1,234,567 or 1234567
OCR_DOTS = re.compile(r"\d{1,3}(?:\.\d{3})+")             # 72.886: OCR read the comma as a dot
DECIMAL = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{1,2}")       # 2,151.2 (accounts in £m)
OCR_DOTS_DECIMAL = re.compile(r"\d{1,3}(?:\.\d{3})+\.\d{1,2}")  # 1.771.0: OCR read 1,771.0 with dots


def amounts(rest):
    """Pull money amounts out of the text after a 'Turnover' label, skipping note references
    (1-2 digit numbers like '3' or '2,3'). A small decimal like '2.1' could be a note number,
    so those only count when the row has several; 1,000.0 and up always count."""
    found = []
    for tok in re.split(r"\s+", rest.replace("£", "").strip()):
        tok = tok.strip("|:;")
        neg = tok.startswith("(") or tok.startswith("-")
        core = tok.strip("()-–")
        if re.fullmatch(r"[0,.]+", core):                    # "£000" is a unit heading, not zero
            continue
        if WHOLE.fullmatch(core):
            found.append((float(core.replace(",", "")), neg, False))
        elif OCR_DOTS.fullmatch(core):
            found.append((float(core.replace(".", "")), neg, False))
        elif DECIMAL.fullmatch(core):
            v = float(core.replace(",", ""))
            found.append((v, neg, v < 1000))
        elif OCR_DOTS_DECIMAL.fullmatch(core):
            whole, _, dec = core.rpartition(".")
            found.append((float(whole.replace(".", "") + "." + dec), neg, False))
    if sum(1 for f in found if f[2]) < 2:
        found = [f for f in found if not f[2]]
    return [-v if neg else v for v, neg, _ in found]


def turnover_from_lines(lines):
    """lines: strings, one per row of a table. Returns (current, prior, row_units, row_index)
    or None, where row_units is (multiplier, label) when the row itself says '£000' or '£m'."""
    for idx, line in enumerate(lines):
        m = LABEL.match(line)
        if not m:
            continue
        # Stop at the next word: either the label carries on ("Revenue from sale of goods", a
        # sub-line, so no amounts before it) or OCR merged in a neighbouring column's row.
        rest = re.split(r"[A-Za-z]{3,}", m.group(1), maxsplit=1)[0]
        vals = amounts(rest)
        if vals:
            units = unit_multiplier(rest)
            return vals[0], (vals[1] if len(vals) > 1 else None), (units if units[0] > 1 else None), idx
    return None


def scaled(found, lines):
    """Apply the units (£, £000, £m) to a turnover_from_lines result. Units come from the row
    itself or the column headings just above it, never the whole page: a stray 'm' elsewhere
    would otherwise turn pounds into millions."""
    idx = found[3]
    mult, label = found[2] or unit_multiplier(" | ".join(lines[max(0, idx - 8):idx]))
    return round(found[0] * mult), (round(found[1] * mult) if found[1] is not None else None), label


def unit_multiplier(text):
    """Column headers like '£m', '€ million' or '£000' say what the figures are in.
    OCR often reads £ as E and sometimes drops the symbol altogether."""
    tokens = set(re.split(r"[\s|]+", text))
    if any(re.fullmatch(r"[£Ef€$]?[’']?(m|million)", t) for t in tokens):   # £m, £'m, Em, f'm, € million
        return 1_000_000, "m"
    if any(re.fullmatch(r"[£E€$]?[’']?000|[£€$]k|[£E€$]?[’']000s?", t) for t in tokens):
        return 1_000, "000"
    return 1, ""


def page_currency(text):
    """GBP unless the page shows another currency. '' when no symbol is visible at all."""
    if re.search(r"€|\beuros?\b", text, re.I):
        return "EUR"
    if re.search(r"US\$|\$m|\$'?000|\bUSD\b|dollars", text):
        return "USD"
    if re.search(r"£|\b[Ef][m’']\b|\bE\s*000|\bGBP\b|sterling", text):
        return "GBP"
    return ""


# ---------------------------------------------------------------- keywords


def _norm(text):
    """Lowercase, rejoin words split across lines ('ex-\nchange'), punctuation to spaces."""
    text = re.sub(r"-\s*\n\s*", "", str(text).lower())
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", text)).strip()


def parse_keywords(text):
    """Keywords typed one per line, or separated by commas or semicolons."""
    return [k.strip() for k in re.split(r"[\n,;]+", text or "") if k.strip()]


class Keywords:
    """Finds which of a list of keywords appear in a document's text, and how often.

    Matching ignores case and punctuation, and counts whole words or phrases only, so
    'derivative' doesn't count 'derivatives' and 'FX' doesn't match inside other words. In OCR
    text, a phrase or long word (8+ letters) that isn't found that way is looked for with spaces
    ignored, because OCR often runs words together ("Foreignexchange")."""

    def __init__(self, words):
        self.words, seen = [], set()
        for w in words:
            n = _norm(w)
            if n and n not in seen:
                seen.add(n)
                self.words.append(w.strip())
        self._tests = []
        for w in self.words:
            n = _norm(w)
            compact = n.replace(" ", "")
            whole = re.compile(rf"(?<![a-z0-9]){re.escape(n)}(?![a-z0-9])")
            self._tests.append((w, whole, compact if " " in n or len(compact) >= 8 else None))
        self.found = set()
        self.counts = {}

    def scan(self, text, ocr=False):
        """Add one piece of text (a page, or a whole document) to the counts. ocr=True for text
        read by OCR, which allows the spaces-ignored fallback."""
        if not text:
            return
        t = _norm(text)
        tc = t.replace(" ", "")
        for w, whole, compact in self._tests:
            n = len(whole.findall(t))
            if not n and compact and ocr:               # OCR may have run the words together
                n = tc.count(compact)
            if n:
                self.found.add(w)
                self.counts[w] = self.counts.get(w, 0) + n

    @property
    def done(self):
        return len(self.found) == len(self.words)

    def result(self):
        """Found keywords in the order they were given."""
        return [w for w in self.words if w in self.found]

    def ranked(self):
        """{keyword: count} for the keywords found, most mentioned first (ties keep the given order)."""
        order = {w: i for i, w in enumerate(self.words)}
        return dict(sorted(self.counts.items(), key=lambda kv: (-kv[1], order[kv[0]])))


# ---------------------------------------------------------------- iXBRL


def parse_number(el):
    txt = el.get_text("", strip=True)
    fmt = (el.get("format") or "").lower()
    if re.fullmatch(r"[-–—]*", txt) or "zerodash" in fmt or "fixed-zero" in fmt:
        return 0.0
    if "numcommadecimal" in fmt or "num-comma-decimal" in fmt:
        txt = txt.replace(".", "").replace(" ", "").replace(",", ".")
    else:
        txt = txt.replace(",", "").replace(" ", "")
    try:
        v = float(re.sub(r"[^\d.]", "", txt))
    except ValueError:
        return None
    v *= 10 ** int(el.get("scale") or 0)
    return -v if el.get("sign") == "-" else v


def months_between(start, end):
    try:
        a, b = date.fromisoformat(start), date.fromisoformat(end)
    except (TypeError, ValueError):
        return None
    return round((b - a).days / 30.44)


def read_ixbrl(html):
    """Return (current, prior, currency, period_end, method, months) or None. months is the
    length of the accounting period when the filing says, else None. html may be a parsed soup."""
    soup = html if isinstance(html, BeautifulSoup) else BeautifulSoup(html, "html.parser")
    # Context priority: 0 = no dimensions, 1 = consolidated group figure, None = some other
    # breakdown (a segment, a subsidiary...) which is not the headline turnover.
    contexts = {}
    for c in soup.find_all(re.compile(r"(^|:)context$")):
        end = c.find(re.compile(r"(^|:)(enddate|instant)$"))
        start = c.find(re.compile(r"(^|:)startdate$"))
        members = [m.get_text(strip=True) for m in c.find_all(re.compile(r"(^|:)explicitmember$"))]
        dims = c.find(re.compile(r"(^|:)(segment|scenario)$"))
        if not dims:
            rank = 0
        elif members and all(m.split(":")[-1].startswith("Consolidated") for m in members):
            rank = 1
        else:
            rank = None
        contexts[c.get("id")] = (end.get_text(strip=True) if end else "",
                                 start.get_text(strip=True) if start else "", rank)
    units = {}
    for u in soup.find_all(re.compile(r"(^|:)unit$")):
        m = u.find(re.compile(r"(^|:)measure$"))
        units[u.get("id")] = m.get_text(strip=True).split(":")[-1] if m else ""

    facts = []
    for el in soup.find_all(re.compile(r"(^|:)nonfraction$")):
        if el.get("name", "").split(":")[-1] not in TURNOVER_TAGS:
            continue
        end, start, rank = contexts.get(el.get("contextref"), ("", "", None))
        v = parse_number(el)
        if v is not None and rank is not None and end:
            facts.append((rank, end, start, v, units.get(el.get("unitref"), "")))
    if facts:
        by_end = {}
        for _rank, end, start, v, cur in sorted(facts):
            by_end.setdefault(end, (v, cur, start))
        ends = sorted(by_end, reverse=True)
        cur_v, cur, start = by_end[ends[0]]
        prior = by_end[ends[1]][0] if len(ends) > 1 else None
        return cur_v, prior, (cur or "GBP").upper(), ends[0], "iXBRL tagged figure", months_between(start, ends[0])

    # Shown in a table but not tagged: read the table rows as text
    lines = [" ".join(td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"]))
             for tr in soup.find_all("tr")]
    found = turnover_from_lines(lines)
    if found:
        current, _prior, _ = scaled(found, lines)    # prior-year columns in loose tables are unreliable
        return current, None, "GBP", "", "iXBRL table (untagged)", None
    return None


def ixbrl_text(soup):
    """The readable text of an iXBRL document, without the hidden tagging header."""
    for header in soup.find_all(re.compile(r"(^|:)header$")):
        header.decompose()
    return soup.get_text(" ", strip=True)


# ---------------------------------------------------------------- PDF + OCR

_ocr = None


class OcrUnavailable(Exception):
    pass


def ocr_available():
    """Whether scanned PDFs can be read here (the OCR package only installs on Python <= 3.12)."""
    import importlib.util
    return importlib.util.find_spec("rapidocr_onnxruntime") is not None


def ocr_engine():
    global _ocr
    if _ocr is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _ocr = RapidOCR()
        except Exception as e:                  # missing package or model on the host
            raise OcrUnavailable(str(e)) from e
    return _ocr


def _rows(boxes):
    """Group OCR boxes [(y_mid, x_left, height, payload)] into rows, each sorted left to right."""
    if not boxes:
        return []
    tol = 0.6 * sorted(b[2] for b in boxes)[len(boxes) // 2]
    rows, current = [], []
    for b in sorted(boxes, key=lambda b: b[0]):
        if current and b[0] - current[-1][0] > tol:
            rows.append(current)
            current = []
        current.append(b)
    rows.append(current)
    return [sorted(r, key=lambda b: b[1]) for r in rows]


def _box(pts, payload):
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return ((min(ys) + max(ys)) / 2, min(xs), max(ys) - min(ys), payload)


def ocr_rows(image):
    """OCR a whole page; returns one string per row of text, left to right."""
    res, _ = ocr_engine()(np.asarray(image.convert("RGB")), use_cls=False)
    rows = _rows([_box(pts, text) for pts, text, _conf in res or []])
    return [" ".join(b[3] for b in r) for r in rows]


def ocr_skim(image):
    """Cheap pass: find all the text, but only read the first box on each row. That's enough to
    spot headings and row labels like 'Revenue', at a fraction of the cost of a full read."""
    arr = np.asarray(image.convert("RGB"))
    eng = ocr_engine()
    found, _ = eng(arr, use_cls=False, use_rec=False)
    if not found:
        return []
    firsts = [r[0] for r in _rows([_box(pts, pts) for pts in found])]
    crops = []
    for _y, _x, _h, pts in firsts:
        xs, ys = [int(p[0]) for p in pts], [int(p[1]) for p in pts]
        crops.append(arr[max(min(ys), 0):max(ys) + 1, max(min(xs), 0):max(xs) + 1])
    texts, _ = eng.text_rec(crops)
    return [t[0] for t in texts]


SKIM_LABEL = re.compile(r"^(group\s*|total\s*)?(turnover|revenue)s?\b", re.I)


def page_order(n):
    """Long annual reports put the financial statements after the strategic report and
    governance sections, so start a quarter of the way in and wrap round."""
    if n <= 40:
        return list(range(n))
    start = int(n * 0.25)
    return list(range(start, n)) + list(range(start))


STATEMENT_HEAD = re.compile(r"income\s*statement|profit\s*(and|&)\s*loss|statement\s*of\s*(comprehensive\s*)?income|"
                            r"statement\s*of\s*profit\s*(or|and)\s*loss", re.I)
NOTES_HEAD = re.compile(r"notes\s*to\s*the|notes\s*\(continued\)", re.I)


def is_statement_page(top_rows):
    """The page is a primary statement when its heading (first few rows) names one, and it
    isn't a note to the accounts (notes repeat 'Revenue' for subsidiaries and segments)."""
    top = " | ".join(top_rows[:8])
    return bool(STATEMENT_HEAD.search(top)) and not NOTES_HEAD.search(top)


@dataclass
class PdfResult:
    turnover: float | None = None
    unit: str = ""
    page: int = 0
    method: str = ""
    currency: str = ""
    pages: int = 0                # pages in the document
    pages_checked: int = 0
    keyword_pages: int = 0        # pages whose full text was searched for keywords
    scanned_skipped: bool = False
    keyword_ocr_skipped: bool = False
    timed_out: bool = False
    problem: str = ""


def read_pdf(data, max_pages=None, progress=None, read_scanned=True, keywords=None, ocr_keywords=True,
             deadline=None):
    """Find the turnover on the income statement page of an accounts PDF and, if keywords
    is given, search every page for them. Scanned pages are then read in full, which is slow;
    ocr_keywords=False searches only pages that have real text. max_pages (None = all) and
    deadline (a time.monotonic() value, None = no limit) stop the reading early."""
    import pypdfium2 as pdfium
    out = PdfResult()
    try:
        pdf = pdfium.PdfDocument(data)
        out.pages = len(pdf)
    except Exception as e:
        out.problem = f"the accounts PDF couldn't be opened ({type(e).__name__})"
        return out
    order = page_order(out.pages)[:max_pages]
    for done, i in enumerate(order, 1):
        want_keywords = keywords is not None             # counts need every page, so never stop early
        if out.turnover is not None and not want_keywords:
            break
        if deadline is not None and time.monotonic() > deadline:
            out.timed_out = True
            break
        out.pages_checked = done
        try:
            page = pdf[i]
            text = page.get_textpage().get_text_range()
        except Exception:
            continue                                       # one unreadable page, try the rest
        if text.strip():                                   # PDF has real text, no OCR needed
            rows, method = text.splitlines(), "PDF text"
            if want_keywords:
                keywords.scan(text)
                out.keyword_pages += 1
        elif not read_scanned:
            out.scanned_skipped = True
            continue
        else:
            if progress:
                what = "searching scanned accounts" if want_keywords else "reading scanned accounts"
                progress(f"{what}, page {done} of {len(order)}")
            if want_keywords and not ocr_keywords:
                want_keywords, out.keyword_ocr_skipped = False, True
            try:
                image = page.render(scale=150 / 72).to_pil()
                if want_keywords:                          # need every word on the page
                    rows = ocr_rows(image)
                    keywords.scan("\n".join(rows), ocr=True)
                    out.keyword_pages += 1
                else:                                      # just looking for the income statement
                    skim = ocr_skim(image)
                    if not (is_statement_page(skim) and any(SKIM_LABEL.match(t.strip()) for t in skim)):
                        continue
                    rows = ocr_rows(image)
                method = "PDF scan (OCR)"
            except OcrUnavailable:
                out.problem = ("this is a scanned PDF and the scanned-PDF reader isn't installed on this "
                               "server, open the PDF to check")
                return out
        if out.turnover is not None or not is_statement_page([r for r in rows if r.strip()]):
            continue
        found = turnover_from_lines(rows)
        if found:
            current, _prior, unit = scaled(found, rows)    # PDF column order varies, skip prior
            out.turnover, out.unit, out.page, out.method = current, unit, i + 1, method
            out.currency = page_currency(" | ".join(rows))
    return out


# ---------------------------------------------------------------- one company


def why_missing(description):
    """Reason a filing has no turnover, from its description; '' when the accounts type
    should normally include it (full, group, audited...)."""
    d = description.lower()
    if "dormant" in d:
        return "dormant company, no turnover"
    if "micro" in d:
        return "micro-entity accounts, turnover not required to be filed"
    if any(k in d for k in ("total exemption", "abridged", "small", "filleted")):
        return "small company accounts, profit and loss not filed"
    if "medium" in d:
        return "medium company accounts, which may leave turnover out"
    return ""


def apply_pdf(res, pdf):
    res.turnover, res.method = pdf.turnover, pdf.method
    res.currency = pdf.currency
    res.method += f", page {pdf.page}" + (f", figures in {pdf.currency or ''}{pdf.unit}" if pdf.unit else "")
    if not res.currency:
        res.notes.append("currency symbol not readable on the page, check it")
    res.status = "VERIFY" if "OCR" in res.method else "FOUND"
    if res.status == "VERIFY":
        res.notes.append("read by OCR, check the figure against the PDF")


def check(row, web, max_pages=None, progress=None, read_scanned=True, keywords=None,
          ocr_keywords=True, time_limit=0):
    """Find the latest turnover for one row: {'name': ..., 'number': ...}. keywords, if given,
    is a list of words or phrases to look for in the same accounts. time_limit is in minutes
    per company (0 = no limit); when it runs out, PDF reading stops and the row says so."""
    deadline = time.monotonic() + time_limit * 60 if time_limit else None
    res = Result()
    kw = Keywords(keywords) if keywords else None
    name = str(row.get("name") or "").strip()
    profile = find_company(web, name, row.get("number"), res)
    if profile is None:
        return finish(res)

    last = res.last_accounts
    if last == "":
        res.status, res.note = "NOT DISCLOSED", "no accounts filed yet (new company)"
        return finish(res)
    filing = latest_accounts(web, res.company_number)
    if not filing:
        res.status = "NOT DISCLOSED" if last is None else "CHECK PDF"
        res.note = ("no accounts filed" if last is None else
                    f"accounts made up to {last} exist but weren't found in the filing history, check the company page")
        return finish(res)
    res.filed_on, res.accounts_type, res.pdf_link, xhtml = filing
    m = re.search(r"made up to (\d{1,2} \w+ \d{4})", res.accounts_type)
    if m:
        made_up = datetime.strptime(m.group(1), "%d %B %Y")
        res.period_end = made_up.date().isoformat()
        if (datetime.now() - made_up).days > STALE_DAYS:
            res.notes.append(f"latest accounts are over 2 years old ({m.group(1)})")
    reason = why_missing(res.accounts_type)

    # 1. Electronic accounts: exact figure from the iXBRL.
    if xhtml:
        try:
            r = web.get(xhtml)
        except SourceError:
            r = None
        if r is None:
            found, xhtml = None, ""             # couldn't download it: fall back to the PDF
        else:
            soup = BeautifulSoup(r.content, "html.parser")
            found = read_ixbrl(soup)
            if kw:                              # the iXBRL is the whole document: search it all
                kw.scan(ixbrl_text(soup))
                res.keywords_found, res.keyword_counts = kw.result(), kw.ranked()
                kw = None                       # done, the PDF needn't be searched again
        if found:
            res.turnover, res.prior_turnover, res.currency, end, res.method, months = found
            res.period_end = end or res.period_end
            res.status = "FOUND"
            if months and not 11 <= months <= 13:
                res.notes.append(f"these accounts cover {months} months, not a normal year")
            return finish(res)
        if xhtml and reason:
            res.status, res.note = "NOT DISCLOSED", reason
            return finish(res)

    # 2. The PDF: no iXBRL, iXBRL failed to download, or full accounts with turnover untagged.
    if not res.pdf_link:
        res.status, res.note = "CHECK PDF", "the accounts filing has no downloadable document"
        return finish(res)
    r = web.get(res.pdf_link)
    if r is None:
        res.status, res.note = "CHECK PDF", "the accounts PDF couldn't be downloaded"
        return finish(res)
    pdf = read_pdf(r.content, max_pages, progress, read_scanned, kw, ocr_keywords, deadline)
    stopped = (f"{time_limit:g}-minute time limit" if pdf.timed_out else "page limit")
    if kw and not pdf.problem:
        res.keywords_found, res.keyword_counts = kw.result(), kw.ranked()
        if pdf.scanned_skipped:
            res.notes.append("keywords not searched in scanned pages (scanned-PDF reading is turned off)")
        elif pdf.keyword_ocr_skipped:
            res.notes.append("keywords not searched in scanned pages (keyword search of scanned PDFs is turned off)")
        elif pdf.keyword_pages < pdf.pages:
            res.notes.append(f"keywords searched in {pdf.keyword_pages} of {pdf.pages} pages ({stopped})")
    if pdf.turnover is not None:
        apply_pdf(res, pdf)
    elif reason:
        res.status, res.note = "NOT DISCLOSED", reason
    elif xhtml:
        res.status, res.note = "NOT DISCLOSED", "turnover not shown in the filed accounts"
    else:
        res.status = "CHECK PDF"
        if pdf.problem:
            res.note = pdf.problem
        elif pdf.timed_out:
            res.note = (f"stopped at the {stopped} after {pdf.pages_checked} of {pdf.pages} pages, "
                        "open the PDF to check")
        elif pdf.scanned_skipped:
            res.note = "scanned PDF not read (scanned-PDF reading is turned off)"
        else:
            res.note = f"no turnover line found in {pdf.pages_checked} pages, open the PDF to check"
    return finish(res)


def finish(res):
    if res.notes:
        res.note = "; ".join(filter(None, [res.note] + res.notes))
    return res


def safe_check(row, web, **kw):
    """check(), but a failure only ever affects this one row."""
    try:
        return check(row, web, **kw)
    except SourceError as e:
        return Result("ERROR", note=str(e))
    except Exception as e:                      # a bug or an unexpected page: report, carry on
        traceback.print_exc(file=sys.stderr)
        return Result("ERROR", note=f"unexpected problem ({type(e).__name__}: {e}), run this row again "
                                    "or check it by hand")


# ---------------------------------------------------------------- Excel

FILL_COLOURS = {"FOUND": "C6EFCE", "VERIFY": "FFEB9C", "NOT DISCLOSED": "FCE4D6",
                "CHECK PDF": "F8CBAD", "NOT FOUND": "FFC7CE", "ERROR": "D9D9D9"}
FILLS = {k: PatternFill("solid", fgColor=v) for k, v in FILL_COLOURS.items()}
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
COLS = [("Turnover Status", 16), ("Turnover", 17), ("Prior Year Turnover", 17), ("Currency", 9),
        ("Accounts Period End", 13), ("Notes", 50), ("Company Name (Companies House)", 34),
        ("Company Number", 12), ("Company Status", 13), ("Latest Accounts", 42), ("Filed On", 12),
        ("Read From", 30), ("Accounts PDF", 13), ("Companies House Page", 13)]
KEYWORD_COLS = [("Keyword Count", 10), ("Keywords Found", 60)]   # after Accounts Period End when searching


def columns(with_keywords=False):
    if not with_keywords:
        return COLS
    at = [h for h, _ in COLS].index("Accounts Period End") + 1
    return COLS[:at] + KEYWORD_COLS + COLS[at:]


def keyword_count(res):
    """How many different keywords were found; None when nothing was searched."""
    return None if res.keywords_found is None else len(res.keywords_found)


def keyword_cell(res):
    """'foreign exchange (14), hedging (3)', most mentioned first; None when nothing was searched."""
    if res.keywords_found is None:
        return None
    if not res.keywords_found:
        return "none found"
    counts = res.keyword_counts or {w: 1 for w in res.keywords_found}
    return ", ".join(f"{w} ({n})" for w, n in counts.items())
MONEY = '#,##0;(#,##0)'
MEANINGS = {
    "FOUND": "Read exactly from the filed accounts.",
    "VERIFY": "Read from a scanned PDF with OCR. Check the figure against the PDF (page given).",
    "NOT DISCLOSED": "The company didn't file a turnover figure. The Notes column says why.",
    "CHECK PDF": "Couldn't read a figure automatically. Open the linked PDF.",
    "NOT FOUND": "Couldn't match the company on the register. Add or check the company number.",
    "ERROR": "Companies House couldn't be reached for this row. Run it again later.",
}


def find_col(ws, header, header_row=1):
    if not header:
        return None
    for cell in ws[header_row]:
        if cell.value is not None and str(cell.value).strip().lower() == header.strip().lower():
            return cell.column
    found = [str(c.value) for c in ws[header_row] if c.value is not None]
    raise KeyError(f"Column '{header}' not found in row {header_row}. Headers there are: {found}")


def write_row(ws, r, start, res, cols=COLS):
    values = {"Turnover Status": res.status, "Turnover": res.turnover, "Prior Year Turnover": res.prior_turnover,
              "Currency": res.currency, "Accounts Period End": res.period_end, "Keyword Count": keyword_count(res),
              "Keywords Found": keyword_cell(res),
              "Notes": res.note, "Company Name (Companies House)": res.company_name,
              "Company Number": res.company_number, "Company Status": res.company_status,
              "Latest Accounts": res.accounts_type, "Filed On": res.filed_on, "Read From": res.method,
              "Accounts PDF": "Open PDF" if res.pdf_link else None,
              "Companies House Page": "Open page" if res.company_number else None}
    links = {"Accounts PDF": res.pdf_link, "Companies House Page": res.company_link}
    for i, (h, _w) in enumerate(cols):
        v = values[h]
        cell = ws.cell(row=r, column=start + i, value=v if v != "" else None)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if h in ("Turnover", "Prior Year Turnover") and v is not None:
            cell.number_format = MONEY
        if links.get(h):
            cell.hyperlink, cell.font = links[h], Font(color="0563C1", underline="single")
    ws.cell(row=r, column=start).fill = FILLS.get(res.status, FILLS["ERROR"])


def process_sheet(ws, cols, web=None, max_pages=None, on_row=None, on_detail=None,
                  checkpoint=None, header_row=1, values_ws=None, read_scanned=True, keywords=None,
                  ocr_keywords=True, time_limit=0):
    """Fill in the turnover columns for every row below header_row. cols maps 'name' and
    'number' to column indexes. values_ws, if given, is the same sheet loaded with cached
    formula results, used for reading (ws keeps the formulas for writing back). keywords, if
    given, are searched for in each company's accounts and listed in a Keywords Found column."""
    web = web or Web()
    src = values_ws or ws
    out_cols = columns(bool(keywords))
    headers = {str(c.value): c.column for c in ws[header_row] if c.value}
    start = headers.get(COLS[0][0], ws.max_column + 1)
    for i, (h, w) in enumerate(out_cols):
        cell = ws.cell(row=header_row, column=start + i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cell.column_letter].width = w

    cache, counts = {}, {}
    rows = [r for r in range(header_row + 1, src.max_row + 1)
            if any(src.cell(row=r, column=c).value not in (None, "") for c in cols.values() if c)]
    for done, r in enumerate(rows, 1):
        row = {k: src.cell(row=r, column=c).value for k, c in cols.items() if c}
        key = (str(row.get("name") or "").strip().lower(), clean_number(row.get("number")))
        if key not in cache:
            label = str(row.get("name") or row.get("number"))
            detail = (lambda msg, label=label: on_detail(f"{label}: {msg}")) if on_detail else None
            cache[key] = safe_check(row, web, max_pages=max_pages, progress=detail, read_scanned=read_scanned,
                                    keywords=keywords, ocr_keywords=ocr_keywords, time_limit=time_limit)
        res = cache[key]
        write_row(ws, r, start, res, out_cols)
        counts[res.status] = counts.get(res.status, 0) + 1
        if res.keywords_found:
            counts["_with_keywords"] = counts.get("_with_keywords", 0) + 1
        if on_row:
            on_row(done, len(rows), str(row.get("name") or row.get("number")), res)
        if checkpoint and done % 10 == 0:
            checkpoint()

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    last_col = get_column_letter(start + len(out_cols) - 1)
    ws.auto_filter.ref = f"A{header_row}:{last_col}{max(ws.max_row, header_row + 1)}"
    return counts


def add_summary(wb, counts, source_name="", keywords=None):
    """A 'Turnover Summary' sheet: counts per status, what each status means, when it ran,
    and which keywords were searched for."""
    title = "Turnover Summary"
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 80
    ws["A1"] = "Automation Tool results"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Run on {datetime.now():%d %B %Y at %H:%M}" + (f" for {source_name}" if source_name else "")
    ws["A3"] = "Source: Companies House (find-and-update.company-information.service.gov.uk)"
    for i, h in enumerate(["Status", "Companies", "Meaning"]):
        c = ws.cell(row=5, column=i + 1, value=h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), HEADER_FILL
    for r, status in enumerate(STATUSES, 6):
        ws.cell(row=r, column=1, value=status).fill = FILLS[status]
        ws.cell(row=r, column=2, value=counts.get(status, 0))
        ws.cell(row=r, column=3, value=MEANINGS[status])
    total = len(STATUSES) + 6
    ws.cell(row=total, column=1, value="Total").font = Font(bold=True)
    ws.cell(row=total, column=2, value=sum(v for k, v in counts.items() if k in STATUSES)).font = Font(bold=True)
    if keywords:
        r = total + 2
        ws.cell(row=r, column=1, value="Keywords").font = Font(bold=True)
        ws.cell(row=r, column=2, value=counts.get("_with_keywords", 0))
        ws.cell(row=r, column=3, value="companies whose accounts mention at least one keyword")
        ws.cell(row=r + 1, column=3, value=f"Searched for {len(keywords)}: " + ", ".join(keywords))
        ws.cell(row=r + 1, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("--name-col", help="column with company names")
    ap.add_argument("--number-col", help="column with company numbers")
    ap.add_argument("--sheet")
    ap.add_argument("--header-row", type=int, default=1, help="row the column headers are in")
    ap.add_argument("--max-pages", type=int, default=0, help="pages read per PDF (0 = no limit)")
    ap.add_argument("--time-limit", type=float, default=0,
                    help="minutes allowed per company for reading its PDF (0 = no limit)")
    ap.add_argument("--no-ocr", action="store_true", help="don't read scanned PDFs (much faster)")
    ap.add_argument("--keywords", help="words or phrases to look for in the accounts, comma-separated")
    ap.add_argument("--keywords-file", help="text file of keywords, one per line")
    ap.add_argument("--no-keyword-ocr", action="store_true",
                    help="don't search scanned PDFs for keywords (they're read page by page, which is slow)")
    ap.add_argument("--output")
    args = ap.parse_args()
    if not (args.name_col or args.number_col):
        sys.exit("Give --name-col, --number-col or both.")

    keywords = parse_keywords(args.keywords) if args.keywords else []
    if args.keywords_file:
        with open(args.keywords_file, encoding="utf-8") as f:
            keywords += parse_keywords(f.read())

    out_path = args.output or re.sub(r"\.xls[xm]?$", "", args.input, flags=re.I) + "_turnover.xlsx"
    wb = load_workbook(args.input)
    values = load_workbook(args.input, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.worksheets[0]
    try:
        cols = {"name": find_col(ws, args.name_col, args.header_row),
                "number": find_col(ws, args.number_col, args.header_row)}
    except KeyError as e:
        sys.exit(e.args[0])

    def show(i, n, name, res):
        amount = f"{res.turnover:,.0f}" if res.turnover is not None else ""
        kws = f"  [{len(res.keywords_found)} keywords]" if res.keywords_found else ""
        print(f"[{i}/{n}] {name[:40]:40} {res.status:13} {amount:>16}{kws}  {res.note[:60]}", flush=True)

    counts = process_sheet(ws, cols, time_limit=args.time_limit, max_pages=args.max_pages or None, on_row=show,
                           header_row=args.header_row,
                           values_ws=values[ws.title], read_scanned=not args.no_ocr, keywords=keywords or None,
                           ocr_keywords=not args.no_keyword_ocr,
                           checkpoint=lambda: wb.save(out_path))
    add_summary(wb, counts, args.input, keywords)
    wb.save(out_path)
    print("\n" + ", ".join(f"{k}: {v}" for k, v in sorted(counts.items()) if k in STATUSES))
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()
