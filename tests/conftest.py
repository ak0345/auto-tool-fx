import sys
from pathlib import Path

import pytest
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).parent / "fixtures"
sys.path.insert(0, str(ROOT))

import turnover  # noqa: E402


def fixture(name, mode="r"):
    path = FIXTURES / name
    return path.read_bytes() if mode == "rb" else path.read_text(encoding="utf-8")


class Response:
    def __init__(self, body):
        self.content = body if isinstance(body, bytes) else body.encode("utf-8")
        self.text = self.content.decode("utf-8", errors="replace")


class FakeWeb:
    """Stands in for turnover.Web. routes maps a path (with '?page=N' for filing history
    pages, or a full URL for documents) to: a str/bytes body, None (404), or an Exception
    to raise. Anything unrouted is a 404. Every request is recorded in .calls."""

    def __init__(self, routes):
        self.routes = routes
        self.calls = []

    def get(self, path, params=None):
        key = path
        if params and "page" in params:
            key = f"{path}?page={params['page']}"
        elif params and "q" in params:
            key = f"{path}?q={params['q']}"
        self.calls.append(key)
        body = self.routes.get(key)
        if isinstance(body, Exception):
            raise body
        return Response(body) if body is not None else None

    def soup(self, path, params=None):
        r = self.get(path, params)
        return BeautifulSoup(r.text, "html.parser") if r is not None else None


def filing_page(*rows, next_page=False):
    """A filing-history page in the same shape as the real site. rows are
    (date, type, description, has_pdf, has_xhtml)."""
    trs = []
    for i, (when, kind, desc, pdf, xhtml) in enumerate(rows):
        links = ""
        if pdf:
            links += f'<a href="/company/X/filing-history/T{i}/document?format=pdf&amp;download=0">View PDF</a>'
        if xhtml:
            links += f'<a href="/company/X/filing-history/T{i}/document?format=xhtml&amp;download=1">Download iXBRL</a>'
        trs.append(f"<tr><td>{when}</td><td>{kind}</td><td>{desc} View PDF {desc}</td><td>{links}</td></tr>")
    pager = '<a href="/company/X/filing-history?page=2">Next</a>' if next_page else ""
    return f'<table id="fhTable"><tr><th>Date</th><th>Type</th><th>Description</th><th>View</th></tr>{"".join(trs)}</table>{pager}'


PDF0 = turnover.BASE + "/company/X/filing-history/T0/document?format=pdf&download=0"
XHTML0 = turnover.BASE + "/company/X/filing-history/T0/document?format=xhtml&download=1"


def company(number, profile="profile_active.html", filings=None, xhtml=None, pdf=None, extra=None):
    """Routes for one company: its page, filing history page 1, and documents."""
    routes = {f"/company/{number}": fixture(profile)}
    if filings is not None:
        routes[f"/company/{number}/filing-history?page=1"] = filings
    if xhtml is not None:
        routes[XHTML0] = xhtml
    if pdf is not None:
        routes[PDF0] = pdf
    routes.update(extra or {})
    return routes


needs_ocr = pytest.mark.skipif(not turnover.ocr_available(), reason="scanned-PDF reader not installed")


@pytest.fixture
def no_sleep(monkeypatch):
    monkeypatch.setattr(turnover.time, "sleep", lambda s: None)
