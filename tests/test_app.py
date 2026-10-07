"""The web app, driven through Streamlit's test harness with Companies House faked."""

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

import turnover as t
from conftest import FakeWeb, company, filing_page, fixture

APP = str(Path(__file__).resolve().parent.parent / "app.py")
CSV = "text/csv"


@pytest.fixture
def fake_ch(monkeypatch):
    routes = company("00445790", filings=filing_page(
        ("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True)),
        xhtml=fixture("ixbrl_tagged.html", "rb"))
    monkeypatch.setattr(t, "Web", lambda *a, **k: FakeWeb(routes))


def start():
    return AppTest.from_file(APP, default_timeout=60).run()


def upload(at, name, data, mime=CSV):
    at.file_uploader[0].upload(name, data, mime).run()
    return at


def test_landing_page_renders():
    at = start()
    assert not at.exception
    assert any("Automation</span> Tool</h1>" in m.value for m in at.markdown)
    assert not at.button                                      # nothing to run until a file is uploaded


def test_upload_guess_columns_run_and_get_results(fake_ch):
    data = b"Client list\n\nCompany Name,Company Number\nTesco PLC,445790\n,\nNobody Ltd,99999999\n"
    at = upload(start(), "clients.csv", data)
    assert not at.exception
    assert [n for n in at.number_input if "Headers" in n.label][0].value == 3   # found below the title
    assert at.selectbox[0].value == "Company Name" and at.selectbox[1].value == "Company Number"
    assert any("2 companies" in c.value for c in at.caption)

    at.button[0].click().run()
    assert not at.exception
    res = at.session_state["result"]
    assert res["final"] and res["done"] == 2 and res["name"] == "clients_turnover.xlsx"
    statuses = [r[1] for r in res["rows"]]
    assert statuses == ["FOUND", "NOT FOUND"]
    assert res["rows"][0][2] == "£44,043,650"
    assert res["bytes"][:2] == b"PK"                          # a real xlsx (zip) file


def test_unreadable_file_shows_an_error():
    at = upload(start(), "broken.xlsx", b"definitely not excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    assert any("Couldn't open that file" in e.value for e in at.error)


def test_sheet_without_company_columns_asks_for_them():
    at = upload(start(), "other.csv", b"Fruit,Colour\nApple,Red\n")
    assert any("Pick a company name column" in w.value for w in at.warning)
    assert not at.button


def test_empty_header_row_is_explained():
    at = upload(start(), "gap.csv", b"Company Name\n\nTesco PLC\n")
    [n for n in at.number_input if "Headers" in n.label][0].set_value(2).run()
    assert not at.exception and any("Row 2 is empty" in e.value for e in at.error)
    [n for n in at.number_input if "Headers" in n.label][0].set_value(1).run()
    assert at.selectbox[0].value == "Company Name"


def test_keyword_box_starts_with_the_default_list():
    at = start()
    box = at.text_area[0]
    assert box.value.splitlines()[0] == "foreign exchange" and len(box.value.splitlines()) == 162
    assert any("162 keywords" in c.value for c in at.caption)


def test_run_with_keywords_lists_them_per_company(fake_ch):
    at = upload(start(), "clients.csv", b"Company Number\n445790\n")
    assert any("searching for 162 keywords" in c.value for c in at.caption)
    at.button[0].click().run()
    row = at.session_state["result"]["rows"][0]
    assert row[1] == "FOUND" and row[4].startswith("derivative (1), presentation currency (1)") and row[6] == 3


def test_clearing_the_box_skips_the_search(fake_ch):
    at = start()
    at.text_area[0].set_value("").run()
    assert any("only turnover will be looked up" in c.value for c in at.caption)
    upload(at, "clients.csv", b"Company Number\n445790\n")
    at.button[0].click().run()
    assert at.session_state["result"]["rows"][0][4] == ""


def test_long_keyword_search_is_warned_about():
    at = upload(start(), "clients.csv", b"Company Number\n445790\n")
    assert any("every page is read to search for the keywords" in c.value for c in at.caption)
    at.toggle[1].set_value(False).run()                          # keyword search of scanned PDFs off
    assert not any("every page is read" in c.value for c in at.caption)


def test_time_limit_box_defaults_to_no_limit_and_shows_in_estimate():
    at = upload(start(), "clients.csv", b"Company Number\n445790\n")
    box = [n for n in at.number_input if "Time limit" in n.label][0]
    assert box.value == 0
    box.set_value(10).run()
    assert any("up to 10 minutes for each large company" in c.value for c in at.caption)


def test_page_slider_defaults_to_no_limit():
    at = start()
    assert at.slider[0].label.startswith("Max pages per PDF") and at.slider[0].value == 0
