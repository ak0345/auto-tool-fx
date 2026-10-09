"""Reading uploads and writing results back into the user's workbook."""

import io

from openpyxl import Workbook, load_workbook

import turnover as t
from conftest import FakeWeb, company, filing_page, fixture
from inputs import NONE, guess, guess_header_row, read_upload

MICRO = filing_page(("29 Jul 2026", "AA", "Micro company accounts made up to 31 October 2025", True, True))
TAGGED = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True))


def xlsx_bytes(rows, title_rows=()):
    wb = Workbook()
    for r in list(title_rows) + list(rows):
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- reading uploads


def test_read_xlsx_keeps_formulas_for_writing_and_values_for_reading():
    wb = Workbook()
    wb.active.append(["Company Name", "Company No", "Joined"])
    wb.active.append(["Tesco PLC", "00445790", "=1+1"])
    buf = io.BytesIO()
    wb.save(buf)
    write_wb, read_wb = read_upload(buf.getvalue(), "clients.xlsx")
    assert write_wb.active["C2"].value == "=1+1"
    assert read_wb.active["A2"].value == "Tesco PLC"


def test_read_csv_with_semicolons_and_windows_encoding():
    data = "Company Name;Company No\nCafé Nero Group Ltd;05000000\n".encode("cp1252")
    wb, _ = read_upload(data, "clients.csv")
    assert [c.value for c in wb.active[2]] == ["Café Nero Group Ltd", "05000000"]


def test_read_csv_with_utf8_bom_and_commas():
    data = "﻿Company Name,Company No\nTesco PLC,00445790\n".encode("utf-8")
    wb, _ = read_upload(data, "clients.csv")
    assert wb.active["A1"].value == "Company Name" and wb.active["B2"].value == "00445790"


def test_read_old_xls():
    wb, _ = read_upload(fixture("clients_old_excel.xls", "rb"), "clients.xls")
    ws = wb["Clients"]
    assert ws["A1"].value == "Company Name" and ws["A2"].value == "Tesco PLC" and ws["B2"].value == 445790


def test_header_row_detection():
    wb = Workbook()
    for r in [["Client list, October 2026"], [], ["Company Name", "Company Number"], ["Tesco PLC", "00445790"]]:
        wb.active.append(r)
    assert guess_header_row(wb.active) == 3
    wb2 = Workbook()
    wb2.active.append(["Client", "Ref"])
    assert guess_header_row(wb2.active) == 1


def test_column_guessing():
    headers = ["Ref", "Company Name", "Company Reg No", "Address"]
    num = guess(headers, ["company no", "company number", "crn", "reg no", "registration", "number"])
    assert num == "Company Reg No"
    assert guess(headers, ["company name", "name"], skip=num) == "Company Name"
    assert guess(["Address"], ["name"]) == NONE


# ---------------------------------------------------------------- writing results

EXPECTED_HEADERS = ["Company Name", "Company Number", "Turnover Status", 'Turnover or "Revenue"',
                    'Prior Year "Turnover" or "Revenue"', "Cash", "Debtors", "Persons with significant control",
                    "linkedin search", "Keyword Count", "Keywords Found", "Notes", "Latest Accounts",
                    "Accounts PDF", "Companies House Page"]


def test_output_has_exactly_the_requested_columns():
    _wb, out = t.new_output()
    assert [c.value for c in out[1]] == EXPECTED_HEADERS
    assert out["A1"].fill.fgColor.rgb.endswith("7030A0") and out.freeze_panes == "A2"


def test_process_sheet_end_to_end():
    data = xlsx_bytes([["Company Name", "Company No", "Owner"],
                       ["Tesco PLC", 445790.0, "=\"A\"&\"B\""],      # float number, formula
                       [None, None, "just a note"],                   # blank company: skipped
                       ["Tesco PLC", "00445790", "dupe"],             # duplicate: looked up once
                       ["Nobody Ltd", "99999999", ""]],               # unknown number
                      title_rows=[["Client list"], []])
    _wb, values = read_upload(data, "c.xlsx")
    src = values.active
    routes = company("00445790", filings=TAGGED, xhtml=fixture("ixbrl_tagged.html", "rb"))
    web = FakeWeb(routes)
    seen = []
    cols = {"name": t.find_col(src, "Company Name", 3), "number": t.find_col(src, "Company No", 3)}
    out_wb, out = t.new_output()
    counts = t.process_sheet(src, cols, out, web=web, header_row=3,
                             on_row=lambda i, n, name, res: seen.append((i, n, res.status)))
    assert counts == {"FOUND": 2, "NOT FOUND": 1}
    assert [s[:2] for s in seen] == [(1, 3), (2, 3), (3, 3)]
    assert web.calls.count("/company/00445790") == 1                  # cached

    col = {h: i + 1 for i, h in enumerate(EXPECTED_HEADERS)}
    assert out.max_row == 4                                            # one row per company, no blanks
    assert out.cell(2, col["Company Name"]).value == "TESCO PLC"       # the registered name
    assert out.cell(2, col["Company Number"]).value == "00445790"
    assert out.cell(2, col['Turnover or "Revenue"']).value == 44_043_650
    assert out.cell(2, col['Prior Year "Turnover" or "Revenue"']).value == 47_988_261
    assert out.cell(2, col['Turnover or "Revenue"']).number_format == t.MONEY
    assert out.cell(2, col["Accounts PDF"]).hyperlink.target.endswith("format=pdf&download=0")
    assert out.cell(2, col["Companies House Page"]).hyperlink.target.endswith("/company/00445790")
    assert out.cell(2, col["Turnover Status"]).fill.fgColor.rgb.endswith(t.FILL_COLOURS["FOUND"])
    assert out.cell(4, col["Company Name"]).value == "Nobody Ltd"      # not found: what was typed
    assert out.cell(4, col["Company Number"]).value == "99999999"
    assert out.cell(4, col["Turnover Status"]).value == "NOT FOUND"
    assert out.auto_filter.ref == "A1:O4"

    t.add_summary(out_wb, counts, "c.xlsx")
    summary = out_wb["Summary"]
    rows = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(6, 14)}
    assert rows["FOUND"] == 2 and rows["NOT FOUND"] == 1 and rows["Total"] == 3

    buf = io.BytesIO()
    out_wb.save(buf)                                                   # saves and reopens cleanly
    assert load_workbook(io.BytesIO(buf.getvalue()))["Summary"]["A1"].value == "Automation Tool results"


def test_error_rows_dont_stop_the_run():
    _wb, values = read_upload(xlsx_bytes([["Company No"], ["00000001"], ["00445790"]]), "c.xlsx")
    routes = company("00445790", filings=MICRO, xhtml=fixture("ixbrl_micro.html", "rb"))
    routes["/company/00000001"] = RuntimeError("boom")
    _out_wb, out = t.new_output()
    counts = t.process_sheet(values.active, {"name": None, "number": 1}, out, web=FakeWeb(routes))
    assert counts == {"ERROR": 1, "NOT DISCLOSED": 1}


def test_checkpoint_called_every_ten_rows():
    _wb, values = read_upload(xlsx_bytes([["Company No"]] + [["99999999"]] * 25), "c.xlsx")
    calls = []
    _out_wb, out = t.new_output()
    t.process_sheet(values.active, {"name": None, "number": 1}, out, web=FakeWeb({}), checkpoint=lambda: calls.append(1))
    assert len(calls) == 2


def test_find_col_reports_available_headers():
    wb = Workbook()
    wb.active.append(["Name", "Number"])
    try:
        t.find_col(wb.active, "Company")
    except KeyError as e:
        assert "Name" in e.args[0] and "Number" in e.args[0]
