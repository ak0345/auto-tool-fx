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


def test_process_sheet_end_to_end():
    data = xlsx_bytes([["Company Name", "Company No", "Owner"],
                       ["Tesco PLC", 445790.0, "=\"A\"&\"B\""],        # float number, formula
                       [None, None, "just a note"],                   # blank company: skipped
                       ["Tesco PLC", "00445790", "dupe"],             # duplicate: looked up once
                       ["", "99999999", ""]],                         # unknown number
                      title_rows=[["Client list"], []])
    wb, values = read_upload(data, "c.xlsx")
    routes = company("00445790", filings=TAGGED, xhtml=fixture("ixbrl_tagged.html", "rb"))
    web = FakeWeb(routes)
    seen = []
    cols = {"name": t.find_col(values.active, "Company Name", 3), "number": t.find_col(values.active, "Company No", 3)}
    counts = t.process_sheet(wb.active, cols, web=web, header_row=3, values_ws=values.active,
                             on_row=lambda i, n, name, res: seen.append((i, n, res.status)))
    assert counts == {"FOUND": 2, "NOT FOUND": 1}
    assert [s[:2] for s in seen] == [(1, 3), (2, 3), (3, 3)]
    assert web.calls.count("/company/00445790") == 1                  # cached

    ws = wb.active
    heads = [c.value for c in ws[3]]
    assert heads[:3] == ["Company Name", "Company No", "Owner"] and heads[3] == "Turnover Status"
    col = {h: i + 1 for i, h in enumerate(heads)}
    assert ws.cell(4, col["Turnover"]).value == 44_043_650
    assert ws.cell(4, col["Turnover"]).number_format == t.MONEY
    assert ws.cell(4, col["Accounts PDF"]).hyperlink.target.endswith("format=pdf&download=0")
    assert ws.cell(4, col["Companies House Page"]).hyperlink.target.endswith("/company/00445790")
    assert ws.cell(4, col["Turnover Status"]).fill.fgColor.rgb.endswith(t.FILL_COLOURS["FOUND"])
    assert ws.cell(5, col["Turnover Status"]).value is None                # blank row untouched
    assert ws.cell(7, col["Turnover Status"]).value == "NOT FOUND"
    assert ws["C4"].value == '="A"&"B"'                                    # formula kept
    assert ws.freeze_panes == "A4" and ws.auto_filter.ref.startswith("A3:")

    t.add_summary(wb, counts, "c.xlsx")
    summary = wb["Turnover Summary"]
    rows = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(6, 14)}
    assert rows["FOUND"] == 2 and rows["NOT FOUND"] == 1 and rows["Total"] == 3

    buf = io.BytesIO()
    wb.save(buf)                                                           # saves and reopens cleanly
    assert load_workbook(io.BytesIO(buf.getvalue()))["Turnover Summary"]["A1"].value == "Automation Tool results"


def test_running_twice_reuses_the_result_columns():
    data = xlsx_bytes([["Company No"], ["00445790"]])
    wb, values = read_upload(data, "c.xlsx")
    routes = company("00445790", filings=MICRO, xhtml=fixture("ixbrl_micro.html", "rb"))
    for _ in range(2):
        t.process_sheet(wb.active, {"name": None, "number": 1}, web=FakeWeb(routes), values_ws=values.active)
        t.add_summary(wb, {"NOT DISCLOSED": 1})
    heads = [c.value for c in wb.active[1]]
    assert heads.count("Turnover Status") == 1 and wb.sheetnames.count("Turnover Summary") == 1


def test_error_rows_dont_stop_the_run():
    data = xlsx_bytes([["Company No"], ["00000001"], ["00445790"]])
    wb, values = read_upload(data, "c.xlsx")
    routes = company("00445790", filings=MICRO, xhtml=fixture("ixbrl_micro.html", "rb"))
    routes["/company/00000001"] = RuntimeError("boom")
    counts = t.process_sheet(wb.active, {"name": None, "number": 1}, web=FakeWeb(routes), values_ws=values.active)
    assert counts == {"ERROR": 1, "NOT DISCLOSED": 1}


def test_checkpoint_called_every_ten_rows():
    data = xlsx_bytes([["Company No"]] + [["99999999"]] * 25)
    wb, values = read_upload(data, "c.xlsx")
    calls = []
    t.process_sheet(wb.active, {"name": None, "number": 1}, web=FakeWeb({}), values_ws=values.active,
                    checkpoint=lambda: calls.append(1))
    assert len(calls) == 2


def test_find_col_reports_available_headers():
    wb = Workbook()
    wb.active.append(["Name", "Number"])
    try:
        t.find_col(wb.active, "Company")
    except KeyError as e:
        assert "Name" in e.args[0] and "Number" in e.args[0]
