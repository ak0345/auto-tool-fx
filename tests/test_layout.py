"""The fields added for the requested output: cash, debtors, prior year, PSCs and the LinkedIn
search link. The two real companies are the ones in the example they sent."""

from urllib.parse import parse_qs, urlparse

import pytest
from bs4 import BeautifulSoup

import turnover as t
from conftest import FakeWeb, company, filing_page, fixture, text_pdf


def soup(name):
    return BeautifulSoup(fixture(name, "rb"), "html.parser")


# ---------------------------------------------------------------- the example companies


def test_cargo_overseas_tagged_figures():
    doc = soup("ixbrl_cash_debtors_tagged.html")
    assert t.read_ixbrl(doc)[:2] == (56_557_394, 57_040_440)
    # the total, not the "due within one year" breakdown (8,569,709) tagged with a dimension
    assert t.read_ixbrl_balance(doc) == (3_535_079, 9_212_197)


def test_eurosonic_turnover_from_positioned_text():
    doc = soup("ixbrl_positioned_text.html")
    turnover = t.read_ixbrl(doc)
    assert turnover[:2] == (43_954_299, 48_827_556) and turnover[4] == "iXBRL table (untagged)"
    assert t.read_ixbrl_balance(doc) == (956_080, 6_667_617)


def test_positioned_rows_are_rebuilt_left_to_right():
    lines = t.ix_lines(soup("ixbrl_positioned_text.html"))
    assert "Turnover 3 43,954,299 48,827,556" in lines
    assert "Cash at bank and in hand 956,080 731,812" in lines


def test_psc_pages():
    corporate = FakeWeb({"/company/X/persons-with-significant-control": fixture("psc_corporate.html")})
    assert t.read_pscs(corporate, "X") == [("Venture Asset Management Limited", True)]
    person = FakeWeb({"/company/X/persons-with-significant-control": fixture("psc_individual.html")})
    assert t.read_pscs(person, "X") == [("Mr Harpreet Singh Chadha", False)]


def test_ceased_pscs_are_left_out():
    page = ('<div class="appointment-1"><span id="psc-name-1"><b>Old Owner</b></span>'
            '<span id="psc-status-tag-1">Ceased</span></div>'
            '<div class="appointment-2"><span id="psc-name-2"><b>Mrs New Owner</b></span>'
            '<span id="psc-status-tag-2">Active</span></div>')
    web = FakeWeb({"/company/X/persons-with-significant-control": page})
    assert t.read_pscs(web, "X") == [("Mrs New Owner", False)]
    assert t.read_pscs(FakeWeb({}), "X") == []                       # no PSC page


# ---------------------------------------------------------------- LinkedIn search


def query(url):
    return parse_qs(urlparse(url).query)["q"][0]


def test_linkedin_search_uses_company_and_first_person():
    res = t.Result(company_name="EUROSONIC GROUP LIMITED", pscs=[("Mr Harpreet Singh Chadha", False)])
    text, url = t.linkedin_search(res)
    assert text == "EUROSONIC GROUP + Mr Harpreet Singh Chadha + LinkedIn"
    assert url.startswith("https://www.google.com/search?q=")
    assert query(url) == "EUROSONIC GROUP Harpreet Singh Chadha LinkedIn"     # title dropped from the search


def test_linkedin_search_prefers_a_person_over_a_company():
    res = t.Result(company_name="ACME PLC", pscs=[("Holdco Limited", True), ("Dr Jane Smith", False)])
    assert query(t.linkedin_search(res)[1]) == "ACME Jane Smith LinkedIn"
    corporate_only = t.Result(company_name="CARGO OVERSEAS LIMITED", pscs=[("Venture Asset Management Limited", True)])
    assert t.linkedin_search(corporate_only)[0] == "CARGO OVERSEAS + Venture Asset Management Limited + LinkedIn"


def test_linkedin_search_without_pscs_or_company():
    assert query(t.linkedin_search(t.Result(company_name="TESCO PLC", pscs=[]))[1]) == "TESCO LinkedIn"
    assert t.linkedin_search(t.Result()) == (None, None)


def test_psc_cell():
    assert t.psc_cell(t.Result(pscs=[("A Ltd", True), ("Mr B", False)])) == "A Ltd; Mr B"
    assert t.psc_cell(t.Result(pscs=[])) == "none registered"
    assert t.psc_cell(t.Result()) is None


# ---------------------------------------------------------------- reading rows


@pytest.mark.parametrize("line, value", [
    ("Cash at bank and in hand 956,080 731,812", 956_080),
    ("Cashatbankandinhand 45,300 30,200", 45_300),                 # OCR ran the words together
    ("Cash and cash equivalents 18 2,345 1,999", 2_345),
    ("Cash 12,000 9,000", 12_000),
])
def test_cash_rows(line, value):
    assert t.figure_from_lines([line], t.CASH_LABEL)[0] == value


@pytest.mark.parametrize("line", [
    "Cash flow from operating activities 1,234 1,000",
    "Cash and cash equivalents at end of year 956,080 731,812",     # cash flow statement, not the balance
])
def test_rows_that_arent_balance_sheet_cash(line):
    assert t.figure_from_lines([line], t.CASH_LABEL) is None


@pytest.mark.parametrize("line, value", [
    ("Debtors 17 6,667,617 8,767,425", 6_667_617),
    ("Trade and other receivables 1,234 1,100", 1_234),
    ("Total debtors 6,667,617 8,767,425", 6_667_617),
])
def test_debtor_rows(line, value):
    assert t.figure_from_lines([line], t.DEBTORS_LABEL)[0] == value


def test_debtors_heading_without_figures_is_skipped():
    rows = ["Debtors: amounts falling due within one year", "Trade debtors 900 800"]
    assert t.figure_from_lines(rows, t.DEBTORS_LABEL) is None


@pytest.mark.parametrize("line, prior", [
    ("Turnover 3 43,954,299 48,827,556", 48_827_556),               # this year, last year
    ("Revenue 2, 3 17,273.6 13,816.8", 13_816.8),
    ("Revenue 2,3 73,712 73,712 69,916 69,916", None),              # extra columns: can't tell which
    ("Turnover   1 5,465,729 11,9 47,210", None),                   # too far apart to be last year
    ("Sales 812,400", None),
])
def test_prior_year_only_when_unambiguous(line, prior):
    assert t.turnover_from_lines([line])[1] == prior


def test_sales_counts_as_turnover_but_cost_of_sales_does_not():
    assert t.turnover_from_lines(["Sales 3 812,400 790,100"])[0] == 812_400
    assert t.turnover_from_lines(["Net sales 1,000 900"])[0] == 1_000
    assert t.turnover_from_lines(["Cost of sales (800,000) (700,000)"]) is None


# ---------------------------------------------------------------- PDFs and whole companies

SMALL_PDF = text_pdf(["Example Ltd", "Profit and Loss Account", "Note 2025 2024", "£ £", "Sales 3 812,400 790,100"],
                     ["Example Ltd", "Balance Sheet", "Note 2025 2024", "£ £", "Current assets",
                      "Debtors 5 120,500 99,000", "Cash at bank and in hand 45,300 30,200"])


def test_text_pdf_gives_turnover_prior_cash_and_debtors():
    r = t.read_pdf(SMALL_PDF)
    assert (r.turnover, r.prior, r.cash, r.debtors) == (812_400, 790_100, 45_300, 120_500)


def test_balance_sheet_in_thousands():
    pdf = text_pdf(["Group Balance Sheet", "2025 2024", "£000 £000", "Debtors 4,210 3,900", "Cash 1,050 990"])
    r = t.read_pdf(pdf)
    assert (r.cash, r.debtors) == (1_050_000, 4_210_000)


def test_balance_sheet_figures_from_notes_pages_are_ignored():
    pdf = text_pdf(["Notes to the financial statements (continued)", "Balance sheet items", "Cash 999,999"])
    assert t.read_pdf(pdf).cash is None


def test_whole_company_row_has_everything():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 30 June 2025", True, False))
    routes = company("01234567", profile="profile_active.html", filings=page, pdf=SMALL_PDF)
    routes["/company/01234567/persons-with-significant-control"] = fixture("psc_individual.html")
    res = t.safe_check({"number": "01234567"}, FakeWeb(routes))
    assert (res.status, res.turnover, res.prior_turnover, res.cash, res.debtors) == \
        ("FOUND", 812_400, 790_100, 45_300, 120_500)
    assert res.pscs == [("Mr Harpreet Singh Chadha", False)]


def test_small_company_still_gets_cash_and_debtors():
    page = filing_page(("01 Mar 2026", "AA", "Total exemption full accounts made up to 31 March 2025", True, False))
    pdf = text_pdf(["Small Ltd", "Balance Sheet", "2025 2024", "£ £", "Debtors 8,100 7,000", "Cash at bank 2,500 1,900"])
    res = t.safe_check({"number": "01234567"}, FakeWeb(company("01234567", filings=page, pdf=pdf)))
    assert res.status == "NOT DISCLOSED" and (res.cash, res.debtors) == (2_500, 8_100)


def test_psc_page_failure_doesnt_lose_the_row():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 30 June 2025", True, False))
    routes = company("01234567", filings=page, pdf=SMALL_PDF)
    routes["/company/01234567/persons-with-significant-control"] = t.SourceError("server error")
    res = t.safe_check({"number": "01234567"}, FakeWeb(routes))
    assert res.status == "FOUND" and res.pscs is None and "persons with significant control couldn't be read" in res.note


def test_non_sterling_accounts_are_noted_and_not_formatted_in_pounds():
    from test_documents import ixbrl
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c" unitRef="EUR">100</ix:nonFraction>',
                [("c", "2024-01-01", "2024-12-31", "")])
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2024", True, True))
    res = t.safe_check({"number": "01234567"}, FakeWeb(company("01234567", filings=page, xhtml=doc)))
    assert res.currency == "EUR" and "figures are in EUR" in res.note
    wb, ws = t.new_output()
    t.write_row(ws, 2, {}, res)
    assert ws.cell(2, 4).number_format == t.PLAIN_MONEY


def test_output_row_for_the_example_company():
    res = t.Result(status="FOUND", company_name="EUROSONIC GROUP LIMITED", company_number="01394171",
                   turnover=43_954_299, prior_turnover=48_827_556, cash=956_080, debtors=6_667_617,
                   pscs=[("Mr Harpreet Singh Chadha", False)], accounts_type="Group of companies' accounts",
                   pdf_link="https://example/pdf")
    _wb, ws = t.new_output()
    t.write_row(ws, 2, {}, res)
    row = {ws.cell(1, c).value: ws.cell(2, c) for c in range(1, 16)}
    assert row["Company Name"].value == "EUROSONIC GROUP LIMITED" and row["Company Number"].value == "01394171"
    assert [row[h].value for h in ('Turnover or "Revenue"', 'Prior Year "Turnover" or "Revenue"', "Cash", "Debtors")] \
        == [43_954_299, 48_827_556, 956_080, 6_667_617]
    assert row["Cash"].number_format == t.MONEY
    assert row["Persons with significant control"].value == "Mr Harpreet Singh Chadha"
    assert row["linkedin search"].value == "EUROSONIC GROUP + Mr Harpreet Singh Chadha + LinkedIn"
    assert row["linkedin search"].hyperlink.target.startswith("https://www.google.com/search?q=EUROSONIC")


@pytest.mark.parametrize("line", ["Debtors 4 0 17,740", "Debtors 4 - 17,740", "Debtors 4 nil 17,740"])
def test_zero_this_year_isnt_replaced_by_last_year(line):
    assert t.figure_from_lines([line], t.DEBTORS_LABEL)[0] == 0


def test_balance_sheet_units_from_a_heading_far_above():
    rows = ["Group balance sheet", "2025 2024", "Notes £m £m", "Non-current assets", "Goodwill 1,200 1,100"] + \
           [f"Other asset {i} 10 10" for i in range(12)] + ["Current assets", "Inventories 300 280",
                                                          "Cash and cash equivalents 71 125"]
    assert t.balance_figure(rows, t.CASH_LABEL)[0] == 71_000_000


def test_stray_m_in_a_sentence_is_not_a_heading():
    rows = ["We served 5 m customers this year"] + ["filler"] * 10 + ["Cash 1,050 990"]
    assert t.balance_figure(rows, t.CASH_LABEL)[0] == 1_050


def test_current_receivables_not_non_current():
    rows = ["Balance sheet", "£m £m", "Non-current assets", "Trade and other receivables 161 150",
            "Current assets", "Inventories 2,500 2,400", "Trade and other receivables 3,948 3,600", "Cash 2,515 2,300"]
    assert t.balance_figure(rows, t.DEBTORS_LABEL)[0] == 3_948_000_000
    assert t.balance_figure(rows, t.CASH_LABEL)[0] == 2_515_000_000


def test_tesco_two_column_balance_sheet():
    """OCR rows from Tesco's 2026 balance sheet, where two columns of the page share each row."""
    rows = ["Groupbalance sheet",
            "Notes 28 February 2026 Em 22 February 2025 m Notes 28 February2026 fm 22 February 2025 Em",
            "Non-current assets Non-current liabilities",
            "Trade and other receivables 18 161 158 Deferred tax liabilities (635) (503)",
            "30,991 30,034 Share premium 5,166 5,165",
            "Current assets Otherreserves 29 3,167 3,140",
            "Inventories 17 2,840 2,768 Equityattributable toowners of the parent 11,463 11,666",
            "Trade and other receivables 18 1,318 1,210 Non-controlling interests (6) (4)",
            "Short-term investments 19 1,429 2.223 The notes on pages 127 to 199 form part of these financial statements.",
            "Cash and cash equivalents 19 2,515 2.255"]
    assert t.balance_figure(rows, t.DEBTORS_LABEL)[0] == 1_318_000_000
    assert t.balance_figure(rows, t.CASH_LABEL)[0] == 2_515_000_000


def test_next_plc_wording():
    rows = ["CONSOLIDATEDBALANCESHEET", "Notes fm m", "Current assets",
            "Customerandotherreceivables 14 1,660.6 1,508.4", "Cashandshortterm deposits 16 96.2 200.4 FinancialStatements"]
    assert t.balance_figure(rows, t.DEBTORS_LABEL)[0] == 1_660_600_000
    assert t.balance_figure(rows, t.CASH_LABEL)[0] == 96_200_000


def test_unreadable_current_assets_give_blank_not_non_current():
    rows = ["Consolidated statement of financial position", "£m £m",
            "Trade and other receivables 17 279.0 382.8 356.7", "Current assets Share premium account 994.6 982.7",
            "Other financial assets Current tax assets Inventories Trade and other receivables"]
    assert t.balance_figure(rows, t.DEBTORS_LABEL) is None


def test_debtors_tagged_as_due_within_one_year():
    assert t.read_ixbrl_balance(soup("ixbrl_tagged.html"))[1] == 1_127_610


def test_small_company_debtors_row_wording():
    assert t.figure_from_lines(["Debtors: amounts falling due within one year 13 1,127,610 1,599,565"],
                               t.DEBTORS_LABEL)[0] == 1_127_610
