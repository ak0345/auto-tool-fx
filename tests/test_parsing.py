"""Reading amounts, units and headings out of lines of text (shared by iXBRL tables and PDFs).
The row texts are real ones from filings and OCR output that caused trouble during testing."""

import pytest

import turnover as t


@pytest.mark.parametrize("line, expected", [
    ("Turnover 3 30,035,247 26,242,640", 30035247),                              # note ref skipped
    ("Revenue 2,3 73,712 73,712 69,916 69,916", 73712),                         # Tesco, "2,3" notes
    ("Revenue 2, 3 17,273.6 13,816.8", 17273.6),                                # M&S, £m decimals
    ("Revenue 2,151.2 2,151.2 2,014.4 2,014.4", 2151.2),                        # Greggs
    ("Revenue(including credit account interest) 1,2 6,901.3 6,118.1", 6901.3),  # Next
    ("Revenue 1.771.0 1.7065 Prnfit fnr the nerind 156.3 151.2", 1771.0),       # Dunelm, OCR dots + merged column
    ("Turnover 2 50,503 52,479 51,680", 50503),                                 # Unilever
    ("Turnover 3 72.886 69,191", 72886),                                        # OCR read a comma as a dot
    ("Revenue 37.3 35.6", 37.3),                                                # small decimals, two of them
    ("Revenue 4.1 2,920.2", 2920.2),                                            # lone small decimal is a note
    ("Turnover 2.1 1,234,567 1,100,000", 1234567),
    ("Turnover £000 12,764", 12764),                                            # £000 is a heading, not zero
    ("Turnover   1 5,465,729 11,9 47,210", 5465729),                            # badly typeset prior year
    ("Revenue (note 3) 5,000", 5000),
    ("Revenue (3) 1,234 1,100", 1234),
    ("TURNOVER 3 10,156,284  9,248,588", 10156284),
    ("Group revenue 4,567,890", 4567890),
    ("Total turnover (1,234)", -1234),                                          # negative in brackets
])
def test_turnover_rows(line, expected):
    found = t.turnover_from_lines([line])
    assert found is not None and found[0] == expected


@pytest.mark.parametrize("line", [
    "Revenue from sale of goods and services 72,886 72,886",   # a sub-line, not the total
    "Turnover is measured at the fair value of the consideration received",  # policy text
    "Turnover",                                                # heading with no figures
    "Cost of sales (800.0) (700.0)",
    "Revenue 2.1",                                             # just a note number
])
def test_rows_that_are_not_turnover(line):
    assert t.turnover_from_lines([line]) is None


def test_first_matching_row_wins_and_reports_its_index():
    found = t.turnover_from_lines(["Notes £m £m", "Revenue from sale of goods 72,886",
                                   "Revenue 2,3 73,712 69,916"])
    assert found[0] == 73712 and found[3] == 2


@pytest.mark.parametrize("header, mult", [
    ("Notes £m £m", 1_000_000), ("Notes m Em", 1_000_000), ("Note 2025 2024 f'm Note 2075 E'm", 1_000_000),
    ("Notes million million million", 1_000_000), ("€m", 1_000_000),
    ("Note £000 £000", 1_000), ("£'000", 1_000), ("Note £ £", 1), ("2025 2024", 1),
    ("system item", 1),                                        # words ending in 'm' aren't units
])
def test_units(header, mult):
    assert t.unit_multiplier(header)[0] == mult


def test_units_come_from_headings_above_the_row_not_the_whole_page():
    lines = ["Strategic report: revenue up 5 m customers", "", "", "", "", "", "", "", "",
             "Note £ £", "Turnover 4 5,983,804 7,790,881"]
    assert t.scaled(t.turnover_from_lines(lines), lines)[0] == 5_983_804


def test_units_on_the_row_itself_win():
    lines = ["Notes £m £m", "Turnover £000 12,764"]
    assert t.scaled(t.turnover_from_lines(lines), lines)[0] == 12_764_000


@pytest.mark.parametrize("text, currency", [
    ("Notes £m £m", "GBP"), ("Notes Em Em", "GBP"), ("€ million", "EUR"), ("US$m", "USD"),
    ("Notes million million", ""),
])
def test_currency(text, currency):
    assert t.page_currency(text) == currency


@pytest.mark.parametrize("rows, expected", [
    (["Consolidated Financial Statements", "Unilever Group", "Consolidated income statement"], True),
    (["STRATECIC REPORT COVERNANCE", "CONSOLIDATEDINCOMESTATEMENT"], True),
    (["Strategic report Governance 121", "Groupincomestatement"], True),
    (["Small Example Limited", "Profit and Loss Account"], True),
    (["Statement of Comprehensive Income"], True),
    (["Consolidated statement of profit or loss"], True),
    (["WhitbreadGroupPLC", "Notestotheconsolidatedfinancialstatements(continued)"], False),
    (["NOTES TO THE CONSOLIDATED FINANCIAL STATEMENTS", "Income statement"], False),
    (["Balance Sheet", "As at 31 March 2025"], False),
])
def test_statement_pages(rows, expected):
    assert t.is_statement_page(rows) is expected


@pytest.mark.parametrize("raw, clean", [
    ("00445790", "00445790"), (445790, "00445790"), (445790.0, "00445790"), (" 0044 5790 ", "00445790"),
    ("sc005261", "SC005261"), ("SC5261", "SC005261"), ("OC318626", "OC318626"), ("NI 12345", "NI012345"),
    (None, ""), ("", ""), ("n/a", "NA"),
])
def test_clean_number(raw, clean):
    assert t.clean_number(raw) == clean


@pytest.mark.parametrize("description, reason", [
    ("Accounts for a dormant company made up to 31 October 2024", "dormant"),
    ("Micro company accounts made up to 31 October 2025", "micro"),
    ("Total exemption full accounts made up to 31 December 2020", "small"),
    ("Unaudited abridged accounts made up to 31 March 2025", "small"),
    ("Accounts for a small company made up to 31 March 2025", "small"),
    ("Accounts for a medium company made up to 31 March 2025", "medium"),
    ("Full accounts made up to 31 December 2024", ""),
    ("Group of companies' accounts made up to 28 February 2026", ""),
])
def test_why_missing(description, reason):
    got = t.why_missing(description)
    assert (reason in got) if reason else got == ""


def test_page_order_starts_a_quarter_in_for_long_reports():
    assert t.page_order(10) == list(range(10))
    order = t.page_order(200)
    assert order[0] == 50 and sorted(order) == list(range(200))


def test_name_matching():
    assert t.name_score("Tesco PLC", "TESCO PLC") == 100
    assert t.name_score("Marks & Spencer Group plc", "MARKS AND SPENCER GROUP P.L.C.") == 100
    assert t.name_score("The Greggs", "GREGGS PLC") == 100
    assert t.name_score("Zxqv Imaginary Trading Ltd", "IMAGINARY BEINGS LTD") < t.MATCH_THRESHOLD
