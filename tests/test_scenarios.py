"""Every route through one company's lookup, against saved Companies House pages."""

import requests

import turnover as t
from conftest import PDF0, XHTML0, FakeWeb, company, filing_page, fixture, needs_ocr

N = "01234567"


def run(routes, row=None, **kw):
    web = FakeWeb(routes)
    return t.safe_check(row or {"number": N}, web, **kw), web


# ---------------------------------------------------------------- the main routes


def test_tagged_ixbrl_is_found_exactly():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True))
    res, _ = run(company(N, filings=page, xhtml=fixture("ixbrl_tagged.html", "rb")))
    assert res.status == "FOUND" and res.turnover == 44_043_650 and res.prior_turnover == 47_988_261
    assert res.period_end == "2025-12-31" and res.method == "iXBRL tagged figure" and res.note == ""


def test_micro_accounts_are_not_disclosed_with_reason_and_no_pdf_download():
    page = filing_page(("29 Jul 2026", "AA", "Micro company accounts made up to 31 October 2025", True, True))
    res, web = run(company(N, filings=page, xhtml=fixture("ixbrl_micro.html", "rb")))
    assert res.status == "NOT DISCLOSED" and "micro-entity" in res.note
    assert PDF0 not in web.calls


def test_real_micro_filing_history_page():
    res, _ = run(company(N, filings=fixture("filing_micro.html"), extra={
        t.BASE + "/company/07000000/filing-history/MzUzNTc4MTMzM2FkaXF6a2N4/document?format=xhtml&download=1":
            fixture("ixbrl_micro.html", "rb")}))
    assert res.status == "NOT DISCLOSED" and "micro" in res.note and res.filed_on == "29 Jul 2026"


@needs_ocr
def test_scanned_pdf_only_is_read_by_ocr_and_marked_verify():
    page = filing_page(("25 Jul 2026", "AA", "Group of companies' accounts made up to 28 February 2026", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("scanned_annual_report.pdf", "rb")))
    assert res.status == "VERIFY" and res.turnover == 1_234_500_000
    assert "page 3" in res.method and "check the figure" in res.note


def test_text_pdf_is_found_without_verify():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 30 June 2025", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("text_small_company.pdf", "rb")))
    assert res.status == "FOUND" and res.turnover == 5_432_100 and res.currency == "GBP"


# ---------------------------------------------------------------- fallbacks


def test_ixbrl_download_failure_falls_back_to_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 30 June 2025", True, True))
    routes = company(N, filings=page, pdf=fixture("text_small_company.pdf", "rb"))
    routes[XHTML0] = t.SourceError("Companies House server error (HTTP 503), tried 4 times")
    res, _ = run(routes)
    assert res.status == "FOUND" and res.turnover == 5_432_100


@needs_ocr
def test_full_accounts_with_untagged_turnover_fall_back_to_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True))
    res, web = run(company(N, filings=page, xhtml=fixture("ixbrl_micro.html", "rb"),
                           pdf=fixture("scanned_annual_report.pdf", "rb")))
    assert PDF0 in web.calls and res.status == "VERIFY" and res.turnover == 1_234_500_000


def test_full_accounts_with_no_turnover_anywhere_are_not_disclosed():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True))
    res, _ = run(company(N, filings=page, xhtml=fixture("ixbrl_micro.html", "rb"),
                         pdf=fixture("scanned_balance_sheet_only.pdf", "rb")))
    assert res.status == "NOT DISCLOSED" and "not shown" in res.note


def test_small_company_scanned_pdf_without_pl_is_not_disclosed():
    page = filing_page(("01 Mar 2026", "AA", "Total exemption full accounts made up to 31 March 2025", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("scanned_balance_sheet_only.pdf", "rb")))
    assert res.status == "NOT DISCLOSED" and "small company" in res.note


@needs_ocr
def test_full_accounts_pdf_with_no_turnover_line_says_check_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 March 2025", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("scanned_balance_sheet_only.pdf", "rb")))
    assert res.status == "CHECK PDF" and "open the PDF" in res.note


def test_pdf_missing_says_check_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 March 2025", True, False))
    res, _ = run(company(N, filings=page))                                     # PDF route is a 404
    assert res.status == "CHECK PDF" and "couldn't be downloaded" in res.note


def test_filing_without_any_document_says_check_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 March 2025", False, False))
    res, _ = run(company(N, filings=page))
    assert res.status == "CHECK PDF" and "no downloadable document" in res.note


def test_corrupt_pdf_says_check_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 March 2025", True, False))
    res, _ = run(company(N, filings=page, pdf=b"%PDF-1.4 garbage"))
    assert res.status == "CHECK PDF" and "couldn't be opened" in res.note


def test_scanned_reading_turned_off_says_check_pdf():
    page = filing_page(("01 Mar 2026", "AA", "Group of companies' accounts made up to 28 February 2026", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("scanned_annual_report.pdf", "rb")), read_scanned=False)
    assert res.status == "CHECK PDF" and "turned off" in res.note


# ---------------------------------------------------------------- choosing the right filing


def test_interim_accounts_are_skipped_for_the_annual_ones():
    res, _ = run(company(N, filings=fixture("filing_interim_first.html")))
    assert "Interim" not in res.accounts_type and "28 June 2025" in res.accounts_type


def test_accounts_on_a_later_history_page_are_found():
    routes = company(N, filings=filing_page(("16 Sep 2026", "SH03", "Purchase of own shares", True, False),
                                            next_page=True))
    routes[f"/company/{N}/filing-history?page=2"] = filing_page(
        ("01 Mar 2026", "AA", "Full accounts made up to 30 June 2025", True, False))
    routes[PDF0] = fixture("text_small_company.pdf", "rb")
    res, _ = run(routes)
    assert res.status == "FOUND" and res.turnover == 5_432_100


def test_stale_accounts_are_noted():
    page = filing_page(("09 Jul 2022", "AA", "Total exemption full accounts made up to 31 December 2020", True, True))
    res, _ = run(company(N, filings=page, xhtml=fixture("ixbrl_micro.html", "rb")))
    assert "over 2 years old" in res.note and res.period_end == "2020-12-31"


def test_non_annual_period_is_noted():
    from test_documents import ixbrl
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c" unitRef="GBP">100</ix:nonFraction>',
                [("c", "2024-07-01", "2025-12-31", "")])
    page = filing_page(("01 Mar 2026", "AA", "Full accounts made up to 31 December 2025", True, True))
    res, _ = run(company(N, filings=page, xhtml=doc))
    assert res.status == "FOUND" and "18 months" in res.note


# ---------------------------------------------------------------- the company itself


def test_new_company_without_accounts_needs_no_filing_lookup():
    res, web = run(company(N, profile="profile_new.html"))
    assert res.status == "NOT DISCLOSED" and "no accounts filed yet" in res.note
    assert not any("filing-history" in c for c in web.calls)


def test_company_with_no_accounts_filings_at_all():
    routes = {f"/company/{N}": '<h1 id="company-name">NO ACCOUNTS LTD</h1><dd id="company-status">Active</dd>',
              f"/company/{N}/filing-history?page=1": filing_page()}
    res, _ = run(routes)
    assert res.status == "NOT DISCLOSED" and "no accounts filed" in res.note


def test_overseas_company_old_accounts_are_found_in_history():
    res, _ = run(company(N, profile="profile_overseas.html", filings=fixture("filing_overseas.html")))
    assert "31 December 2014" in res.accounts_type and res.company_status == "Active"


def test_dissolved_company_is_noted():
    page = filing_page(("01 Mar 2008", "AA", "Group of companies' accounts made up to 3 February 2008", True, False))
    res, _ = run(company(N, profile="profile_dissolved.html", filings=page, pdf=fixture("text_small_company.pdf", "rb")))
    assert res.company_status == "Dissolved" and "dissolved" in res.note and "over 2 years old" in res.note


def test_liquidation_and_overdue_accounts_are_noted():
    page = filing_page(("01 Jan 2019", "AA", "Group of companies' accounts made up to 30 September 2018", True, False))
    res, _ = run(company(N, profile="profile_liquidation.html", filings=page, pdf=fixture("text_small_company.pdf", "rb")))
    assert "liquidation" in res.note and "accounts overdue" in res.note


def test_accounts_known_but_not_found_in_history():
    res, _ = run(company(N, filings=filing_page(("16 Sep 2026", "SH03", "Purchase of own shares", True, False))))
    assert res.status == "CHECK PDF" and "weren't found" in res.note


def test_unknown_number_is_not_found():
    res, _ = run({})
    assert res.status == "NOT FOUND" and "not on the register" in res.note


def test_number_formats_from_excel_are_cleaned():
    page = filing_page(("01 Mar 2026", "AA", "Micro company accounts made up to 31 October 2025", True, True))
    res, web = run(company("00445790", filings=page, xhtml=fixture("ixbrl_micro.html", "rb")),
                   row={"number": 445790.0})
    assert res.company_number == "00445790" and res.status == "NOT DISCLOSED"


def test_name_and_number_that_disagree_are_flagged():
    page = filing_page(("01 Mar 2026", "AA", "Micro company accounts made up to 31 October 2025", True, True))
    res, _ = run(company(N, filings=page, xhtml=fixture("ixbrl_micro.html", "rb")),
                 row={"name": "Unilever PLC", "number": N})
    assert "the number belongs to TESCO PLC" in res.note


def test_search_by_name_picks_the_right_company():
    routes = {"/search/companies?q=Greggs plc": fixture("search_greggs.html")}
    routes.update(company("00502851", profile="profile_new.html"))
    res, _ = run(routes, row={"name": "Greggs plc"})
    assert res.company_number == "00502851"


def test_search_with_no_good_match_is_not_found():
    res, _ = run({"/search/companies?q=Zxqv Imaginary Trading Ltd": fixture("search_nomatch.html")},
                 row={"name": "Zxqv Imaginary Trading Ltd"})
    assert res.status == "NOT FOUND" and "no confident match" in res.note


def test_empty_row_is_not_found():
    res, _ = run({}, row={"name": "", "number": None})
    assert res.status == "NOT FOUND" and "no company name or number" in res.note


# ---------------------------------------------------------------- failures stay on their own row


def test_network_failure_is_an_error_row():
    res, _ = run({f"/company/{N}": t.SourceError("network error (ConnectionError), tried 4 times")})
    assert res.status == "ERROR" and "network error" in res.note


def test_unexpected_bug_is_an_error_row_not_a_crash():
    res, _ = run({f"/company/{N}": ValueError("something odd")})
    assert res.status == "ERROR" and "unexpected problem" in res.note


# ---------------------------------------------------------------- the real Web class's retries


class FlakySession:
    """requests.Session stand-in: plays back a list of status codes / exceptions."""

    def __init__(self, script):
        self.script, self.headers, self.calls = list(script), {}, 0

    def get(self, url, params=None, timeout=None):
        self.calls += 1
        step = self.script.pop(0)
        if isinstance(step, Exception):
            raise step
        r = requests.Response()
        r.status_code, r._content = step, b"ok"
        return r


def web_with(script):
    w = t.Web(delay=0)
    w.s = FlakySession(script)
    return w


def test_retries_then_succeeds(no_sleep):
    w = web_with([requests.ConnectionError(), 503, 429, 200])
    assert w.get("/x").content == b"ok" and w.s.calls == 4


def test_gives_up_after_retries(no_sleep):
    w = web_with([503] * t.RETRIES)
    try:
        w.get("/x")
    except t.SourceError as e:
        assert "server error" in str(e) and "tried 4 times" in str(e)
    else:
        raise AssertionError("expected SourceError")


def test_404_is_none_and_403_is_refused(no_sleep):
    assert web_with([404]).get("/x") is None
    try:
        web_with([403]).get("/x")
    except t.SourceError as e:
        assert "refused" in str(e)


@needs_ocr
def test_scanned_page_without_currency_symbol_is_pounds_without_a_warning():
    page = filing_page(("25 Jul 2026", "AA", "Group of companies' accounts made up to 28 February 2026", True, False))
    res, _ = run(company(N, filings=page, pdf=fixture("scanned_annual_report.pdf", "rb")))
    assert res.status == "VERIFY" and res.currency == "GBP" and "currency" not in res.note
