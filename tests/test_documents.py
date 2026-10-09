"""Reading turnover out of real iXBRL filings and PDF accounts."""

import pytest

import turnover as t
from conftest import fixture, needs_ocr

# ---------------------------------------------------------------- iXBRL (real filings)


@pytest.mark.parametrize("name, current, prior, method", [
    ("ixbrl_tagged.html", 44_043_650, 47_988_261, "iXBRL tagged figure"),
    ("ixbrl_group_consolidated.html", 30_035_247, 26_242_640, "iXBRL tagged figure"),   # Consolidated dimension
    ("ixbrl_thousands_row.html", 12_764_297, None, "iXBRL tagged figure"),
    ("ixbrl_untagged_table.html", 10_491_000, None, "iXBRL table (untagged)"),
    ("ixbrl_stray_m.html", 5_983_804, 7_790_881, "iXBRL table (untagged)"),             # not £5.98 trillion
])
def test_real_ixbrl_filings(name, current, prior, method):
    got = t.read_ixbrl(fixture(name, "rb"))
    assert got[0] == current and got[1] == prior and got[4] == method and got[2] == "GBP"


def test_micro_entity_filing_has_no_turnover():
    assert t.read_ixbrl(fixture("ixbrl_micro.html", "rb")) is None


def test_tagged_filing_reports_period_and_length():
    got = t.read_ixbrl(fixture("ixbrl_tagged.html", "rb"))
    assert got[3] == "2025-12-31" and got[5] == 12


def ixbrl(facts, contexts):
    """A minimal iXBRL document."""
    ctx = "".join(f'<xbrli:context id="{cid}"><xbrli:entity><xbrli:identifier>1</xbrli:identifier>{seg}'
                  f'</xbrli:entity><xbrli:period><xbrli:startDate>{s}</xbrli:startDate>'
                  f'<xbrli:endDate>{e}</xbrli:endDate></xbrli:period></xbrli:context>'
                  for cid, s, e, seg in contexts)
    return (f'<html><body><ix:header><ix:resources>{ctx}<xbrli:unit id="GBP"><xbrli:measure>iso4217:GBP'
            f'</xbrli:measure></xbrli:unit><xbrli:unit id="EUR"><xbrli:measure>iso4217:EUR</xbrli:measure>'
            f'</xbrli:unit></ix:resources></ix:header>{facts}</body></html>')


SEGMENT = ('<xbrli:segment><xbrldi:explicitMember dimension="bus:X">bus:SegmentOne'
           '</xbrldi:explicitMember></xbrli:segment>')


def test_ixbrl_number_formats_scale_and_sign():
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c1" unitRef="EUR" scale="3" '
                'format="ixt:numcommadecimal">1.234,5</ix:nonFraction>',
                [("c1", "2024-01-01", "2024-12-31", "")])
    got = t.read_ixbrl(doc)
    assert got[0] == 1_234_500 and got[2] == "EUR"
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c1" unitRef="GBP" sign="-">500'
                '</ix:nonFraction>', [("c1", "2024-01-01", "2024-12-31", "")])
    assert t.read_ixbrl(doc)[0] == -500


def test_ixbrl_ignores_segment_breakdowns_and_picks_latest_year():
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="seg" unitRef="GBP">999</ix:nonFraction>'
                '<ix:nonFraction name="core:TurnoverRevenue" contextRef="py" unitRef="GBP">800</ix:nonFraction>'
                '<ix:nonFraction name="core:TurnoverRevenue" contextRef="cy" unitRef="GBP">900</ix:nonFraction>',
                [("seg", "2024-01-01", "2024-12-31", SEGMENT), ("py", "2023-01-01", "2023-12-31", ""),
                 ("cy", "2024-01-01", "2024-12-31", "")])
    got = t.read_ixbrl(doc)
    assert (got[0], got[1], got[3]) == (900, 800, "2024-12-31")


def test_ixbrl_long_period_is_measured():
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c" unitRef="GBP">100</ix:nonFraction>',
                [("c", "2023-07-01", "2024-12-31", "")])
    assert t.read_ixbrl(doc)[5] == 18


def test_ixbrl_dash_means_zero():
    doc = ixbrl('<ix:nonFraction name="core:TurnoverRevenue" contextRef="c" unitRef="GBP" '
                'format="ixt:fixed-zero">-</ix:nonFraction>', [("c", "2024-01-01", "2024-12-31", "")])
    assert t.read_ixbrl(doc)[0] == 0


# ---------------------------------------------------------------- PDFs


@needs_ocr
def test_scanned_annual_report_skips_notes_page_and_reads_statement():
    got = t.read_pdf(fixture("scanned_annual_report.pdf", "rb"))
    assert got.turnover == 1_234_500_000 and got.page == 3 and got.method == "PDF scan (OCR)"
    assert got.unit == "m" and got.currency in ("GBP", "")


def test_text_pdf_needs_no_ocr(monkeypatch):
    def boom():
        raise AssertionError("OCR should not run for a text PDF")
    monkeypatch.setattr(t, "ocr_engine", boom)
    got = t.read_pdf(fixture("text_small_company.pdf", "rb"))
    assert got.turnover == 5_432_100 and got.method == "PDF text" and got.currency == "GBP"


@needs_ocr
def test_balance_sheet_only_pdf_has_no_turnover():
    got = t.read_pdf(fixture("scanned_balance_sheet_only.pdf", "rb"))
    assert got.turnover is None and got.pages_checked == 1 and not got.problem


@needs_ocr
def test_page_limit_is_respected():
    got = t.read_pdf(fixture("scanned_annual_report.pdf", "rb"), max_pages=2)
    assert got.turnover is None and got.pages_checked == 2


def test_scanned_pages_can_be_skipped():
    got = t.read_pdf(fixture("scanned_annual_report.pdf", "rb"), read_scanned=False)
    assert got.turnover is None and got.scanned_skipped


def test_corrupt_pdf_is_reported_not_raised():
    got = t.read_pdf(b"this is not a pdf")
    assert got.turnover is None and "couldn't be opened" in got.problem


def test_missing_ocr_engine_is_reported_not_raised(monkeypatch):
    def unavailable():
        raise t.OcrUnavailable("no module named onnxruntime")
    monkeypatch.setattr(t, "ocr_engine", unavailable)
    got = t.read_pdf(fixture("scanned_annual_report.pdf", "rb"))
    assert got.turnover is None and "isn't working" in got.problem


@needs_ocr
def test_progress_messages_are_sent():
    seen = []
    t.read_pdf(fixture("scanned_annual_report.pdf", "rb"), progress=seen.append)
    assert seen and "page 1 of 3" in seen[0]


def test_reader_that_cant_start_is_reported_and_logged(monkeypatch, capsys):
    import builtins
    real_import = builtins.__import__

    def broken(name, *a, **k):
        if name == "rapidocr_onnxruntime":
            raise ImportError("libGL.so.1: cannot open shared object file")
        return real_import(name, *a, **k)

    monkeypatch.setattr(t, "_ocr", None)
    monkeypatch.setattr(t, "_ocr_problem", None)
    monkeypatch.setattr(builtins, "__import__", broken)
    assert t.ocr_available() is False
    assert "libGL.so.1" in capsys.readouterr().err                       # the reason is in the logs
    got = t.read_pdf(fixture("scanned_annual_report.pdf", "rb"))
    assert got.turnover is None and "isn't working" in got.problem
