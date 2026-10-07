"""Checks against the real Companies House website. Slow (a few minutes) and depends on what
companies have filed, so they only run when asked:

    RUN_LIVE=1 python -m pytest tests/test_live.py -v

Figures are checked as ranges, because companies file new accounts every year.
"""

import os

import pytest

import turnover as t

pytestmark = pytest.mark.skipif(not os.environ.get("RUN_LIVE"), reason="set RUN_LIVE=1 to run live checks")

WEB = t.Web()


def check(**row):
    return t.safe_check(row, WEB)


def test_large_company_scanned_report_by_number():
    res = check(number="00445790")                     # Tesco: 200+ page scanned annual report
    assert res.status == "VERIFY" and 50e9 < res.turnover < 120e9 and res.company_name == "TESCO PLC"


def test_large_company_by_name_only():
    res = check(name="Greggs plc")
    assert res.company_number == "00502851" and res.status == "VERIFY" and 1e9 < res.turnover < 5e9


def test_name_written_differently_to_the_register():
    res = check(name="Marks & Spencer Group plc")      # register says MARKS AND SPENCER GROUP P.L.C.
    assert res.company_number == "04256886"


def test_micro_company():
    res = check(number="07000000")
    assert res.status == "NOT DISCLOSED" and "micro" in res.note


def test_dormant_company():
    res = check(number="11000000")
    assert res.status == "NOT DISCLOSED" and "dormant" in res.note


def test_new_company_without_accounts():
    res = check(number="16900000")
    assert res.status == "NOT DISCLOSED" and "no accounts" in res.note


def test_dissolved_company():
    res = check(number="03855289")                     # Woolworths Group plc, dissolved 2015
    assert res.company_status == "Dissolved" and "dissolved" in res.note


def test_unknown_number():
    res = check(number="ZZ123456")
    assert res.status == "NOT FOUND"


def test_unknown_name():
    res = check(name="Zxqv Imaginary Trading Ltd")
    assert res.status == "NOT FOUND" and "no confident match" in res.note


def test_keywords_in_electronic_accounts():
    res = t.safe_check({"number": "00442329"}, WEB, keywords=["derivative", "presentation currency", "zzqx"])
    assert res.keywords_found == ["derivative", "presentation currency"]
