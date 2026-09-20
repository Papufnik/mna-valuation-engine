"""
Tests for lbo_model/build_lbo_model.py.

Two layers, matching how a real bug in this model actually got caught
during development:

1. Structural checks that don't need a spreadsheet engine -- every
   projection-period cell in the SBA amortization block must reference the
   PRIOR period's ENDING balance, not some other row. An earlier version of
   this script had an off-by-one (D6 referenced C9, "Principal Payment",
   instead of C10, "Ending Balance"), which produced a workbook that
   recalculated with zero formula errors and completely wrong downstream
   numbers (negative debt balances, an 13x+ MOIC from an arithmetic bug
   rather than the deal). A clean recalc proves formulas evaluate, not that
   they're right -- these tests check the formula text itself, not just
   that it runs.

2. An end-to-end check that actually recalculates the workbook with
   LibreOffice (`soffice --convert-to xlsx`, which forces a full
   recalculation as part of the conversion) and pins down the numbers a
   reviewer would sanity-check first: Sources & Uses balances, the debt
   schedule amortizes to a smaller balance every year, and IRR is
   internally consistent with MOIC. Skipped if soffice isn't installed.
"""
import os
import shutil
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(HERE)
LBO_DIR = os.path.join(REPO_ROOT, "lbo_model")
sys.path.insert(0, LBO_DIR)

import build_lbo_model  # noqa: E402

openpyxl = pytest.importorskip("openpyxl")

SOFFICE_AVAILABLE = shutil.which("soffice") is not None


@pytest.fixture(scope="module")
def built_wb_path():
    build_lbo_model.build()
    return build_lbo_model.OUT_FILE


@pytest.fixture(scope="module")
def wb(built_wb_path):
    return openpyxl.load_workbook(built_wb_path, data_only=False)


def test_all_expected_sheets_present(wb):
    assert wb.sheetnames == [
        "Assumptions",
        "Sources & Uses",
        "Operating Model",
        "Debt Schedule",
        "Returns Analysis",
        "Notes & Limitations",
    ]


def test_sba_amortization_references_prior_year_ending_balance_not_principal(wb):
    """Regression test for the off-by-one bug: Beginning Balance in each
    projection year (D6:G6) must reference the PRIOR column's row 10
    (Ending Balance), never row 9 (Principal Payment)."""
    ds = wb["Debt Schedule"]
    prior_col = {"D": "C", "E": "D", "F": "E", "G": "F"}
    for col, prev in prior_col.items():
        formula = ds[f"{col}6"].value
        assert formula == f"={prev}10", (
            f"Debt Schedule!{col}6 should reference {prev}10 (Ending Balance) "
            f"but was {formula!r}"
        )


def test_seller_note_amortization_references_prior_year_ending_balance(wb):
    ds = wb["Debt Schedule"]
    prior_col = {"D": "C", "E": "D", "F": "E", "G": "F"}
    for col, prev in prior_col.items():
        formula = ds[f"{col}14"].value
        assert formula == f"={prev}18", (
            f"Debt Schedule!{col}14 should reference {prev}18 (Ending Balance) "
            f"but was {formula!r}"
        )


def test_sources_and_uses_use_formulas_not_hardcoded_totals(wb):
    su = wb["Sources & Uses"]
    assert str(su["B6"].value).startswith("=")  # Total Uses
    assert str(su["B12"].value).startswith("=")  # Total Sources
    assert str(su["B13"].value).startswith("=")  # Check


def test_returns_analysis_irr_uses_moic_power_formula(wb):
    ra = wb["Returns Analysis"]
    assert ra["B12"].value == "=(B11)^(1/B4)-1"


@pytest.mark.skipif(not SOFFICE_AVAILABLE, reason="LibreOffice not available to recalculate")
def test_recalculated_model_produces_internally_consistent_values(built_wb_path, tmp_path):
    result = subprocess.run(
        [
            "soffice", "--headless",
            "--convert-to", "xlsx:Calc MS Excel 2007 XML",
            "--outdir", str(tmp_path),
            built_wb_path,
        ],
        capture_output=True, text=True, timeout=90,
    )
    recalced_path = tmp_path / "LBO_Model.xlsx"
    assert recalced_path.exists(), result.stdout + result.stderr

    wb_vals = openpyxl.load_workbook(recalced_path, data_only=True)

    su = wb_vals["Sources & Uses"]
    assert su["B6"].value == pytest.approx(su["B12"].value)  # Uses == Sources
    assert su["B13"].value == pytest.approx(0)  # Check row

    ds = wb_vals["Debt Schedule"]
    sba_balances = [ds[f"{c}10"].value for c in ["C", "D", "E", "F", "G"]]
    assert all(b > 0 for b in sba_balances)
    assert all(sba_balances[i] > sba_balances[i + 1] for i in range(len(sba_balances) - 1)), (
        "SBA ending balance must decline every year -- a flat or rising "
        "balance means the amortization schedule is broken"
    )

    ra = wb_vals["Returns Analysis"]
    moic = ra["B11"].value
    irr = ra["B12"].value
    years = ra["B4"].value
    assert moic > 1  # equity grew
    assert irr == pytest.approx(moic ** (1 / years) - 1, rel=1e-6)
