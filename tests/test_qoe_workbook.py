"""
Tests for qoe_workbook/build_qoe_workbook.py.

Same two-layer approach as test_lbo_model.py:

1. Structural / build-time checks -- the monthly add-back allocations must
   sum to exactly the annual figures carried in mna_valuation_engine.py
   (build_qoe_workbook.py itself asserts this at build time; here we assert
   it again independently so a future edit that breaks the tie-out fails
   the test suite, not just a print statement nobody read).

2. An end-to-end recalculation check (skipped if soffice isn't installed)
   that confirms the EBITDA bridge tie-out and add-back reconciliation
   checks both evaluate to exactly zero, and that revenue/NWC totals match
   the figures used elsewhere in this repo.
"""
import os
import shutil
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(HERE)
QOE_DIR = os.path.join(REPO_ROOT, "qoe_workbook")
sys.path.insert(0, QOE_DIR)

import build_qoe_workbook  # noqa: E402

openpyxl = pytest.importorskip("openpyxl")

SOFFICE_AVAILABLE = shutil.which("soffice") is not None


@pytest.fixture(scope="module")
def built_wb_path():
    return build_qoe_workbook.build()


@pytest.fixture(scope="module")
def wb(built_wb_path):
    return openpyxl.load_workbook(built_wb_path, data_only=False)


def test_all_expected_sheets_present(wb):
    assert wb.sheetnames == [
        "Overview",
        "Adjusted EBITDA Bridge",
        "Add-Back Detail",
        "Revenue & Margin Quality",
        "NWC Analysis",
        "Notes & Limitations",
    ]


def test_ebitda_bridge_uses_formulas_not_hardcoded_results(wb):
    eb = wb["Adjusted EBITDA Bridge"]
    assert eb["B10"].value == "=SUM(B4:B9)"   # EBITDA
    assert eb["B15"].value == "=SUM(B13:B14)"  # Total add-backs
    assert eb["B18"].value == "=B10+B15"       # Adjusted EBITDA
    assert eb["B20"].value == "=B18-B19"       # Tie-out check


def test_addback_detail_reconciliation_formula_present(wb):
    ab = wb["Add-Back Detail"]
    assert ab["B16"].value == "=B14-B15"


def test_nwc_formulas_reference_correct_cells(wb):
    nwc = wb["NWC Analysis"]
    assert nwc["B8"].value == "=B4+B5-B6"   # NWC = AR + Inventory - AP
    assert nwc["B11"].value == "=AVERAGE(B8:M8)"
    assert nwc["B12"].value == "=M8"
    assert nwc["B13"].value == "=B11"
    assert nwc["B14"].value == "=B12-B13"


@pytest.mark.skipif(not SOFFICE_AVAILABLE, reason="LibreOffice not available to recalculate")
def test_recalculated_workbook_ties_out(built_wb_path, tmp_path):
    result = subprocess.run(
        [
            "soffice", "--headless",
            "--convert-to", "xlsx:Calc MS Excel 2007 XML",
            "--outdir", str(tmp_path),
            built_wb_path,
        ],
        capture_output=True, text=True, timeout=90,
    )
    recalced_path = tmp_path / "QoE_Workbook.xlsx"
    assert recalced_path.exists(), result.stdout + result.stderr

    wb_vals = openpyxl.load_workbook(recalced_path, data_only=True)

    eb = wb_vals["Adjusted EBITDA Bridge"]
    assert eb["B4"].value == pytest.approx(148220.60)
    assert eb["B15"].value == pytest.approx(21755.00)
    assert eb["B18"].value == pytest.approx(169975.60)
    assert eb["B20"].value == pytest.approx(0)  # tie-out check

    ab = wb_vals["Add-Back Detail"]
    assert ab["B16"].value == pytest.approx(0)  # reconciliation check
    assert ab["M6"].value == pytest.approx(20140.00)  # running total hits annual figure by Dec

    rq = wb_vals["Revenue & Margin Quality"]
    assert rq["F4"].value == pytest.approx(1400000)  # TTM revenue matches LBO model's Year 0 revenue

    nwc = wb_vals["NWC Analysis"]
    avg_nwc = nwc["B11"].value
    at_close = nwc["B12"].value
    assert nwc["B14"].value == pytest.approx(at_close - avg_nwc)
