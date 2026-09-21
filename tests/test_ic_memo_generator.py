"""
Tests for ic_memo_generator.py.

This generator reads four upstream outputs (valuation JSON, market comps
JSON, LBO_Model.xlsx, QoE_Workbook.xlsx) and writes nothing of its own --
so the tests focus on two things: (1) it fails loudly and specifically when
a prerequisite is missing or unrecalculated, rather than silently writing
"None" into a document someone might present as final, and (2) the
end-to-end memo actually contains the real, cross-checked figures pulled
from all four sources, not a fifth independent copy of any of them.

The LBO/QoE reads use the real, already-recalculated
lbo_model/LBO_Model.xlsx and qoe_workbook/QoE_Workbook.xlsx checked into
this repo -- these tests never call build_lbo_model.build() or
build_qoe_workbook.build(), so they can't leave those committed files in
the unrecalculated state test_lbo_model.py/test_qoe_workbook.py guard
against.
"""
import os
import sys
import json
import shutil

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(HERE)
sys.path.insert(0, REPO_ROOT)

from mna_valuation_engine import calculate_mna_valuation
import market_comps
import ic_memo_generator as gen

openpyxl = pytest.importorskip("openpyxl")
docx = pytest.importorskip("docx")

OUT_DIR = os.path.join(REPO_ROOT, "dashboard_app")
DEAL_ROOM_DIR = os.path.join(REPO_ROOT, "exit_package")
LBO_XLSX_PRESENT = os.path.exists(os.path.join(REPO_ROOT, "lbo_model", "LBO_Model.xlsx"))
QOE_XLSX_PRESENT = os.path.exists(os.path.join(REPO_ROOT, "qoe_workbook", "QoE_Workbook.xlsx"))


@pytest.fixture(autouse=True)
def clean_generated_output():
    for d in (OUT_DIR, DEAL_ROOM_DIR):
        if os.path.exists(d):
            shutil.rmtree(d)
    yield
    for d in (OUT_DIR, DEAL_ROOM_DIR):
        if os.path.exists(d):
            shutil.rmtree(d)


def _write_valuation_and_comps(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()
    market_comps.benchmark_multiples()


def test_raises_when_valuation_model_missing():
    with pytest.raises(FileNotFoundError, match="mna_valuation_engine"):
        gen.generate_ic_memo()


def test_raises_when_market_comps_missing(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()  # valuation exists, but market_comps.py was never run
    with pytest.raises(FileNotFoundError, match="market_comps"):
        gen.generate_ic_memo()


def test_raises_with_clear_message_when_lbo_workbook_unrecalculated(monkeypatch, tmp_path):
    """A workbook that exists but was never recalculated should fail with a
    message naming the exact cell and telling the user what to run --
    never silently write None/blank into the memo.

    read_lbo_figures(business_dir) builds its file path as
    <business_dir>/lbo_model/LBO_Model.xlsx, so pointing it at a throwaway
    tmp_path laid out the same way exercises the real function directly,
    with no monkeypatching of ic_memo_generator itself required.
    """
    sys.path.insert(0, os.path.join(REPO_ROOT, "lbo_model"))
    import build_lbo_model

    fake_business_dir = tmp_path
    lbo_dir = fake_business_dir / "lbo_model"
    lbo_dir.mkdir()
    unrecalculated_path = lbo_dir / "LBO_Model.xlsx"

    original_out_file = build_lbo_model.OUT_FILE
    build_lbo_model.OUT_FILE = str(unrecalculated_path)
    try:
        build_lbo_model.build()  # writes formulas only, never recalculated
    finally:
        build_lbo_model.OUT_FILE = original_out_file

    with pytest.raises(ValueError, match="hasn't been recalculated"):
        gen.read_lbo_figures(str(fake_business_dir))


@pytest.mark.skipif(not LBO_XLSX_PRESENT, reason="lbo_model/LBO_Model.xlsx not present in this checkout")
@pytest.mark.skipif(not QOE_XLSX_PRESENT, reason="qoe_workbook/QoE_Workbook.xlsx not present in this checkout")
def test_end_to_end_memo_contains_cross_checked_figures(monkeypatch):
    _write_valuation_and_comps(monkeypatch)

    out_file = gen.generate_ic_memo()
    assert os.path.exists(out_file)

    doc = docx.Document(out_file)
    full_text = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text += "\n" + cell.text

    # Figures that must appear, each traceable to one specific upstream file.
    assert "$169,975.60" in full_text or "169,975.60" in full_text  # valuation engine recastSDE
    assert "$509,926.80" in full_text  # valuation engine baseValuation
    assert "8.53x" in full_text  # LBO Returns Analysis MOIC
    assert "53.5%" in full_text  # LBO Returns Analysis IRR
    assert "2.63x" in full_text  # market_comps retail sector comp
    assert "148,220.60" in full_text  # QoE reported net income (verified figure)

    # The comps headline finding (the whole point of Section 4) must survive
    # into the document verbatim, not be paraphrased into something weaker.
    assert "sits above every cited market comp" in full_text
