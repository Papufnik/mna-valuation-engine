"""
Tests for mna_valuation_engine.py -- pin down the two things the README
calls out explicitly: net income must be read LIVE from the verified tax
package (never hardcoded), and it must fail loudly rather than guess when
that source data isn't trustworthy.
"""
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(HERE)
sys.path.insert(0, REPO_ROOT)

from mna_valuation_engine import get_verified_net_income, calculate_mna_valuation


def test_get_verified_net_income_reads_live_value(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    sample_dir = tmp_path / "sample_data"
    sample_dir.mkdir()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2 - 2025 Income Statement"
    ws["A59"] = "Net Income"
    ws["B59"] = 148220.60
    wb.save(sample_dir / "business_tax_package_latest.xlsx")

    assert get_verified_net_income(str(tmp_path)) == 148220.60


def test_get_verified_net_income_fails_loudly_when_file_missing(tmp_path):
    with pytest.raises(FileNotFoundError):
        get_verified_net_income(str(tmp_path))


def test_get_verified_net_income_fails_loudly_on_non_numeric_cell(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    sample_dir = tmp_path / "sample_data"
    sample_dir.mkdir()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2 - 2025 Income Statement"
    ws["B59"] = "TBD -- pending reconciliation"  # not a number
    wb.save(sample_dir / "business_tax_package_latest.xlsx")

    with pytest.raises(ValueError):
        get_verified_net_income(str(tmp_path))


def test_recast_sde_and_valuation_math_matches_verified_net_income(monkeypatch):
    """End-to-end regression: with net income of $148,220.60 and the two
    documented (unverified but included) add-backs totaling $21,755.00,
    recast SDE must be exactly $169,975.60, and the base valuation must be
    exactly 3.00x that -- pinning down the same numbers the README's demo
    output shows ("Base SDE $169,975.60 -> Base Value $509,926.80")."""
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    model = calculate_mna_valuation()

    assert model["reportedNetIncome"] == 148220.60
    assert model["totalAddBacks"] == pytest.approx(21755.00)
    assert model["recastSDE"] == pytest.approx(169975.60)
    assert model["valuations"]["baseValuation"] == pytest.approx(169975.60 * 3.00)
    assert model["valuations"]["conservativeValuation"] == pytest.approx(169975.60 * 2.50)
    assert model["valuations"]["aggressiveValuation"] == pytest.approx(169975.60 * 3.50)
    assert model["dealRoomReadiness"].startswith("NOT READY")


def test_add_backs_are_individually_flagged_unverified(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    model = calculate_mna_valuation()
    assert len(model["addBackBreakdown"]) == 2
    assert all(item["verified"] is False for item in model["addBackBreakdown"])
