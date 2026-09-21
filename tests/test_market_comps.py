"""
Tests for market_comps.py -- pins down that it reads the valuation engine's
CURRENT assumed multiples (never a hardcoded copy), and that the comparison
logic correctly identifies which scenarios fall inside vs. outside the
cited comps range.
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

OUT_DIR = os.path.join(REPO_ROOT, "dashboard_app")


@pytest.fixture(autouse=True)
def clean_generated_output():
    if os.path.exists(OUT_DIR):
        shutil.rmtree(OUT_DIR)
    yield
    if os.path.exists(OUT_DIR):
        shutil.rmtree(OUT_DIR)


def test_raises_when_valuation_model_missing():
    with pytest.raises(FileNotFoundError):
        market_comps.benchmark_multiples()


def test_reads_live_multiples_not_a_hardcoded_copy(monkeypatch):
    """If mna_valuation_engine.py's multiples ever change, this check must
    compare against the NEW numbers automatically -- not a stale copy."""
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()

    # Prove the read is live: patch the JSON on disk to a different base
    # multiple and confirm the check picks it up.
    model_path = os.path.join(OUT_DIR, "mna_valuation_model.json")
    with open(model_path) as f:
        model = json.load(f)
    model["multiples"]["base"] = 2.45  # inside the 2.13x-2.8x comps range
    with open(model_path, "w") as f:
        json.dump(model, f)

    result = market_comps.benchmark_multiples()
    base_finding = next(f for f in result["assumedMultipleFindings"] if f["scenario"] == "base")
    assert base_finding["assumedMultiple"] == 2.45
    assert "within the cited comps range" in base_finding["readVsComps"]


def test_base_case_flagged_above_comps_with_real_assumptions(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()
    result = market_comps.benchmark_multiples()

    findings = {f["scenario"]: f for f in result["assumedMultipleFindings"]}
    assert "within the cited comps range" in findings["conservative"]["readVsComps"]
    assert "ABOVE every cited comp" in findings["base"]["readVsComps"]
    assert "ABOVE every cited comp" in findings["aggressive"]["readVsComps"]
    assert "sits above every cited market comp" in result["headlineFinding"]


def test_comps_are_real_sourced_data_not_placeholders():
    """Every cited comp must carry a real URL and a positive multiple --
    guards against someone swapping in illustrative/placeholder figures for
    what's supposed to be genuine third-party market data."""
    for comp in market_comps.MARKET_COMPS:
        assert comp["url"].startswith("https://")
        assert comp["sdeMultiple"] > 0
        assert comp["source"]
        assert comp["segment"]


def test_headline_finding_and_range_are_internally_consistent(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()
    result = market_comps.benchmark_multiples()

    comp_values = [c["sdeMultiple"] for c in market_comps.MARKET_COMPS]
    assert result["compsRange"]["low"] == min(comp_values)
    assert result["compsRange"]["high"] == max(comp_values)
    assert result["compsRange"]["average"] == pytest.approx(
        sum(comp_values) / len(comp_values), abs=0.01
    )
