"""
Tests for mna_deal_room_generator.py -- confirms it refuses to run without
a valuation model on disk, and that the draft banner / unverified-claim
labeling actually appear when the valuation isn't marked ready (which is
the only state the sample data ever produces, since add-backs always need
owner sign-off first).

These run against the real dashboard_app/ and exit_package/ output
locations, the same way demo.py does -- both are gitignored, so running
this suite doesn't dirty the repo.
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
import mna_deal_room_generator as gen

OUT_DIR = os.path.join(REPO_ROOT, "dashboard_app")
DEAL_ROOM_DIR = os.path.join(REPO_ROOT, "exit_package")


@pytest.fixture(autouse=True)
def clean_generated_output():
    """Runs before and after every test so this suite is idempotent and
    never depends on a previous test's leftover files."""
    for d in (OUT_DIR, DEAL_ROOM_DIR):
        if os.path.exists(d):
            shutil.rmtree(d)
    yield
    for d in (OUT_DIR, DEAL_ROOM_DIR):
        if os.path.exists(d):
            shutil.rmtree(d)


def test_raises_when_valuation_model_missing():
    with pytest.raises(FileNotFoundError):
        gen.generate_deal_room_package()


def test_draft_banner_and_unverified_labels_appear_when_not_ready(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()  # writes the real mna_valuation_model.json

    summary = gen.generate_deal_room_package()

    with open(summary["executiveTeaser"], "r", encoding="utf-8") as f:
        html = f.read()

    assert "DRAFT -- NOT FOR EXTERNAL USE" in html
    assert "[UNVERIFIED -- confirm against build_gmroi.py before use]" in html
    assert "$169,976" in html  # recast SDE, comma-formatted, rounded per the template's :,.0f
    assert summary["status"] == "DRAFT -- pending owner sign-off"


def test_deal_room_summary_json_is_also_written(monkeypatch):
    monkeypatch.setattr(
        "mna_valuation_engine.get_verified_net_income", lambda business_dir: 148220.60
    )
    calculate_mna_valuation()
    gen.generate_deal_room_package()

    summary_path = os.path.join(OUT_DIR, "mna_deal_room_summary.json")
    assert os.path.exists(summary_path)
    with open(summary_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    assert data["recastSDE"] == pytest.approx(169975.60)
