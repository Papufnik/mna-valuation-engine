import os
import sys
import json
from datetime import datetime

try:
    import openpyxl
except ImportError:
    openpyxl = None

TAX_PACKAGE_REL = os.path.join(
    "sample_data", "business_tax_package_latest.xlsx"
)


def get_verified_net_income(business_dir):
    """Pull 2025 Net Income live from the tax package instead of hardcoding
    it. Falls back loudly (not silently) if the file or openpyxl is missing,
    since a stale hardcoded number here is exactly the bug this replaced --
    see CREDENTIALS_README.md sibling note in this folder for the fix
    history (2026-07-31)."""
    path = os.path.join(business_dir, TAX_PACKAGE_REL)
    if openpyxl is None:
        raise RuntimeError(
            "openpyxl not available -- cannot safely compute a live valuation. "
            "Install openpyxl rather than falling back to a guessed number."
        )
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Tax package not found at {path} -- cannot safely compute a live "
            "valuation without it."
        )
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["2 - 2025 Income Statement"]
    net_income = ws["B59"].value
    if not isinstance(net_income, (int, float)):
        raise ValueError(
            f"Tax package B59 (Net Income) read back as {net_income!r}, not a "
            "number -- the cached formula value may be stale. Re-open the "
            "workbook in Excel (or re-run the cached-value injection) before "
            "trusting this valuation."
        )
    return round(float(net_income), 2)


def calculate_mna_valuation():
    business_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(business_dir, "dashboard_app")
    os.makedirs(out_dir, exist_ok=True)

    print(f"=== M&A Exit Valuation & Recast SDE Engine [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ===")

    # Financial recasting & Seller's Discretionary Earnings (SDE) model.
    # Net income is read LIVE from the verified tax package (Tab 2, B59) --
    # it used to be hardcoded at $164,500.00, which was stale/wrong (the
    # real, fully-reconciled 2025 figure is $148,220.60 as of 2026-07-31;
    # see that tab's own methodology notes for how it was derived and
    # verified). Never hardcode this number again -- if the tax package
    # changes, this valuation should move with it automatically.
    reported_net_income = get_verified_net_income(business_dir)

    # Discretionary & non-recurring add-backs.
    # Two previous entries were REMOVED here (2026-07-31 review) because they
    # contradict facts already verified elsewhere in this engagement:
    #   - "Local Sales Tax Restatement Adjustment" ($9,240): sales tax
    #     is already correctly excluded from revenue in the verified income
    #     statement (see Tab 3 / Tab 0 methodology -- gross receipts are
    #     tax-inclusive and net revenue already backs out the 8.7% tax).
    #     Adding this back on top would double-count.
    #   - "Related-Entity Loan Interest Elimination" ($4,890): the related
    #     entity's intercompany loan was confirmed to be 0%-interest (see the
    #     transmittal letter and Tax Package Tab 10, item on that
    #     loan) -- there is no interest expense to eliminate.
    # The two remaining add-backs are still UNVERIFIED against source
    # documents (unlike everything in the tax package itself, which was
    # independently reconciled) -- flagged here rather than presented with
    # false confidence. Confirm both with the owner before using this in
    # front of a buyer or broker.
    add_backs = [
        {"item": "Owner Discretionary Draw & Health Benefits", "amount": 20140.00,
         "category": "Owner Compensation", "verified": False,
         "note": "Source not yet traced to bank/ledger detail -- confirm with owner before relying on this figure."},
        {"item": "One-Time Store Improvement (Wayfair Fixtures)", "amount": 1615.00,
         "category": "Non-Recurring CapEx", "verified": False,
         "note": "Source not yet traced to bank/ledger detail -- confirm with owner before relying on this figure."},
    ]

    total_add_backs = sum(a["amount"] for a in add_backs)
    recast_sde = round(reported_net_income + total_add_backs, 2)

    # Multiple valuation range (Specialty Retail Hybrids: 2.5x - 3.5x SDE)
    # NOTE: the multiple range itself is a judgment call / market comp, not
    # derived from this business's own data -- treat it the same way the
    # rest of this suite treats a risk-tolerance constant: confirm before
    # trusting, cite a source (broker comps, BizBuySell data, etc.) if one exists.
    multiples = {
        "conservative": 2.50,
        "base": 3.00,
        "aggressive": 3.50
    }

    valuation_range = {
        "conservativeValuation": round(recast_sde * multiples["conservative"], 2),
        "baseValuation": round(recast_sde * multiples["base"], 2),
        "aggressiveValuation": round(recast_sde * multiples["aggressive"], 2),
        "assetLiquidationFloor": 216340.00,
        "assetLiquidationFloorVerified": False,
        "assetLiquidationFloorNote": (
            "UNVERIFIED estimate carried over from before the 2026-07-31 review "
            "($60.4k cash + $155.9k inventory). Pull live cash + inventory "
            "figures from business_balance_sheet.xlsx before using this number."
        ),
    }

    valuation_model = {
        "asOfDate": datetime.now().strftime("%B %d, %Y"),
        "reportedNetIncome": reported_net_income,
        "netIncomeSource": "Live from business_tax_package_latest.xlsx, Tab 2 B59 (fully reconciled 2026-07-31)",
        "totalAddBacks": total_add_backs,
        "recastSDE": recast_sde,
        "multiples": multiples,
        "valuations": valuation_range,
        "addBackBreakdown": add_backs,
        "dealRoomReadiness": "NOT READY -- add-backs and asset liquidation floor need owner sign-off before external use (see addBackBreakdown/assetLiquidationFloorNote)",
    }

    out_file = os.path.join(out_dir, "mna_valuation_model.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(valuation_model, f, indent=2)

    print(f"[SUCCESS] Valuation Calculated: Base SDE ${recast_sde:,.2f} -> Base Value ${valuation_range['baseValuation']:,.2f} -> {out_file}")
    return valuation_model

if __name__ == "__main__":
    calculate_mna_valuation()

