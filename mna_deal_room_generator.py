import os
import sys
import json
from datetime import datetime

def generate_deal_room_package():
    business_dir = os.path.dirname(os.path.abspath(__file__))
    deal_room_dir = os.path.join(business_dir, "exit_package")
    out_dir = os.path.join(business_dir, "dashboard_app")
    os.makedirs(deal_room_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    print(f"=== M&A Turn-Key Deal Room Generator [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ===")

    # Pull numbers from mna_valuation_model.json (mna_valuation_engine.py's
    # output) instead of hardcoding a second, independent copy of the same
    # figures here. Before 2026-07-31 this file hardcoded its own
    # $201,455 SDE / $604,365 valuation / 58.4% margin directly in the HTML
    # template -- a second stale copy of the same bug the valuation engine
    # had, since the two never actually referenced each other.
    model_path = os.path.join(out_dir, "mna_valuation_model.json")
    if not os.path.exists(model_path):
        raise FileNotFoundError(
            f"{model_path} not found -- run mna_valuation_engine.py first so "
            "this generator has real numbers to read instead of guessing."
        )
    with open(model_path, "r", encoding="utf-8") as f:
        model = json.load(f)

    recast_sde = model["recastSDE"]
    base_valuation = model["valuations"]["baseValuation"]
    ready = model.get("dealRoomReadiness", "")
    is_ready = ready.strip().upper().startswith("100%") or ready.strip().upper().startswith("READY")

    draft_banner = "" if is_ready else (
        '<div style="background:#FEF3C7;border:1px solid #D97706;color:#92400E;'
        'padding:12px 16px;border-radius:8px;margin-bottom:24px;font-size:13px;">'
        '<strong>DRAFT -- NOT FOR EXTERNAL USE.</strong> '
        f'{ready} Figures below (SDE, valuation) are live from the reconciled '
        'tax package, but the "Investment Highlights" claims (margin %, YoY '
        'growth) have NOT been re-verified against the current reporting '
        'suite and may be stale -- check against build_gmroi.py / '
        'build_yoy_comparison.py output before sharing.</div>'
    )

    # 1. Generate Executive Teaser HTML
    teaser_html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Confidential Business Review — [Business Name Redacted]</title>
  <style>
    body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background: #FAF7F2; color: #2D1820; margin: 0; padding: 40px; }}
    .deck {{ max-width: 800px; margin: 0 auto; background: #FFFFFF; border-radius: 12px; border: 1px solid #E5D5C5; padding: 40px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); }}
    .header {{ text-align: center; border-bottom: 2px solid #C9908A; padding-bottom: 24px; margin-bottom: 32px; }}
    .title {{ font-size: 28px; font-weight: bold; color: #8B3A4A; letter-spacing: 0.5px; }}
    .subtitle {{ font-size: 14px; color: #7A6066; margin-top: 6px; text-transform: uppercase; letter-spacing: 1px; }}
    .grid {{ display: table; width: 100%; margin: 24px 0; }}
    .cell {{ display: table-cell; width: 33%; text-align: center; padding: 16px; background: #F7EFE9; border-radius: 8px; font-weight: bold; }}
    .val {{ font-size: 24px; color: #047857; margin-top: 6px; }}
    .lbl {{ font-size: 11px; text-transform: uppercase; color: #7A6066; letter-spacing: 0.5px; }}
    .section {{ margin-top: 32px; }}
    .section-h {{ font-size: 18px; font-weight: bold; color: #8B3A4A; border-bottom: 1px solid #E5D5C5; padding-bottom: 8px; margin-bottom: 16px; }}
    ul {{ line-height: 1.6; color: #4A353B; }}
    .footer {{ text-align: center; margin-top: 40px; font-size: 11px; color: #9E858B; border-top: 1px solid #E5D5C5; padding-top: 16px; }}
  </style>
</head>
<body>
  <div class="deck">
    {draft_banner}
    <div class="header">
      <div class="title">THE BUSINESS</div>
      <div class="subtitle">the operating LLC • Confidential Business Overview • [address redacted]</div>
    </div>

    <div class="grid">
      <div class="cell">
        <div class="lbl">Recast Annual SDE</div>
        <div class="val">${recast_sde:,.0f}</div>
      </div>
      <div class="cell">
        <div class="lbl">Target Enterprise Value</div>
        <div class="val">${base_valuation:,.0f}</div>
      </div>
      <div class="cell">
        <div class="lbl">Gross Margin</div>
        <div class="val">UNVERIFIED</div>
      </div>
    </div>

    <div class="section">
      <div class="section-h">Investment Highlights</div>
      <ul>
        <li><strong>Prime Downtown Retail Footprint</strong>: Flagship location in a high-traffic downtown retail corridor of a mountain tourist destination town.</li>
        <li><strong>[UNVERIFIED -- confirm against build_gmroi.py before use]</strong> High-Margin Hybrid Model claim (Permanent Jewelry / Wine Bar category margins).</li>
        <li><strong>[UNVERIFIED -- confirm against build_yoy_comparison.py before use]</strong> Revenue and YoY growth claim.</li>
        <li><strong>Turn-Key Operations</strong>: Automated the POS platform, Plaid Bank Integration, and AI Analytics Engine installed and operating.</li>
      </ul>
    </div>

    <div class="footer">
      CONFIDENTIAL • Prepared for Qualified Buyers & Investors • the operating LLC
    </div>
  </div>
</body>
</html>
"""
    teaser_file = os.path.join(deal_room_dir, "01_Executive_Teaser_Summary.html")
    with open(teaser_file, "w", encoding="utf-8") as f:
        f.write(teaser_html)

    deal_room_summary = {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "dealRoomFolder": deal_room_dir,
        "executiveTeaser": teaser_file,
        "recastSDE": recast_sde,
        "askingValuation": base_valuation,
        "status": "DRAFT -- pending owner sign-off" if not is_ready else "Buyer-Ready Data Room Assembled",
    }

    out_file = os.path.join(out_dir, "mna_deal_room_summary.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(deal_room_summary, f, indent=2)

    print(f"[SUCCESS] Deal Room Package Assembled: Teaser created at {teaser_file} -> {out_file}")
    return deal_room_summary

if __name__ == "__main__":
    generate_deal_room_package()

