"""
market_comps.py -- checks mna_valuation_engine.py's SDE multiple assumption
against real, sourced market data instead of leaving it as an unexamined
judgment call.

WHY THIS EXISTS: mna_valuation_engine.py has always been explicit that its
conservative/base/aggressive multiples (2.50x / 3.00x / 3.50x) are "a market
judgment call, not derived from the business's own data" -- flagged, but
never actually checked against anything. This script closes that gap. It
pulls real, publicly published comps for small-business/specialty-retail
transactions -- not the illustrative business's own numbers, since these
are genuine third-party market statistics, not confidential financials --
and reports where the existing assumption actually falls relative to them.

DATA SOURCES (real, dated, cited -- not illustrative):

1. BizBuySell, "Business Valuation Multiples by Industry" -- 5-year trailing
   dataset of closed sales, Q3 2021 through Q2 2026.
   https://www.bizbuysell.com/learning-center/industry-valuation-multiples/
   - Retail sector (all retail categories, n = large): 2.63x SDE, $286,000
     median sale price.
   - Clothing & Accessory Stores (closest specific-category match to a
     boutique): 2.32x SDE, $200,000 median sale price.
   - Jewelry Stores (this business's single largest revenue category per
     its own item-level sales data): 2.13x SDE, $202,257 median sale price.

2. BizBuySell Q2 2026 Insight Report (quarterly, not the 5-year trailing
   table above -- directional, not an absolute multiple).
   https://www.bizbuysell.com/insight-report/
   - Retail sector Q2 2026: median sale price held at $250,000 despite
     softer revenue and cash flow; average cash-flow multiple up 6%
     year-over-year. Retail had the sharpest YoY decline in deal volume
     (-15%) of any sector tracked.

3. IBBA (International Business Brokers Association) & M&A Source,
   Market Pulse Survey, Q1 2026 (300 brokers, 203 closed transactions).
   https://www.prnewswire.com/news-releases/the-market-pulse-survey-q1-2026-reports-the-latest-trends-in-business-sales-up-to-50m-302813653.html
   - Median multiple by DEAL SIZE (not industry): 2.0x SDE under $500K,
     2.8x SDE from $500K-$1M, 3.0x SDE from $1M-$2M. This illustrative
     deal's purchase price ($509,926.80) falls just inside the $500K-$1M
     bracket -> 2.8x SDE is the size-matched comp.

None of these sources know anything about this specific (illustrative)
business -- they are general market statistics, which is exactly what a
judgment-call multiple should be checked against.
"""

import json
import os
from datetime import datetime

VALUATION_JSON_REL = os.path.join("dashboard_app", "mna_valuation_model.json")

MARKET_COMPS = [
    {
        "source": "BizBuySell -- Business Valuation Multiples by Industry (5-yr trailing, Q3 2021-Q2 2026)",
        "url": "https://www.bizbuysell.com/learning-center/industry-valuation-multiples/",
        "segment": "Retail sector (all categories)",
        "sdeMultiple": 2.63,
        "medianSalePrice": 286000,
        "basisForRelevance": "Broadest, most directly comparable sector-level figure.",
    },
    {
        "source": "BizBuySell -- Business Valuation Multiples by Industry (5-yr trailing, Q3 2021-Q2 2026)",
        "url": "https://www.bizbuysell.com/learning-center/industry-valuation-multiples/",
        "segment": "Clothing & Accessory Stores",
        "sdeMultiple": 2.32,
        "medianSalePrice": 200000,
        "basisForRelevance": "Closest BizBuySell subcategory to a specialty boutique.",
    },
    {
        "source": "BizBuySell -- Business Valuation Multiples by Industry (5-yr trailing, Q3 2021-Q2 2026)",
        "url": "https://www.bizbuysell.com/learning-center/industry-valuation-multiples/",
        "segment": "Jewelry Stores",
        "sdeMultiple": 2.13,
        "medianSalePrice": 202257,
        "basisForRelevance": (
            "This business's own item-level sales data (see best-sellers "
            "reporting) shows jewelry as the single largest unit-volume "
            "category -- the closest single-category comp to its actual mix."
        ),
    },
    {
        "source": "IBBA & M&A Source -- Market Pulse Survey, Q1 2026",
        "url": "https://www.prnewswire.com/news-releases/the-market-pulse-survey-q1-2026-reports-the-latest-trends-in-business-sales-up-to-50m-302813653.html",
        "segment": "Deal-size bracket $500K-$1M (this deal's $509,926.80 purchase price falls here)",
        "sdeMultiple": 2.8,
        "medianSalePrice": None,
        "basisForRelevance": "Size-matched comp, industry-agnostic -- deal size drives multiple as much as sector does at this scale.",
    },
]


def get_assumed_multiples(business_dir):
    """Read the valuation engine's own output rather than re-hardcoding the
    multiples here -- if mna_valuation_engine.py's assumptions ever change,
    this check should automatically compare against the new numbers, not a
    stale copy."""
    path = os.path.join(business_dir, VALUATION_JSON_REL)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"{path} not found -- run mna_valuation_engine.py first so this "
            "script has a real assumption to check against."
        )
    with open(path, "r", encoding="utf-8") as f:
        model = json.load(f)
    return model["multiples"], model["recastSDE"]


def benchmark_multiples():
    business_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"=== Market Comps Check [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ===")

    assumed_multiples, recast_sde = get_assumed_multiples(business_dir)

    comp_values = [c["sdeMultiple"] for c in MARKET_COMPS]
    comps_low = min(comp_values)
    comps_high = max(comp_values)
    comps_avg = round(sum(comp_values) / len(comp_values), 2)

    findings = []
    for label, multiple in assumed_multiples.items():
        if multiple < comps_low:
            read = f"below every cited comp ({comps_low}x-{comps_high}x) -- genuinely conservative"
        elif multiple <= comps_high:
            read = f"within the cited comps range ({comps_low}x-{comps_high}x)"
        else:
            read = (
                f"ABOVE every cited comp ({comps_low}x-{comps_high}x) -- not supported by "
                "these market sources; would need a deal-specific justification (embedded "
                "real estate, SBA-prequalification, diversified wholesale/Faire revenue, "
                "no owner-key-person risk) to defend using it as a base case rather than an "
                "upside case"
            )
        findings.append({
            "scenario": label,
            "assumedMultiple": multiple,
            "impliedValuation": round(multiple * recast_sde, 2),
            "readVsComps": read,
        })

    base_finding = next(f for f in findings if f["scenario"] == "base")
    conservative_finding = next(f for f in findings if f["scenario"] == "conservative")
    base_above_comps = base_finding["assumedMultiple"] > comps_high
    conservative_in_range = comps_low <= conservative_finding["assumedMultiple"] <= comps_high

    headline_finding = (
        f"The valuation engine's BASE CASE multiple ({base_finding['assumedMultiple']}x) "
        f"sits {'above every cited' if base_above_comps else 'within the cited'} market comp "
        f"({comps_low}x-{comps_high}x, average {comps_avg}x) for this deal's size and sector."
    )
    if base_above_comps and conservative_in_range:
        headline_finding += (
            f" The CONSERVATIVE multiple ({conservative_finding['assumedMultiple']}x) is the "
            "one that actually lands inside the comps range. Real market data centers closer "
            "to this deal's conservative case than its base case -- the base case should be "
            "treated as an upside scenario pending a specific, articulated reason this "
            "business should command a premium to comparable specialty retailers, not as the "
            "most likely outcome."
        )

    result = {
        "asOfDate": datetime.now().strftime("%B %d, %Y"),
        "recastSDE": recast_sde,
        "marketComps": MARKET_COMPS,
        "compsRange": {"low": comps_low, "high": comps_high, "average": comps_avg},
        "assumedMultipleFindings": findings,
        "headlineFinding": headline_finding,
    }

    out_dir = os.path.join(business_dir, "dashboard_app")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "market_comps_check.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)

    print(f"Comps range: {comps_low}x - {comps_high}x SDE (avg {comps_avg}x) across {len(MARKET_COMPS)} sourced comps")
    for finding in findings:
        print(f"  {finding['scenario']:>12}: {finding['assumedMultiple']}x -> {finding['readVsComps']}")
    print(f"[SUCCESS] Market comps check written -> {out_file}")
    return result


if __name__ == "__main__":
    benchmark_multiples()
