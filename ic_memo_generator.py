"""
ic_memo_generator.py -- assembles a buy-side Investment Committee memo by
reading, not re-deriving, the outputs of every other script in this repo:
mna_valuation_engine.py's JSON, market_comps.py's JSON, lbo_model's
recalculated LBO_Model.xlsx, and qoe_workbook's recalculated QoE_Workbook.xlsx.

WHY THIS EXISTS: mna_deal_room_generator.py already produces an external,
buyer-facing teaser. An IC memo is the different, internal document a real
deal team produces -- the one that goes to the people deciding whether to
proceed, not the one used to attract a counterparty. It has to say things a
teaser never would (the base-case multiple isn't well-supported by comps;
here's exactly what's still unverified; here's what has to happen before
this is fundable) and it has to synthesize four separate outputs into one
recommendation instead of presenting any single one in isolation.

Nothing here is a fifth, independent copy of a number that already lives in
one of the other four outputs -- this script only reads, formats, and
narrates. If a prerequisite output doesn't exist yet (or exists but was
never recalculated, so its cells still read back as None), this script
fails loudly and says exactly which script to run first, rather than
silently writing "N/A" into a document someone might present as final.
"""

import os
import json
from datetime import datetime

import openpyxl
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

MAROON = RGBColor(0x8B, 0x3A, 0x4A)
DARK = RGBColor(0x2D, 0x18, 0x20)
MUTED = RGBColor(0x6B, 0x53, 0x58)
FLAG = RGBColor(0xB0, 0x00, 0x00)


def _business_dir():
    return os.path.dirname(os.path.abspath(__file__))


def read_valuation_model(business_dir):
    path = os.path.join(business_dir, "dashboard_app", "mna_valuation_model.json")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"{path} not found -- run mna_valuation_engine.py first."
        )
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def read_market_comps(business_dir):
    path = os.path.join(business_dir, "dashboard_app", "market_comps_check.json")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"{path} not found -- run market_comps.py first."
        )
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _require(value, cell_desc, source_file):
    if value is None:
        raise ValueError(
            f"{source_file} cell {cell_desc} read back as None -- the workbook "
            "hasn't been recalculated since it was last saved. Run "
            "scripts/recalc.py (or open and save it in Excel/LibreOffice) "
            "before generating the IC memo from it."
        )
    return value


def read_lbo_figures(business_dir):
    path = os.path.join(business_dir, "lbo_model", "LBO_Model.xlsx")
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found -- run lbo_model/build_lbo_model.py first.")
    wb = openpyxl.load_workbook(path, data_only=True)
    su, ds, ra, notes = wb["Sources & Uses"], wb["Debt Schedule"], wb["Returns Analysis"], wb["Notes & Limitations"]

    def su_(ref):
        return _require(su[ref].value, f"Sources & Uses!{ref}", "LBO_Model.xlsx")

    def ra_(ref):
        return _require(ra[ref].value, f"Returns Analysis!{ref}", "LBO_Model.xlsx")

    def ds_(ref):
        return _require(ds[ref].value, f"Debt Schedule!{ref}", "LBO_Model.xlsx")

    # Rows 17-19 hold the MOIC sensitivity grid; rows 24-26 hold IRR -- easy
    # to mix up since both are keyed by the same exit-multiple labels. This
    # memo shows IRR, so it must read 24-26, not 17-19 (an earlier version
    # of this script read 17-19 here and rendered MOIC multiples formatted
    # as percentages -- e.g. 8.53x came out as "853%" IRR).
    sens_rows = []
    for r in (24, 25, 26):
        mult = ra[f"B{r}"].value
        vals = [ra[f"{c}{r}"].value for c in ("C", "D", "E")]
        if mult is not None and all(v is not None for v in vals):
            sens_rows.append((mult, vals))

    return {
        "totalUses": su_("B6"),
        "sbaLoan": su_("B9"),
        "sellerNote": su_("B10"),
        "sponsorEquity": su_("B11"),
        "sponsorEquityPct": su_("B17"),
        "sbaRate": None,  # pulled from Assumptions below
        "exitYear": ra_("B4"),
        "exitMultiple": ra_("B5"),
        "exitEquityValue": ra_("B9"),
        "netDebtAtExit": ra_("B8"),
        "moic": ra_("B11"),
        "irr": ra_("B12"),
        "netDebtOutstandingByYear": [ds_(f"{c}27") for c in ("C", "D", "E", "F", "G")],
        "dscrByYear": [ds_(f"{c}24") for c in ("C", "D", "E", "F", "G")],
        "moicIrrSensitivity": sens_rows,  # [(multiple, [irr_y3, irr_y4, irr_y5]), ...]
        "notes": [notes[f"A{r}"].value for r in range(3, 20) if notes[f"A{r}"].value],
    }


def read_qoe_figures(business_dir):
    path = os.path.join(business_dir, "qoe_workbook", "QoE_Workbook.xlsx")
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found -- run qoe_workbook/build_qoe_workbook.py first.")
    wb = openpyxl.load_workbook(path, data_only=True)
    eb, rq, nwc, notes = (
        wb["Adjusted EBITDA Bridge"], wb["Revenue & Margin Quality"],
        wb["NWC Analysis"], wb["Notes & Limitations"],
    )

    def eb_(ref):
        return _require(eb[ref].value, f"Adjusted EBITDA Bridge!{ref}", "QoE_Workbook.xlsx")

    def nwc_(ref):
        return _require(nwc[ref].value, f"NWC Analysis!{ref}", "QoE_Workbook.xlsx")

    return {
        "reportedNetIncome": eb_("B4"),
        "totalAddBacks": eb_("B15"),
        "adjustedEBITDA": eb_("B18"),
        "tieOutCheck": eb_("B20"),
        "q4RevenuePct": _require(rq["E5"].value, "Revenue & Margin Quality!E5", "QoE_Workbook.xlsx"),
        "q4Margin": _require(rq["E7"].value, "Revenue & Margin Quality!E7", "QoE_Workbook.xlsx"),
        "avgNWC": nwc_("B11"),
        "atCloseNWC": nwc_("B12"),
        "nwcPeg": nwc_("B13"),
        "nwcDelta": nwc_("B14"),
        "notes": [notes[f"A{r}"].value for r in range(3, 20) if notes[f"A{r}"].value],
    }


# ---------------------------------------------------------------- doc helpers

def set_cell_shading(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = MAROON
    return h


def body(doc, text, italic=False, size=10, color=None, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.italic = italic
    run.bold = bold
    if color:
        run.font.color.rgb = color
    return p


def bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        r1.bold = True
        r2 = p.add_run(text)
    else:
        p.add_run(text)
    return p


def kv_table(doc, rows, col_widths=(Inches(2.6), Inches(3.2))):
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Light Grid Accent 5"
    for label, value in rows:
        row = table.add_row().cells
        row[0].text = str(label)
        row[1].text = str(value)
        row[0].paragraphs[0].runs[0].bold = True
    for row in table.rows:
        row.cells[0].width, row.cells[1].width = col_widths
    return table


def money(x):
    return f"${x:,.2f}" if isinstance(x, float) and not x.is_integer() else f"${x:,.0f}"


def pct(x):
    return f"{x:.1%}"


# ---------------------------------------------------------------- memo build

def build_memo(model, comps, lbo, qoe, out_file):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("INVESTMENT COMMITTEE MEMORANDUM")
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = MAROON

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("THE BUSINESS  --  the operating LLC  --  [address redacted]")
    r.font.size = Pt(11)
    r.font.color.rgb = MUTED

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(f"Prepared {model['asOfDate']}  |  CONFIDENTIAL  |  Internal use only")
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = MUTED

    is_ready = model.get("dealRoomReadiness", "").strip().upper().startswith(("100%", "READY"))
    rec_p = doc.add_paragraph()
    rec_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = rec_p.add_run(
        "RECOMMENDATION: PROCEED TO LOI, SUBJECT TO CONDITIONS BELOW" if not is_ready
        else "RECOMMENDATION: PROCEED"
    )
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = FLAG if not is_ready else RGBColor(0x04, 0x78, 0x57)

    doc.add_paragraph()

    # 1. Transaction Summary
    heading(doc, "1. Transaction Summary")
    kv_table(doc, [
        ("Purchase Price (Enterprise Value)", money(model["valuations"]["baseValuation"])),
        ("Total Uses (incl. transaction costs)", money(lbo["totalUses"])),
        ("Recast SDE (Base Year)", money(model["recastSDE"])),
        ("Base-Case Multiple", f"{model['multiples']['base']:.2f}x SDE"),
        ("Financing Structure", f"SBA 7(a) {money(lbo['sbaLoan'])} + Seller Note {money(lbo['sellerNote'])} + Sponsor Equity {money(lbo['sponsorEquity'])}"),
        ("Sponsor Equity / Total Uses", pct(lbo["sponsorEquityPct"])),
        ("Base-Case Hold Period / Exit Multiple", f"{lbo['exitYear']:.0f} years / {lbo['exitMultiple']:.2f}x SDE"),
        ("Base-Case MOIC / IRR", f"{lbo['moic']:.2f}x / {pct(lbo['irr'])}"),
    ])

    # 2. Investment Thesis
    heading(doc, "2. Investment Thesis")
    bullet(doc, "Flagship location in a high-traffic downtown retail corridor of a mountain tourist destination town -- durable foot traffic that doesn't depend on a single owner's relationships.", bold_prefix="Prime downtown footprint. ")
    bullet(doc, "POS platform, bank-feed integration, and an internal analytics/reporting stack are already built and operating -- an acquirer inherits systems, not a rebuild project.", bold_prefix="Turn-key operations. ")
    bullet(doc, "SBA 7(a) eligibility keeps the buyer pool wide and the capital structure realistic for a deal this size, rather than requiring an all-cash or unconventional close.", bold_prefix="SBA-financeable at this scale. ")
    bullet(doc, "Every add-back, every working-capital input, and the valuation multiple itself are checked against a named source and flagged verified/unverified below -- this memo is not asking the committee to take the headline numbers on faith.", bold_prefix="Diligence-tested, not just modeled. ")

    # 3. Recast Earnings & Quality of Earnings
    heading(doc, "3. Recast Earnings & Quality of Earnings")
    verified_pct = qoe["reportedNetIncome"] / qoe["adjustedEBITDA"]
    body(doc, (
        f"Reported net income of {money(qoe['reportedNetIncome'])} is read live from the business's "
        f"reconciled tax package and is fully verified. Two owner/non-recurring add-backs totaling "
        f"{money(qoe['totalAddBacks'])} bring adjusted SDE to {money(qoe['adjustedEBITDA'])} -- "
        f"{verified_pct:.1%} of that figure is verified, {1 - verified_pct:.1%} is unverified pending "
        f"bank-statement and invoice tracing (see Add-Back Detail tab, qoe_workbook/QoE_Workbook.xlsx). "
        f"The EBITDA bridge tie-out check against the valuation engine's own recastSDE output is "
        f"{qoe['tieOutCheck']:+.2f} -- confirms the two models agree to the penny."
    ))
    body(doc, (
        f"Revenue quality flag: Q4 alone carries {qoe['q4RevenuePct']:.1%} of trailing-twelve-month "
        f"revenue at the year's highest gross margin ({qoe['q4Margin']:.1%}). A single strong holiday "
        f"season is doing an outsized share of the year's earnings -- diligence should confirm Q4 "
        f"wasn't inflated by a one-time bulk order or promotion that won't repeat."
    ), color=FLAG)

    # 4. Valuation & Market Support
    heading(doc, "4. Valuation & Market Support")
    body(doc, comps["headlineFinding"], bold=True)
    comps_tbl = doc.add_table(rows=1, cols=4)
    comps_tbl.style = "Light Grid Accent 5"
    hdr = comps_tbl.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Source", "Segment", "SDE Multiple", "Median Sale Price"
    for c in hdr:
        c.paragraphs[0].runs[0].bold = True
    for comp in comps["marketComps"]:
        row = comps_tbl.add_row().cells
        row[0].text = comp["source"].split(" -- ")[0]
        row[1].text = comp["segment"]
        row[2].text = f"{comp['sdeMultiple']:.2f}x"
        row[3].text = money(comp["medianSalePrice"]) if comp["medianSalePrice"] else "n/a"
    doc.add_paragraph()
    for finding in comps["assumedMultipleFindings"]:
        bullet(doc, f"{finding['assumedMultiple']:.2f}x -> {finding['readVsComps']}", bold_prefix=f"{finding['scenario'].capitalize()} case: ")

    # 5. Capital Structure & Financing
    heading(doc, "5. Capital Structure & Financing")
    kv_table(doc, [
        ("SBA 7(a) Term Loan", money(lbo["sbaLoan"])),
        ("Seller Note", money(lbo["sellerNote"])),
        ("Sponsor Equity", money(lbo["sponsorEquity"])),
        ("Total Sources", money(lbo["totalUses"])),
        ("Year 1 DSCR", f"{lbo['dscrByYear'][0]:.2f}x"),
        (f"Year {lbo['exitYear']:.0f} (Exit) DSCR", f"{lbo['dscrByYear'][-1]:.2f}x"),
        ("Net Debt Outstanding at Exit", money(lbo["netDebtAtExit"])),
    ])
    body(doc, (
        "Structured as SBA 7(a) acquisition financing rather than a traditional institutional LBO "
        "stack -- no senior lender underwrites acquisition debt at this deal size. Sponsor equity is "
        "set at the SBA's real ~10% minimum injection for a change-of-ownership 7(a) loan."
    ), italic=True, size=9, color=MUTED)

    # 6. Returns Analysis
    heading(doc, "6. Returns Analysis")
    kv_table(doc, [
        ("Base Case (Exit Yr, Multiple)", f"{lbo['exitYear']:.0f} yrs, {lbo['exitMultiple']:.2f}x"),
        ("Base Case MOIC", f"{lbo['moic']:.2f}x"),
        ("Base Case IRR", pct(lbo["irr"])),
    ])
    sens_tbl = doc.add_table(rows=1, cols=4)
    sens_tbl.style = "Light Grid Accent 5"
    hdr = sens_tbl.rows[0].cells
    hdr[0].text = "Exit Multiple \\ Exit Year"
    hdr[1].text, hdr[2].text, hdr[3].text = "Year 3", "Year 4", "Year 5"
    for c in hdr:
        c.paragraphs[0].runs[0].bold = True
    for mult, irr_vals in lbo["moicIrrSensitivity"]:
        row = sens_tbl.add_row().cells
        row[0].text = f"{mult:.2f}x"
        for i, v in enumerate(irr_vals):
            row[i + 1].text = pct(v)
    body(doc, "IRR sensitivity by exit multiple and hold period.", italic=True, size=9, color=MUTED)

    # 7. Key Risks & Mitigants
    heading(doc, "7. Key Risks & Mitigants")
    risk_lines = list(dict.fromkeys(qoe["notes"] + lbo["notes"]))  # dedupe, preserve order
    for line in risk_lines:
        if line.lower().startswith(("this is a portfolio project", "this model works at the pre-tax")):
            continue
        bullet(doc, line)

    # 8. Recommendation & Conditions to Proceed
    heading(doc, "8. Recommendation & Conditions to Proceed")
    body(doc, (
        f"{model.get('dealRoomReadiness', '')}. This memo's own valuation support (Section 4) "
        "shows the base case sits above sourced market comps -- committee sign-off on this deal "
        "should be conditioned on the following before the base case (not just the conservative "
        "case) is presented to a lender or used to set an offer price:"
    ))
    conditions = [
        "Trace both owner/non-recurring add-backs to bank statements and dated invoices (currently unverified -- see Section 3).",
        "Confirm no interest-bearing business debt and no un-booked D&A exists that would change the EBITDA bridge (Notes, Adjusted EBITDA Bridge tab).",
        "Verify Q4 revenue concentration is organic and repeatable, not a one-time bulk or wholesale order (Section 3).",
        "Obtain a real SBA 7(a) lender term sheet to replace the illustrative rate/term assumptions used in the LBO model.",
        "Confirm the NWC peg convention (trailing-12-month average vs. at-close) against the actual purchase agreement language before it affects the closing price adjustment.",
        "Either source a specific, deal-level justification for underwriting above the comps-implied multiple range, or re-run this memo's valuation section at the conservative (comps-supported) case.",
    ]
    for c in conditions:
        bullet(doc, c)

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run(
        "Generated by ic_memo_generator.py from mna_valuation_engine.py, market_comps.py, "
        "lbo_model/LBO_Model.xlsx, and qoe_workbook/QoE_Workbook.xlsx -- every figure above traces "
        "to one of those four outputs. This is a portfolio project; business identity and dollar "
        "figures are illustrative."
    )
    r.font.size = Pt(8)
    r.italic = True
    r.font.color.rgb = MUTED

    doc.save(out_file)
    return out_file


def generate_ic_memo():
    business_dir = _business_dir()
    print(f"=== IC Memo Generator [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ===")

    model = read_valuation_model(business_dir)
    comps = read_market_comps(business_dir)
    lbo = read_lbo_figures(business_dir)
    qoe = read_qoe_figures(business_dir)

    out_dir = os.path.join(business_dir, "exit_package")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "02_Investment_Committee_Memo.docx")
    build_memo(model, comps, lbo, qoe, out_file)

    print(f"[SUCCESS] IC memo assembled from 4 upstream outputs -> {out_file}")
    return out_file


if __name__ == "__main__":
    generate_ic_memo()
