"""
build_qoe_workbook.py -- generates QoE_Workbook.xlsx, a buy-side Quality of
Earnings workbook for the SAME illustrative deal mna_valuation_engine.py and
lbo_model/build_lbo_model.py already value (recast SDE $169,975.60, reported
net income $148,220.60, two owner/non-recurring add-backs totaling
$21,755.00).

WHY THIS EXISTS ALONGSIDE THE VALUATION ENGINE: a recast-SDE valuation
answers "what is this worth if the earnings are real." A QoE workbook is the
diligence step that tests whether they're real -- it breaks the same annual
add-backs down to monthly/source-document granularity, checks whether
revenue and margin are trending or lumpy, and normalizes working capital to
set a fair closing-date peg. A buyer (or a seller's advisor preparing for
buyer diligence) builds this BEFORE relying on the sell-side number, not
after. This workbook is the "test the earnings" counterpart to the "price
the earnings" work mna_valuation_engine.py already does.

Every hardcoded figure is either sourced from the SAME verified/unverified
figures the valuation engine and LBO model already carry, or is a new
illustrative assumption flagged the same way -- consistent with this repo's
verified/unverified labeling discipline throughout.
"""

import os

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(OUT_DIR, "QoE_Workbook.xlsx")

FONT_NAME = "Arial"
BLUE = Font(name=FONT_NAME, color="0000FF")
BLACK = Font(name=FONT_NAME, color="000000")
BOLD = Font(name=FONT_NAME, bold=True)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
HEADER_FILL = PatternFill(start_color="8B3A4A", end_color="8B3A4A", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
SECTION_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
FLAG_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")

CUR = '$#,##0;($#,##0);"-"'
PCT = "0.0%"
MULT = "0.00x"

MONTH_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
QTR_COLS = ["B", "C", "D", "E"]


def set_cell(ws, ref, value, font=BLACK, fmt=None, bold=False, note=None, fill=None):
    cell = ws[ref]
    cell.value = value
    cell.font = Font(name=FONT_NAME, color=font.color, bold=bold) if font else Font(name=FONT_NAME, bold=bold)
    if fmt:
        cell.number_format = fmt
    if note:
        cell.comment = Comment(note, "build_qoe_workbook.py")
    if fill:
        cell.fill = fill
    return cell


def section(ws, ref, text):
    cell = ws[ref]
    cell.value = text
    cell.font = BOLD
    cell.fill = SECTION_FILL


def title(ws, text):
    ws["A1"] = text
    ws["A1"].font = TITLE_FONT


def widen(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def header_row(ws, row, labels, start_col="B"):
    cols = MONTH_COLS if len(labels) == 12 else QTR_COLS
    for col, label in zip(cols, labels):
        cell = ws[f"{col}{row}"]
        cell.value = label
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def build_overview(wb):
    ws = wb.create_sheet("Overview")
    title(ws, "Quality of Earnings Workbook")
    widen(ws, {"A": 100})

    lines = [
        "Illustrative buy-side QoE analysis for the same anonymized deal valued in "
        "mna_valuation_engine.py and financed in lbo_model/LBO_Model.xlsx -- reported "
        "net income $148,220.60, two owner/non-recurring add-backs totaling $21,755.00, "
        "recast SDE $169,975.60.",
        "",
        "Scope: Adjusted EBITDA bridge (with monthly add-back detail, not just the annual "
        "total), quarterly revenue and gross margin trend, and net working capital "
        "normalization with a proposed closing peg. This is the standard first pass of a "
        "buy-side or sell-side-readiness QoE for a deal this size -- it does not include a "
        "full independent audit, customer/vendor confirmation letters, or a formal tax "
        "opinion, all of which a real engagement would add.",
        "",
        "Every add-back and every NWC input below is flagged verified or unverified against "
        "a specific source, the same discipline mna_valuation_engine.py established. Nothing "
        "here should be presented to a lender, buyer, or seller as final until every "
        "unverified item is traced to source documents (bank statements, POS exports, AP/AR "
        "aging) -- see Notes & Limitations.",
    ]
    for i, line in enumerate(lines, start=3):
        cell = ws[f"A{i}"]
        cell.value = line
        cell.font = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    section(ws, "A12", "Contents")
    contents = [
        "Adjusted EBITDA Bridge -- annual bridge from reported net income to adjusted SDE, "
        "reconciled to mna_valuation_engine.py",
        "Add-Back Detail -- the two annual add-backs broken into monthly/source-traceable line items",
        "Revenue & Margin Quality -- quarterly revenue and gross margin trend, seasonality, and flags",
        "NWC Analysis -- trailing-12-month monthly working capital, average vs. at-close, proposed peg",
        "Notes & Limitations",
    ]
    for i, line in enumerate(contents, start=13):
        cell = ws[f"A{i}"]
        cell.value = f"- {line}"
        cell.font = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    return ws


def build_ebitda_bridge(wb):
    ws = wb.create_sheet("Adjusted EBITDA Bridge")
    title(ws, "Adjusted EBITDA Bridge (TTM)")
    widen(ws, {"A": 48, "B": 16, "C": 40})

    section(ws, "A3", "Starting Point")
    set_cell(ws, "A4", "Reported Net Income")
    set_cell(ws, "B4", 148220.60, font=BLUE, fmt=CUR,
             note="Source: mna_valuation_engine.py get_verified_net_income() -- live from "
                  "business_tax_package_latest.xlsx, Tab 2 B59. Verified.")
    set_cell(ws, "C4", "Verified -- traces to reconciled tax package")

    section(ws, "A6", "Above-the-Line Adjustments (to get from Net Income to EBITDA)")
    set_cell(ws, "A7", "+ Interest Expense")
    set_cell(ws, "B7", 0, font=BLUE, fmt=CUR,
             note="No interest-bearing business debt identified in the tax package as of "
                  "this engagement -- confirm against the balance sheet before assuming $0.")
    set_cell(ws, "C7", "Unverified -- confirm no interest-bearing debt on balance sheet")
    set_cell(ws, "A8", "+ Income Taxes")
    set_cell(ws, "B8", 0, font=BLUE, fmt=CUR,
             note="Pass-through entity (net income already reported pre-entity-level-tax on "
                  "the owner's return) -- consistent with SDE convention, not a true GAAP "
                  "EBITDA bridge. See Notes & Limitations.")
    set_cell(ws, "C8", "N/A -- pass-through entity, see Notes")
    set_cell(ws, "A9", "+ Depreciation & Amortization")
    set_cell(ws, "B9", 0, font=BLUE, fmt=CUR,
             note="No D&A separately broken out in the tax package income statement as "
                  "provided -- if fixed-asset depreciation exists, it is already embedded in "
                  "reported net income and NOT separately added back here. Confirm against "
                  "the depreciation schedule/Form 4562 before relying on this.")
    set_cell(ws, "C9", "Unverified -- confirm no embedded D&A being missed")
    set_cell(ws, "A10", "EBITDA (as reported)", bold=True)
    set_cell(ws, "B10", "=SUM(B4:B9)", fmt=CUR, bold=True)

    section(ws, "A12", "Normalizing Add-Backs (owner/non-recurring)")
    set_cell(ws, "A13", "+ Owner Discretionary Draw & Health Benefits")
    set_cell(ws, "B13", 20140.00, font=BLUE, fmt=CUR,
             note="Source: mna_valuation_engine.py addBackBreakdown. Not yet traced to "
                  "bank/ledger detail at the annual level -- see Add-Back Detail tab for the "
                  "monthly breakdown attempted here.")
    set_cell(ws, "C13", "Unverified -- see Add-Back Detail tab")
    set_cell(ws, "A14", "+ One-Time Store Improvement (Wayfair Fixtures)")
    set_cell(ws, "B14", 1615.00, font=BLUE, fmt=CUR,
             note="Source: mna_valuation_engine.py addBackBreakdown. Not yet traced to "
                  "bank/ledger detail -- see Add-Back Detail tab.")
    set_cell(ws, "C14", "Unverified -- see Add-Back Detail tab")
    set_cell(ws, "A15", "Total Normalizing Add-Backs", bold=True)
    set_cell(ws, "B15", "=SUM(B13:B14)", fmt=CUR, bold=True)

    section(ws, "A17", "Adjusted EBITDA / SDE")
    set_cell(ws, "A18", "Adjusted EBITDA (= Recast SDE)", bold=True)
    set_cell(ws, "B18", "=B10+B15", fmt=CUR, bold=True)
    set_cell(ws, "A19", "Reconciles to mna_valuation_engine.py recastSDE")
    set_cell(ws, "B19", 169975.60, font=BLUE, fmt=CUR,
             note="Cross-check figure, not an independent input -- pulled from the same "
                  "engine's output for the tie-out check in B20.")
    set_cell(ws, "A20", "Tie-out check (should be 0)")
    set_cell(ws, "B20", "=B18-B19", fmt=CUR)

    section(ws, "A22", "Quality Read")
    set_cell(ws, "A23",
             "Of the $169,975.60 adjusted figure, $148,220.60 (87.2%) is verified reported "
             "net income and $21,755.00 (12.8%) is unverified add-backs. A buyer should treat "
             "the unverified 12.8% as the highest-risk slice of the number until it's traced "
             "to source documents -- see Add-Back Detail.")
    ws["A23"].font = Font(name=FONT_NAME, size=10, italic=True)
    ws["A23"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A23:C23")
    ws.row_dimensions[23].height = 48

    return ws


def build_addback_detail(wb):
    ws = wb.create_sheet("Add-Back Detail")
    title(ws, "Add-Back Detail -- Monthly Trace Attempt")
    widen(ws, {"A": 40})
    for col in MONTH_COLS:
        widen(ws, {col: 11})

    set_cell(ws, "A2",
             "Illustrative monthly allocation of the two annual add-backs from "
             "mna_valuation_engine.py. Real owner draws and one-time purchases don't land "
             "evenly across the year -- this tab exists to show the kind of month-by-month "
             "tracing a real QoE would do against bank/ledger detail, not to claim these "
             "specific monthly splits are themselves verified.",
             fmt=None)
    ws["A2"].font = Font(name=FONT_NAME, size=9, italic=True)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:M2")
    ws.row_dimensions[2].height = 30

    header_row(ws, 4, MONTH_NAMES)
    set_cell(ws, "A4", "Owner Discretionary Draw & Health Benefits", bold=True)
    ws["A4"].font = HEADER_FONT
    ws["A4"].fill = HEADER_FILL

    # Owner draw: irregular, higher in months with larger discretionary spend
    # (illustrative -- sums to exactly $20,140.00, the annual figure).
    draw_vals = [1200, 900, 1800, 1500, 2400, 1650, 1200, 3200, 1450, 1100, 1240, 2500]
    assert sum(draw_vals) == 20140.00, "monthly draw allocation must sum to the annual add-back"
    for col, v in zip(MONTH_COLS, draw_vals):
        set_cell(ws, f"{col}5", v, font=BLUE, fmt=CUR)
    set_cell(ws, "A5", "  Monthly draw amount")
    set_cell(ws, "A6", "  Running total (should hit $20,140.00 at Dec)")
    set_cell(ws, "B6", "=SUM($B$5:B5)", fmt=CUR)
    for i, col in enumerate(MONTH_COLS[1:], start=1):
        prev = MONTH_COLS[i - 1]
        set_cell(ws, f"{col}6", f"=SUM($B$5:{col}5)", fmt=CUR)
    set_cell(ws, "A7", "  Verified against bank statement?")
    for col in MONTH_COLS:
        set_cell(ws, f"{col}7", "No", font=Font(name=FONT_NAME, italic=True, color="B00000"))

    set_cell(ws, "A9", "One-Time Store Improvement (Wayfair Fixtures)", bold=True)
    ws["A9"].font = HEADER_FONT
    ws["A9"].fill = HEADER_FILL
    fixture_vals = [0, 0, 0, 0, 0, 1615.00, 0, 0, 0, 0, 0, 0]
    for col, v in zip(MONTH_COLS, fixture_vals):
        set_cell(ws, f"{col}10", v, font=BLUE, fmt=CUR)
    set_cell(ws, "A10", "  Monthly amount")
    set_cell(ws, "A11", "  Verified against invoice/receipt?")
    for col, v in zip(MONTH_COLS, fixture_vals):
        label = "Unverified -- invoice not yet located" if v else "N/A"
        set_cell(ws, f"{col}11", label, font=Font(name=FONT_NAME, italic=True, size=8))

    section(ws, "A13", "Reconciliation")
    set_cell(ws, "A14", "Sum of monthly draw allocations")
    set_cell(ws, "B14", "=SUM(B5:M5)", fmt=CUR)
    set_cell(ws, "A15", "Annual figure per mna_valuation_engine.py")
    set_cell(ws, "B15", 20140.00, font=BLUE, fmt=CUR)
    set_cell(ws, "A16", "Check (should be 0)")
    set_cell(ws, "B16", "=B14-B15", fmt=CUR)

    return ws


def build_revenue_quality(wb):
    ws = wb.create_sheet("Revenue & Margin Quality")
    title(ws, "Revenue & Gross Margin Quality (TTM by Quarter)")
    widen(ws, {"A": 34})
    for col in QTR_COLS:
        widen(ws, {col: 14})
    widen(ws, {"F": 16})

    header_row(ws, 3, ["Q1 (Jan-Mar)", "Q2 (Apr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dec)"])
    set_cell(ws, "F3", "TTM Total", bold=True)
    ws["F3"].font = HEADER_FONT
    ws["F3"].fill = HEADER_FILL

    set_cell(ws, "A4", "Revenue")
    rev = [280000, 322000, 308000, 490000]
    assert sum(rev) == 1400000, "quarterly revenue must sum to the Year 0 LTM revenue used elsewhere"
    for col, v in zip(QTR_COLS, rev):
        set_cell(ws, f"{col}4", v, font=BLUE, fmt=CUR,
                 note="Illustrative quarterly split of the Year 0 LTM revenue used in "
                      "lbo_model's Assumptions tab -- specialty-retail seasonality "
                      "(post-holiday Q1 trough, Q4 holiday peak)." if col == "B" else None)
    set_cell(ws, "F4", "=SUM(B4:E4)", fmt=CUR, bold=True)

    set_cell(ws, "A5", "Revenue, % of TTM")
    for col in QTR_COLS:
        set_cell(ws, f"{col}5", f"={col}4/$F$4", fmt=PCT)

    set_cell(ws, "A7", "Gross Margin %")
    gm = [0.47, 0.50, 0.51, 0.53]
    for col, v in zip(QTR_COLS, gm):
        set_cell(ws, f"{col}7", v, font=BLUE, fmt=PCT,
                 note="Illustrative -- Q1 reflects post-holiday clearance markdowns; Q4 "
                      "reflects full-price holiday selling. Confirm against actual COGS by "
                      "month before relying on this pattern." if col == "B" else None)

    set_cell(ws, "A8", "Gross Profit ($)")
    for col in QTR_COLS:
        set_cell(ws, f"{col}8", f"={col}4*{col}7", fmt=CUR)
    set_cell(ws, "F8", "=SUM(B8:E8)", fmt=CUR, bold=True)

    set_cell(ws, "A10", "Blended TTM Gross Margin %")
    set_cell(ws, "B10", "=F8/F4", fmt=PCT)

    section(ws, "A12", "Flags")
    set_cell(ws, "A13",
             "Q4 carries 35.0% of TTM revenue at the highest quarterly margin (53.0%) -- a "
             "single strong holiday season is doing an outsized share of the year's earnings. "
             "A buyer should pressure-test Q4 specifically (was there a one-time bulk/wholesale "
             "order, a new account, or a promotion that won't repeat?) rather than assuming "
             "next year's Q4 repeats at the same scale.",
             fill=FLAG_FILL)
    ws["A13"].font = Font(name=FONT_NAME, size=10)
    ws["A13"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A13:F13")
    ws.row_dimensions[13].height = 48

    return ws


def build_nwc_analysis(wb):
    ws = wb.create_sheet("NWC Analysis")
    title(ws, "Net Working Capital Analysis (Trailing 12 Months)")
    widen(ws, {"A": 40})
    for col in MONTH_COLS:
        widen(ws, {col: 11})
    widen(ws, {"N": 14})

    header_row(ws, 3, MONTH_NAMES)
    set_cell(ws, "N3", "Avg / Total", bold=True)
    ws["N3"].font = HEADER_FONT
    ws["N3"].fill = HEADER_FILL

    set_cell(ws, "A4", "Accounts Receivable (Faire wholesale)")
    ar = [4200, 3800, 5100, 6200, 7400, 8100, 6900, 5600, 6300, 9800, 12400, 7100]
    for col, v in zip(MONTH_COLS, ar):
        set_cell(ws, f"{col}4", v, font=BLUE, fmt=CUR,
                 note="Illustrative -- most sales are POS/card at point of sale with no "
                      "receivable; AR here is wholesale (Faire) invoices outstanding. "
                      "Confirm against the Faire payment tracker and AR aging report." if col == "B" else None)

    set_cell(ws, "A5", "Inventory (at cost)")
    inv = [142000, 138000, 133000, 137000, 145000, 150000, 148000, 152000, 161000, 178000, 186000, 151000]
    for col, v in zip(MONTH_COLS, inv):
        set_cell(ws, f"{col}5", v, font=BLUE, fmt=CUR,
                 note="Illustrative -- builds ahead of Q4 holiday selling, sells through in "
                      "December. Confirm against physical/cycle count, not just the perpetual "
                      "POS inventory value, before relying on this." if col == "B" else None)

    set_cell(ws, "A6", "Accounts Payable (trade)")
    ap = [38000, 35000, 33000, 36000, 41000, 44000, 42000, 45000, 52000, 61000, 58000, 40000]
    for col, v in zip(MONTH_COLS, ap):
        set_cell(ws, f"{col}6", v, font=BLUE, fmt=CUR,
                 note="Illustrative -- follows inventory purchasing with roughly a "
                      "one-month lag. Confirm against actual vendor AP aging." if col == "B" else None)

    set_cell(ws, "A8", "Net Working Capital (AR + Inventory - AP)", bold=True)
    for col in MONTH_COLS:
        set_cell(ws, f"{col}8", f"={col}4+{col}5-{col}6", fmt=CUR, bold=True)

    section(ws, "A10", "Peg Calculation")
    set_cell(ws, "A11", "Trailing 12-Month Average NWC")
    set_cell(ws, "B11", "=AVERAGE(B8:M8)", fmt=CUR)
    set_cell(ws, "A12", "At-Close NWC (most recent month, Dec)")
    set_cell(ws, "B12", "=M8", fmt=CUR)
    set_cell(ws, "A13", "Proposed NWC Peg (12-month average)", bold=True)
    set_cell(ws, "B13", "=B11", fmt=CUR, bold=True,
             note="Standard convention: peg to a trailing average, not a single point-in-"
                  "time balance, so neither party is rewarded for timing the close around a "
                  "seasonal low or high. Confirm this convention against the actual purchase "
                  "agreement language before relying on it.")
    set_cell(ws, "A14", "Delta: At-Close vs. Peg (positive = seller delivers extra NWC)")
    set_cell(ws, "B14", "=B12-B13", fmt=CUR)
    set_cell(ws, "A15", "Implied Purchase Price Adjustment")
    set_cell(ws, "B15", "=B14", fmt=CUR,
             note="Mechanically, a positive delta increases the price paid at close (seller "
                  "delivered more working capital than the peg); a negative delta decreases "
                  "it. Sign convention should be confirmed against the actual purchase "
                  "agreement before use.")

    section(ws, "A17", "Flags")
    set_cell(ws, "A18",
             "December's NWC is a poor stand-in for a 'normal' month -- inventory is still "
             "elevated from holiday buying and hasn't fully sold through. If closing lands "
             "near year-end, pegging to the trailing-12-month average (rather than the "
             "at-close balance) meaningfully protects the buyer from overpaying for seasonal "
             "inventory that's about to be marked down.",
             fill=FLAG_FILL)
    ws["A18"].font = Font(name=FONT_NAME, size=10)
    ws["A18"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A18:N18")
    ws.row_dimensions[18].height = 48

    return ws


def build_notes(wb):
    ws = wb.create_sheet("Notes & Limitations")
    title(ws, "Notes & Limitations")
    widen(ws, {"A": 100})
    notes = [
        "This is an SDE-convention bridge (pre-tax, pass-through), not a GAAP EBITDA bridge "
        "-- interest, tax, and D&A lines are included for format familiarity but are $0 or "
        "N/A for this entity type, consistent with mna_valuation_engine.py and "
        "lbo_model/build_lbo_model.py elsewhere in this repo.",
        "",
        "The monthly add-back allocation on the Add-Back Detail tab is illustrative, not "
        "independently verified -- it exists to demonstrate the level of granularity a real "
        "QoE traces to (bank statement line items, dated receipts), not to certify these "
        "specific monthly splits. Every cell there is explicitly marked unverified.",
        "",
        "Quarterly revenue and gross margin figures, and all twelve months of AR/Inventory/AP "
        "on the NWC Analysis tab, are illustrative and were not pulled from a live POS or "
        "accounting export for this portfolio version -- see each tab's cell comments for "
        "what a real engagement would confirm them against (Faire payment tracker, AR aging, "
        "physical inventory count, vendor AP aging).",
        "",
        "A real QoE engagement would also include: customer/vendor concentration analysis, "
        "related-party transaction review, a cash-basis-to-accrual-basis reconciliation, "
        "lease and other off-balance-sheet commitment review, and (for a deal this size) a "
        "review of owner-vs-business expense commingling beyond the two add-backs already "
        "identified. None of those are in scope here -- this workbook demonstrates the "
        "EBITDA-bridge and NWC-peg mechanics, not a complete diligence engagement.",
        "",
        "This is a portfolio project. The deal, business, and dollar figures are illustrative "
        "and tie to the same anonymized scenario as mna_valuation_engine.py and "
        "lbo_model/LBO_Model.xlsx elsewhere in this repo -- no real business's financial data "
        "appears here.",
    ]
    for i, line in enumerate(notes, start=3):
        cell = ws[f"A{i}"]
        cell.value = line
        cell.font = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def build():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_overview(wb)
    build_ebitda_bridge(wb)
    build_addback_detail(wb)
    build_revenue_quality(wb)
    build_nwc_analysis(wb)
    build_notes(wb)
    wb.save(OUT_FILE)
    print(f"[SUCCESS] QoE workbook written -> {OUT_FILE}")
    return OUT_FILE


if __name__ == "__main__":
    build()
