# M&A Valuation Engine, LBO Model, QoE Workbook & IC Memo

Anonymized portfolio version of six connected pieces I directed the AI-assisted development of for a real small business exit-planning engagement -- covering what the business is worth, whether that number holds up against real market data, how a buyer could actually finance it, whether the earnings backing it are real, and how it all gets synthesized and presented to two different audiences. Business name, real dollar figures, and file paths below are illustrative, not real -- but every number across all six pieces ties to the same single scenario, so nothing here is siloed by design.

## The problem

Small-business valuation for an eventual sale usually starts from "recast earnings" -- net income plus discretionary add-backs the owner wouldn't need going forward (personal draws, one-time expenses) -- times an industry multiple. That's a simple formula. The actual engineering problem is trust: every number that formula touches has to be traceable to a real, verified source, and every add-back has to be honestly labeled as verified or not, because this kind of document eventually sits in front of a buyer, lender, or investment committee.

An earlier version of this valuation had net income **hardcoded as a literal number** directly in the script. It was wrong -- stale from an earlier, less-reconciled pass at the books -- and nothing would have caught it silently drifting further from the real, audited figure over time.

## What `mna_valuation_engine.py` does

- Pulls Net Income **live** from a specific verified cell in the business's own reconciled tax package, instead of a hardcoded number. If the source file is missing, unreadable, or that cell doesn't contain a real number, it fails loudly with a specific error message -- it never silently falls back to a guess.
- Applies discretionary add-backs to compute recast SDE, with each add-back individually flagged `verified: True/False` in the output rather than presented as uniformly trustworthy.
- Two add-back line items that had been sitting in an earlier version were removed after review because they contradicted facts already verified elsewhere in the same engagement (one double-counted a tax treatment already handled upstream; the other eliminated interest on a related-entity loan that turned out to carry 0% interest in the first place, so there was nothing to eliminate). The removal reasoning is preserved in the script's own comments rather than silently deleted, because "why isn't this here anymore" matters as much as "why is this here."
- Produces a conservative/base/aggressive valuation range off an industry-multiple assumption -- checked against real market data by `market_comps.py` below, rather than left as an unexamined judgment call.
- The output JSON carries its own `dealRoomReadiness` field -- "NOT READY -- needs owner sign-off" until a human explicitly clears it -- so nothing downstream can accidentally treat draft figures as final.

## What `mna_deal_room_generator.py` does

Reads the valuation engine's JSON output and assembles an HTML executive teaser from it -- rather than an earlier version that had its own **independent, second hardcoded copy** of the same SDE and valuation figures baked directly into the HTML template. That second copy was already stale the moment the valuation engine's numbers changed, since the two files never actually referenced each other. Now there is exactly one source of truth for these numbers, and the generator refuses to run at all if the valuation engine hasn't produced fresh output yet.

Every claim in the generated teaser that isn't a verified number from the valuation JSON -- margin percentage, year-over-year growth -- is explicitly labeled `[UNVERIFIED -- confirm against X.py before use]` directly in the generated document, and a draft banner renders automatically across the top of the whole page until the valuation is marked ready. The goal is a document that's honest about its own confidence level, not one that looks more finished than the underlying numbers actually are.

## What `lbo_model/build_lbo_model.py` does

Takes the valuation engine's output (recast SDE $169,975.60, base-case 3.00x multiple, $509,926.80 enterprise value) and builds a fully formula-driven Excel LBO model answering the buyer's side of the same deal: what capital structure could actually finance it, and what would a sponsor's return look like.

- **Deliberately doesn't use a textbook institutional LBO capital structure.** A deal this size (sub-$200K SDE) doesn't clear for a syndicated senior term loan + high-yield bond stack -- no institutional lender underwrites acquisition debt at that scale. The model uses SBA 7(a) term debt + a seller note + sponsor equity, the structure independent sponsors and search funds actually use at this deal size, with the sponsor-equity assumption set at exactly the SBA's real ~10% minimum injection for a change-of-ownership 7(a) loan.
- Six sheets: Assumptions, Sources & Uses, Operating Model, Debt Schedule, Returns Analysis, and Notes & Limitations -- 152 live formulas, zero hardcoded results, built and verified with `openpyxl` + LibreOffice per this repo's spreadsheet conventions.
- **A real bug caught before shipping**: an early version of the debt schedule had an off-by-one row reference (each year's beginning balance pulled from the prior year's *principal payment* row instead of its *ending balance* row). LibreOffice recalculated that version with **zero formula errors** -- a clean recalc only proves formulas evaluate, not that they're right. The bug only surfaced by manually tracing the amortization schedule and noticing balances going negative and MOIC coming out at an implausible 13x+. Fixed, and now covered by a regression test (`test_sba_amortization_references_prior_year_ending_balance_not_principal`) that checks the formula text itself, not just that the file recalculates cleanly.
- Base-case returns (5-year hold, 3.00x exit multiple) come out to an 8.53x MOIC / 53.5% IRR -- high for a standard institutional LBO, but that's the mechanical result of a ~10% equity check amplifying the same dollar of EBITDA growth and debt paydown, not an aggressive growth assumption doing the work. The Notes tab spells out why, and that the same leverage cuts both ways if SDE declines instead of growing.

## What `qoe_workbook/build_qoe_workbook.py` does

The diligence counterpart to the valuation engine: a recast-SDE valuation answers "what is this worth if the earnings are real," and a Quality of Earnings workbook is the step that tests whether they're real, before a buyer (or a seller's advisor prepping for buyer diligence) relies on the number.

- **Adjusted EBITDA Bridge** that reconciles, formula-for-formula, back to the exact same verified net income ($148,220.60) and add-backs ($21,755.00) the valuation engine uses -- with a tie-out check cell that must equal zero.
- **Add-Back Detail** tab that breaks the two annual add-back totals into an illustrative monthly trace, the level of granularity a real QoE would build against actual bank/ledger line items -- every monthly cell is explicitly marked unverified rather than implicitly borrowing credibility from the verified annual figure it rolls up to.
- **Revenue & Margin Quality** tab showing the same Year-0 revenue used in the LBO model broken into quarters, flagging that 35% of the year's revenue lands in one holiday quarter at the year's highest margin -- exactly the kind of concentration a buyer should pressure-test rather than assume repeats.
- **NWC Analysis** tab with a trailing-12-month monthly net working capital build (AR + Inventory − AP), pegged to the 12-month average rather than the at-close balance specifically so a year-end closing doesn't reward either party for seasonal inventory timing -- with the purchase-price adjustment mechanics spelled out on the sheet.
- Every hardcoded input carries a cell comment naming what it would need to be confirmed against in a real engagement (Faire wholesale AR aging, physical inventory count, vendor AP aging, bank statements) -- the same verified/unverified discipline as the rest of this repo, applied at a finer grain.

## What `market_comps.py` does

Closes a gap the valuation engine had always flagged but never actually checked: its conservative/base/aggressive multiples were labeled "a market judgment call, not derived from the business's own data" from day one, but nothing tested that judgment call against anything real.

- Cites four **real, dated, sourced** market comps -- not illustrative figures, since these are genuine third-party statistics, not the business's own confidential financials: BizBuySell's five-year trailing industry multiple table (retail sector, clothing & accessory stores, jewelry stores) and the IBBA/M&A Source Q1 2026 Market Pulse Survey's deal-size-bracket multiple. Every comp carries a live URL.
- **The finding**: the valuation engine's base-case multiple (3.00x) sits above every cited comp (2.13x-2.80x); the conservative multiple (2.50x) is the one that actually lands inside real market data. Rather than quietly changing the existing, already-tested valuation assumptions, this script surfaces that finding as its own output -- the same "flag it, don't silently fix it" discipline this repo already applies to unverified add-backs and the asset liquidation floor.
- Reads the valuation engine's live JSON output rather than a hardcoded copy of its multiples, so if those assumptions ever change, this check automatically compares against the new numbers.

## What `ic_memo_generator.py` does

`mna_deal_room_generator.py` produces the external, buyer-facing document. An Investment Committee memo is the different, internal document a real deal team produces -- the one that goes to the people deciding whether to proceed, and it has to say things a teaser never would.

- Reads all four upstream outputs -- the valuation JSON, the market comps JSON, the recalculated `LBO_Model.xlsx`, and the recalculated `QoE_Workbook.xlsx` -- and synthesizes them into one polished Word document with a real recommendation, not four documents a reader has to reconcile themselves.
- Sections: Transaction Summary, Investment Thesis, Recast Earnings & QoE, Valuation & Market Support (with the comps table and the above-market-multiple finding front and center), Capital Structure & Financing, Returns Analysis with sensitivity grid, Key Risks & Mitigants (pulled from both workbooks' own Notes tabs), and Recommendation & Conditions to Proceed.
- Writes nothing that isn't traceable to one of the four source outputs, and fails loudly -- naming the exact cell and which script to re-run -- if a prerequisite is missing or was never recalculated, rather than silently rendering a blank or stale figure into a document someone might present as final.
- **A real bug caught before shipping**: an early version pulled the LBO model's MOIC sensitivity grid (rows 17-19) for the memo's IRR sensitivity table instead of the actual IRR grid (rows 24-26) -- both are keyed by the same exit-multiple labels, so it was an easy mix-up. The output still generated cleanly; it just rendered MOIC multiples formatted as percentages (an 8.53x came out as "853%" IRR). Caught by reading the rendered PDF, not by trusting that "it ran without an error" meant it was right -- the same lesson as the LBO model's own off-by-one bug, applied one script later.

## My role

I specified what counted as a verified vs. unverified figure, traced the stale hardcoded net-income bug back to its source before trusting anything built on top of it, and made the call to remove the two contradictory add-back line items after checking them against facts already established elsewhere in the same engagement. I also designed the "nothing downstream can silently treat a draft as final" pattern -- the readiness flag, the unverified-claim labeling, the draft banner -- rather than shipping a valuation document that looks more authoritative than it actually is.

For the LBO model, I made the underwriting judgment call on capital structure (SBA acquisition financing over a textbook institutional stack, sized to match the actual deal size) rather than accepting a generic template, and caught the off-by-one amortization bug by manually verifying computed values instead of trusting a clean recalc -- then wrote a regression test that checks the formula text so the same bug class can't silently ship again. For the QoE workbook, I chose which sections a diligence workbook at this deal size actually needs (EBITDA bridge, add-back trace, revenue/margin trend, NWC peg) versus a full institutional engagement's scope, and made sure every new number introduced at the monthly/quarterly grain still reconciled back to the verified annual figures already established elsewhere in the repo.

For the comps check, I chose real, citable sources over convenient ones, picked the specific BizBuySell subcategories that actually match this business's product mix (jewelry is its largest unit-volume category per its own sales data) rather than only the broad retail-sector average, and made the deliberate call to surface the finding that the base case runs ahead of market data rather than quietly retuning the multiple to make the comps check pass. For the IC memo, I decided what an internal decision document needs that an external teaser can't say, caught the MOIC/IRR sensitivity mix-up by reading the actual rendered output instead of trusting a clean script run, and found -- then fixed -- a second, unrelated bug: running the test suite was silently corrupting the checked-in `LBO_Model.xlsx` and `QoE_Workbook.xlsx` by rebuilding them in place without recalculating, which I traced to the test fixtures and fixed by pointing them at a throwaway temp path instead.

## Run it

```bash
pip install -r requirements.txt
python build_sample_data.py   # writes an illustrative sample tax workbook
python demo.py                 # runs the valuation engine, then the deal room generator
python lbo_model/build_lbo_model.py       # builds LBO_Model.xlsx
python qoe_workbook/build_qoe_workbook.py # builds QoE_Workbook.xlsx
python market_comps.py                    # benchmarks the valuation multiple against real comps
python ic_memo_generator.py               # assembles the IC memo from all four outputs above
```

No real business data is required or included.

## Tests

```bash
pytest tests/
```

28 tests. The ones that matter most: `test_recast_sde_and_valuation_math_matches_verified_net_income` pins the exact numbers down (net income $148,220.60 + $21,755.00 in add-backs = recast SDE $169,975.60, base valuation exactly 3.00x that); `test_get_verified_net_income_fails_loudly_*` confirms the engine refuses to guess when the tax package is missing or invalid; `test_sba_amortization_references_prior_year_ending_balance_not_principal` is the regression test for the LBO model's off-by-one debt-schedule bug; `test_reads_live_multiples_not_a_hardcoded_copy` proves the comps check compares against the valuation engine's *current* assumptions, not a frozen copy; and `test_end_to_end_memo_contains_cross_checked_figures` confirms the generated IC memo actually contains the real, cross-checked figures from all four upstream outputs (including the corrected IRR sensitivity numbers). The recalculation-dependent workbook tests skip automatically if LibreOffice isn't installed, and none of the test fixtures touch the committed, already-recalculated `LBO_Model.xlsx` / `QoE_Workbook.xlsx` in place -- they build into a throwaway temp path instead, precisely because an earlier version of those fixtures didn't and silently broke both files every time the suite ran.

## Stack

Python, `openpyxl` for reading the live tax workbook and building/reading the LBO/QoE Excel models (formula-driven, LibreOffice-recalculated, zero hardcoded results), `python-docx` for the generated IC memo, JSON as the interchange format connecting all six pieces, plain HTML/CSS for the generated teaser.
