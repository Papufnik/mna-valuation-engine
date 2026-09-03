# M&A Valuation Engine & Deal Room Generator

Anonymized portfolio version of two connected scripts I directed the AI-assisted development of for a real small business exit-planning engagement. The first turns a verified tax filing into a recast Seller's Discretionary Earnings (SDE) valuation; the second reads that valuation and assembles a buyer-facing deal room package from it. Business name, real dollar figures, and file paths below are illustrative, not real.

## The problem

Small-business valuation for an eventual sale usually starts from "recast earnings" -- net income plus discretionary add-backs the owner wouldn't need going forward (personal draws, one-time expenses) -- times an industry multiple. That's a simple formula. The actual engineering problem is trust: every number that formula touches has to be traceable to a real, verified source, and every add-back has to be honestly labeled as verified or not, because this kind of document eventually sits in front of a buyer or broker.

An earlier version of this valuation had net income **hardcoded as a literal number** directly in the script. It was wrong -- stale from an earlier, less-reconciled pass at the books -- and nothing would have caught it silently drifting further from the real, audited figure over time.

## What `mna_valuation_engine.py` does

- Pulls Net Income **live** from a specific verified cell in the business's own reconciled tax package, instead of a hardcoded number. If the source file is missing, unreadable, or that cell doesn't contain a real number, it fails loudly with a specific error message -- it never silently falls back to a guess.
- Applies discretionary add-backs to compute recast SDE, with each add-back individually flagged `verified: True/False` in the output rather than presented as uniformly trustworthy.
- Two add-back line items that had been sitting in an earlier version were removed after review because they contradicted facts already verified elsewhere in the same engagement (one double-counted a tax treatment already handled upstream; the other eliminated interest on a related-entity loan that turned out to carry 0% interest in the first place, so there was nothing to eliminate). The removal reasoning is preserved in the script's own comments rather than silently deleted, because "why isn't this here anymore" matters as much as "why is this here."
- Produces a conservative/base/aggressive valuation range off an industry-multiple assumption that's explicitly labeled as a market judgment call, not something derived from the business's own data.
- The output JSON carries its own `dealRoomReadiness` field -- "NOT READY -- needs owner sign-off" until a human explicitly clears it -- so nothing downstream can accidentally treat draft figures as final.

## What `mna_deal_room_generator.py` does

Reads the valuation engine's JSON output and assembles an HTML executive teaser from it -- rather than an earlier version that had its own **independent, second hardcoded copy** of the same SDE and valuation figures baked directly into the HTML template. That second copy was already stale the moment the valuation engine's numbers changed, since the two files never actually referenced each other. Now there is exactly one source of truth for these numbers, and the generator refuses to run at all if the valuation engine hasn't produced fresh output yet.

Every claim in the generated teaser that isn't a verified number from the valuation JSON -- margin percentage, year-over-year growth -- is explicitly labeled `[UNVERIFIED -- confirm against X.py before use]` directly in the generated document, and a draft banner renders automatically across the top of the whole page until the valuation is marked ready. The goal is a document that's honest about its own confidence level, not one that looks more finished than the underlying numbers actually are.

## My role

I specified what counted as a verified vs. unverified figure, traced the stale hardcoded net-income bug back to its source before trusting anything built on top of it, and made the call to remove the two contradictory add-back line items after checking them against facts already established elsewhere in the same engagement. I also designed the "nothing downstream can silently treat a draft as final" pattern -- the readiness flag, the unverified-claim labeling, the draft banner -- rather than shipping a valuation document that looks more authoritative than it actually is.

## Run it

```bash
pip install -r requirements.txt
python build_sample_data.py   # writes an illustrative sample tax workbook
python demo.py                 # runs the valuation engine, then the deal room generator
```

No real business data is required or included.

## Tests

```bash
pytest tests/
```

8 tests. The two that matter most: `test_recast_sde_and_valuation_math_matches_verified_net_income` pins the exact numbers down (net income $148,220.60 + $21,755.00 in add-backs = recast SDE $169,975.60, base valuation exactly 3.00x that), and `test_get_verified_net_income_fails_loudly_*` confirms the engine refuses to guess -- raising instead of silently defaulting -- when the tax package is missing or its net-income cell doesn't hold a real number, which is exactly the failure mode a hardcoded fallback would have hidden.

## Stack

Python, `openpyxl` for reading the live tax workbook, JSON as the interchange format between the two scripts, plain HTML/CSS for the generated teaser.
