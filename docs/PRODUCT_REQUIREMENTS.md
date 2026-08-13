# Product Requirements - DealCheck Pro 10X

## Vision

DealCheck Pro is a spreadsheet-native underwriting and decision system for residential real-estate investors. It turns uncertain property inputs into an explainable recommendation, safe offer range, downside analysis, and next action.

## Primary user

New and intermediate residential investors analyzing one- to four-unit properties for rental, flip, BRRRR, wholesale, or short-term rental strategies.

## Jobs to be done

1. Decide whether a property deserves more time or an offer.
2. Compare strategies without entering the property twice.
3. Detect missing expenses and weak assumptions.
4. Calculate the price or terms required to meet a personal buy box.
5. Communicate the opportunity to a partner, lender, or client.
6. Track several opportunities and prioritize limited capital.

## Product promise

Enter a property once. Compare five strategies. Stress-test the assumptions. Identify hidden risks. Calculate the safest offer. Produce an investor-ready recommendation in under five minutes.

## Core experience

### Quick Mode

Requires only purchase price, ARV, rehab, rent, taxes, insurance, financing rate, term, and down payment. Returns a preliminary Rental, Flip, and Wholesale comparison.

### Pro Mode

Adds detailed expenses, BRRRR, short-term rental, capital stack, comps, confidence grading, stress tests, offer optimization, due diligence, partnership economics, and an investment memo.

## Must-have capabilities

- Centralized property and financing inputs
- Personalized Buy Box
- Rental, Flip, Wholesale, BRRRR, and STR engines
- Deal DNA dimension scores and explainable deductions
- Strategy Battle comparison
- Offer Optimizer with opening, target, maximum, and walk-away prices
- Deal Rescue recommendations
- Assumption Confidence grading
- Red-Flag Radar
- Conservative, expected, and optimistic scenarios
- One-page investment memo
- Visible model checks and formula audit status

## UX rules

- Blue text: editable inputs
- Black text: formulas
- Green text: cross-sheet references
- Red text: external links
- Yellow fill: missing or attention-required inputs
- Green/yellow/red verdicts: favorable/review/fail
- No macros, paid APIs, live scraping, or credentials in Version 1
- Core formulas must work in current Excel and Google Sheets
- Formula sheets may be hidden but never unauditable

## Non-functional requirements

- No unhandled formula errors
- All ratios guard zero and missing denominators
- All business assumptions live in visible cells
- All strategy outputs reconcile to component schedules
- Printable sheets fit their intended page size
- A beginner can complete Quick Mode without reading a manual
- A worked sample is included
- Version and changelog are visible

## Launch scope

Version 1 launches Rental, Flip, Wholesale, Dashboard, Buy Box, Deal Tracker, and Investor Memo. BRRRR, STR, Deal Rescue, and deeper stress testing follow as shippable increments.

## Commercial model

Initial SKUs: Wholesale MAO, Rental Lite, Flip Analyzer, DealCheck Pro, and Elite Bundle. Etsy is the discovery channel; Payhip or Gumroad follows for bundles and customer ownership.

## Disclaimer

The workbook provides educational estimates and analysis. It is not legal, tax, lending, appraisal, financial, or investment advice and does not guarantee results. Users must independently verify all property, market, financing, and regulatory data.
