# Sprint 1 - Fast Decision MVP

## Sprint goal

A user can enter one residential property in Quick Mode and receive reliable Rental, Flip, and Wholesale results with clear inputs, cash required, expected outcome, and calculation checks.

## In scope

- Start Here
- Quick Analysis
- Property Input
- Financing
- Buy Box baseline
- Rental engine
- Flip engine
- Wholesale engine
- Checks
- Worked sample property

## Out of scope

BRRRR, short-term rental, capital stack, comps, partner waterfall, automated scraping, VBA, and paid APIs.

## User stories

1. As a beginner, I can distinguish inputs from formulas and complete a preliminary analysis in under five minutes.
2. As a rental investor, I can see NOI, cap rate, DSCR, monthly cash flow, cash-on-cash return, and cash required.
3. As a flipper, I can see project cost, net profit, ROI, break-even sale price, and maximum offer.
4. As a wholesaler, I can calculate an editable buyer-driven MAO and confirm buyer spread.
5. As a reviewer, I can see whether required inputs and formulas pass visible checks.

## Acceptance test fixtures

- Financed positive-cash-flow rental
- All-cash rental
- Negative-cash-flow rental
- Profitable flip
- ARV below total project cost
- Rehab 20 percent over expected
- Wholesale assignment fee above buyer spread
- Zero rent
- Zero debt service
- Missing ARV
- Zero vacancy
- 100 percent vacancy

## Sprint exit criteria

- Key outputs reconcile to supporting schedules
- No #REF!, #DIV/0!, #VALUE!, #NAME?, or unhandled #N/A errors
- Every user-facing sheet is visually reviewed
- Sample deal results are internally consistent
- Exported .xlsx opens successfully
- Known Google Sheets limitations are documented
