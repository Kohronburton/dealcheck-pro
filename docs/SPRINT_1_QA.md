# Sprint 1 QA Evidence

## Build

- Workbook: DealCheck Pro Sprint 1 MVP
- Version: 0.1.0
- Sheets: Start Here, Quick Analysis, Property Input, Financing, Rental, Flip, Wholesale, Checks
- Sample property: 125 Sample Ave, Miami, FL

## Automated checks

- Formula-error scan: PASS - no #REF!, #DIV/0!, #VALUE!, #NAME?, or unhandled #N/A matches
- Financing tie-out: PASS
- Rental NOI tie-out: PASS
- Purchase price required: PASS
- ARV required: PASS
- Flip project cost positive: PASS
- Wholesale buyer spread nonnegative: PASS

## Visual review

All eight sheets were rendered and reviewed. Titles, inputs, formulas, warnings, tables, and primary outputs are visible without clipping. Excess blank formatted columns were removed after the first render pass.

## Sample result

The included sample deliberately demonstrates a weak deal:

- Rental cash flow is approximately $8 per month, below the $300 Buy Box target
- Flip profit is negative
- Proposed purchase price is above the buyer-driven wholesale MAO
- Model integrity checks remain OK

This distinction is intentional: model integrity can pass while the investment itself fails.

## Remaining Sprint 1 tests

- All-cash rental
- Zero-interest amortization
- Zero rent
- 100 percent vacancy
- Missing ARV
- Rehab overrun
- Assignment fee above buyer spread
- Excel desktop smoke test
- Google Sheets import smoke test
