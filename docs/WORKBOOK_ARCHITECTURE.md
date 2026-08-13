# Workbook Architecture

## Design principle

The workbook is a layered decision system. Inputs are entered once, calculation engines remain auditable, and decision outputs never hide hard-stop risks.

## Sheet map

| Layer | Sheets |
|---|---|
| Onboarding | Start Here, Quick Analysis |
| Core inputs | Property Input, Financing, Buy Box |
| Strategy engines | Rental, Flip, Wholesale, BRRRR, Short-Term Rental |
| Intelligence | Strategy Battle, Deal DNA, Risk Radar, Stress Lab, Deal Rescue |
| Execution | Offer Builder, Due Diligence, Deal Tracker |
| Outputs | Executive Dashboard, Investment Memo |
| Governance | Settings, Glossary, Checks, Changelog |

## Formula flow

Property Input -> Financing -> Strategy Engines -> Intelligence -> Dashboard/Investment Memo

Buy Box and Settings feed every strategy engine and decision rule. Checks inspect inputs, schedules, and outputs.

## Scoring framework

| Dimension | Weight |
|---|---:|
| Profitability and return | 30 |
| Cash flow and debt coverage | 25 |
| Equity and purchase discount | 20 |
| Downside protection | 15 |
| Data confidence | 10 |

Score bands:

- 80-100: Strong candidate
- 65-79: Promising - verify assumptions
- 50-64: High caution
- 0-49: Does not meet criteria

Hard stops override the display verdict when material risks exist, including negative cash flow, missing ARV, insufficient buyer spread, DSCR below a required floor, or capital required above the user's limit.

## Core formulas

Rental:
- Effective income = scheduled rent + other income - vacancy
- NOI = effective income - operating expenses
- Cap rate = NOI / purchase price
- DSCR = NOI / annual debt service
- Cash-on-cash = annual pre-tax cash flow / total cash invested

Flip:
- Net profit = sale price - purchase - rehab - acquisition - holding - financing - selling costs
- Break-even sale price = total project cost / (1 - selling cost rate)
- MAO is solved from the user's minimum profit and cost assumptions

Wholesale:
- Buyer price = contract price + assignment fee
- Buyer projected profit = ARV - buyer price - repairs - buyer closing/holding/selling costs
- MAO uses an editable buyer profit requirement; no fixed 70 percent rule is imposed

## Audit rules

- No magic numbers in formulas
- Cross-sheet formulas quote sheet names
- Ratios guard missing and zero denominators
- Inputs, formulas, links, and warnings follow consistent color conventions
- A visible Checks sheet lists Actual, Expected, Difference, Tolerance, Status, and Notes
- Sensitivity outputs recalculate underlying economics
