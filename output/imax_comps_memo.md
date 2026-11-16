# Valuation memo: IMAX CORP (IMAX)

## Valuation summary

- **Business:** IMAX CORP generated revenue of $410 million in the fiscal year ended 2025-12-31, up 16.5% on the prior year, with EBITDA of $147 million (35.8% margin) and EBIT of $84 million.
- **Against peers:** Its EBITDA margin is 17.9 points above the peer median of 17.9%, and its revenue growth of 16.5% compares with a peer median of 4.0%.
- **Leverage:** Net debt is $127 million, or 0.9x EBITDA.
- **Market value:** At $48.29 per share the market values IMAX at an enterprise value of $2,883 million: 7.0x revenue and 19.7x EBITDA, a 13% premium to the peer median of 17.3x.
- **Value from comps:** Applying the peer median EV/EBITDA of 17.3x to EBITDA of $147 million implies an enterprise value of $2,544 million, or $42.15 per share.
- **Value from DCF:** The DCF, at a 7.4% discount rate and 2.5% terminal growth, implies an enterprise value of $1,626 million and $25.53 per share, 47% below the current price. Across the sensitivity grid, value per share runs from $17.59 to $44.30.
- **Reading:** Both methods put value more than 10% below the current price. This is the model's output under its stated assumptions, not a recommendation.


**Audit result: 7 warnings across 18 checks.**

## Model audit

- **Judgment assumptions [WARNING]:** 2 of 3 judgment assumptions (discount rate, terminal growth, working capital) are still tool defaults. Set them in the Override column on the DCF sheet; until then the DCF value is not defensible.
- **Discount rate source [OK]:** The discount rate of 7.4% is built on the WACC sheet: risk-free 5.1%, beta 0.50, equity risk premium 5.0%, equity weight 91%.
- **Beta [WARNING]:** Beta is 0.50. Below 0.6 is unusually low for an operating company and pulls the discount rate down. A beta measured on one stock is noisy: compare it with the peers' betas before relying on it.
- **Cash flow reconciliation [OK]:** The model projects year-1 free cash flow of $67 million. Last year the company actually generated $85 million of cash from operations after all investing.
- **Terminal value weight [WARNING]:** Terminal value is 80% of enterprise value. Above 75%, the valuation rests mostly on what happens after year 5, which is the least knowable part.
- **Discount rate vs growth [OK]:** The discount rate exceeds terminal growth by 4.9 points.
- **Implied exit multiple [OK]:** The terminal value equals 8.5x year-5 EBITDA, against a peer median trading multiple of 17.3x.
- **Margin assumption [WARNING]:** Projected EBIT margin is 20.5% against a 3-year average of 15.5%. The model holds the latest year flat, and the latest year is well above the average: you are assuming a peak margin lasts forever.
- **Growth assumption [OK]:** Year-1 revenue growth is 12.6%, fading to 4.5% by year 5.
- **Reinvestment [OK]:** D&A is 15.2% of revenue and reinvestment is 16.5%.
- **DCF vs market [OK]:** The DCF value is 47% below the share price.
- **Profitability [OK]:** EBIT is positive, so margin-based projections are mechanically valid.
- **Share prices [OK]:** All share prices are entered. Check the price dates are close together.
- **IMAX data quality [WARNING]:** Approximations or gaps in the target's data: Debt = convertible debt + credit facility (no standard debt tag). Verify against the 10-K.
- **Peer data quality [WARNING]:** Peers with a gap or approximation (see Checks sheet): DLB.
- **Fiscal year alignment [WARNING]:** Peers whose fiscal year ends more than 3 months from the target's: DLB (2025-09-26), DIS (2025-09-27). Their figures cover a different 12 months.
- **Share count [INFO]:** Market cap uses the share count on the cover of the latest filing (2026-09-30). It excludes dilution from options and convertible debt.
- **Formula integrity [OK]:** The centre of the sensitivity grid ties to the DCF value per share.
