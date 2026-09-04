# AQDQ tracker — field dictionary & extraction rules

Authoritative spec for turning an Accumulator/Decumulator termsheet or trade recap into the
15-column tracker row. Validated 2026-06-03 against a real Goldman Sachs / UOB Kay Hian
leveraged accumulator on CRWD (trade ref SDB1453800981).

## The 15 output columns (tracker columns A–O)

| # | Col | Field | Format | Source / rule |
|---|---|---|---|---|
| 1 | A | AQ/DQ | `AQ` or `DQ` | product classification (below) |
| 2 | B | Trade Date | M/D/YYYY | "Trade Date" |
| 3 | C | Initial Obs Date | M/D/YYYY | **Effective Date** — when daily accrual starts (= Period 1 start). Do NOT use the first bi-weekly Valuation Date. |
| 4 | D | GTD Period End | M/D/YYYY | end of the guaranteed period = the Valuation/End Date of the **last guaranteed period** ("Guaranteed Period: 1st to 2nd period" + "Minimum Share Accrual Period ... to and including the 2nd Valuation Date" ⇒ the 2nd Valuation Date) |
| 5 | E | GTD Days | integer | scheduled-trading-day count of the guaranteed period (below) |
| 6 | F | Final Obs Date | M/D/YYYY | last Valuation Date / "Scheduled Valuation Date for final Calculation Period" |
| 7 | G | Total Expiries | integer | total scheduled trading days of the whole schedule (below) |
| 8 | H | Notional | number, no commas | document "Notional Amount" if stated; else Total Expiries × Shares/Day × Trade Price |
| 9 | I | Currency | e.g. USD | the termsheet's settlement currency |
| 10 | J | Leverage | integer | below-forward quantity ÷ base quantity (below) |
| 11 | K | Shares per Day | integer | the **base** daily quantity (below) |
| 12 | L | Stock | Yahoo ticker | `CRWD.OQ` / `CRWD UW` → `CRWD`; non-US gets the Yahoo suffix (0700.HK, 7203.T, 005930.KS) |
| 13 | M | Trade Price | 4 dp as shown | "Spot" / "Reference Spot Price" |
| 14 | N | KO Level | **decimal** (1.07 = 107%) | "Knock-out Price ... % of Spot" |
| 15 | O | Strike Level | **decimal** (0.747 = 74.70%) | "Forward Price" / "Strike" / "Exercise Price" ... % of Spot |

Tracker columns P (KO Price = M×N) and Q (Strike Price = M×O) and everything after (Database
Last Close, statuses, P&L analytics) are computed by the sheet — **not output**.

KO/Strike are emitted as decimals, not "107.00%" text, because the user needs values that always
participate in Excel calculations; percent-strings usually auto-convert on paste but can land as
text in some import paths. The tracker's % cell format displays 1.07 as 107.00%.

## Shares/Day vs Leverage (critical)
The termsheet states something like: "Daily Number of Shares: **1** Share if the Closing Price is
equal to or greater than the Forward Price, or **2** Shares if it is less."
- **Shares per Day = the at-or-above quantity (the base)** → 1 here.
- **Leverage = below-forward quantity ÷ base** → 2 here. Non-leveraged products ⇒ Leverage 1.

Getting this backwards doubles the notional. Always verify with the reconciliation:
`Notional ≈ Total Expiries × Shares/Day × Trade Price`
(CRWD: 250 × 1 × 622.8267 = 155,706.68 ≈ the stated USD 155,706.67 ✓ — proving base = 1.)
`gen_aqdq_csv.py` re-runs this check and warns on >0.5% mismatch.

## Total Expiries (priority order)
1. **Explicit field** for the whole schedule: "Total Expiries", "Currently expected total number
   of Scheduled Trading Days" (CRWD: 250), "Number of Observations". Use as-is.
2. **Sum the SCHEDULE table's per-period trading-day counts** N(period_j) — exact and already
   holiday-aware (CRWD: the 26 periods sum to 250).
3. Last resort only: count Mon–Fri between Initial Obs and Final Obs inclusive and **flag it as
   approximate** — a plain weekday count ignores exchange holidays and overstates a 1-year trade
   by ~10 days.
Never confuse Guarantee Days with Total Expiries, and never use min/max ranges.

## Guarantee Days
Sum N(period_j) over the guaranteed periods from the SCHEDULE — e.g. guaranteed = periods 1–2 ⇒
N(1)+N(2) = 9+10 = **19**. A naive weekday count gives 20 here because it misses the Memorial-Day
holiday — the schedule's N column is the source of truth. If there is no per-period table, count
trading days from the Effective Date to the guaranteed-period end and say it's approximate.

**No guarantee at all (order table "Guarantee Period: 0W"):** some termsheets — e.g. the Barclays
leveraged decumulators — carry no Minimum Share Accrual / guaranteed-period clause; the knock-out
is live from the Trade Date. Then **GTD Days = 0** and **GTD Period End is left BLANK** (user's
convention, confirmed 2026-08-05) — do not substitute the trade or effective date. Pass
`"gtd_period_end": ""` in the JSON; the generator writes an empty cell.

## AQ / DQ classification
- **AQ (Accumulator)** — client *buys/accumulates* shares daily, "Buy Below Market": strike /
  forward **below** spot (<100%), knock-out **above** spot (>100%).
- **DQ (Decumulator)** — client *sells* shares daily, "Sell Above Market": strike **above** spot
  (>100%), knock-out **below** spot (<100%).
Classify from the product name/mechanics; then sanity-check the KO/Strike directions (the script
warns if they contradict). Keep a contradicting value only if the document is explicit.

## Label synonyms (map by meaning, not label)
- Trade Price ← "Spot", "Reference Spot Price".
- Strike ← "Forward Price", "Strike Price", "Exercise Price", "Settlement Price";
  "Leverage Trigger Price" usually equals the Forward and is not separately tracked.
- KO ← "Knock-out Price", "Target Redemption", "KO Barrier".
- If only price levels are printed (no %), derive the level = price ÷ spot.

## Worked example (CRWD leveraged accumulator, GS/UOBKH, validated)
"Accumulator Leveraged Knock-out Forward with Minimum Share Accrual Period" on CRWD —
Trade 19 May 2026, Effective 20 May 2026, Spot USD 622.8267, Forward 465.2515 (74.70%),
Knock-out 666.4246 (107.00%), "1 Share if ≥ Forward / 2 if <", Guaranteed = periods 1–2,
26 bi-weekly valuation dates ending 18 May 2027, "total Scheduled Trading Days: 250",
Notional 155,706.67:

`AQ | 5/19/2026 | 5/20/2026 | 6/16/2026 | 19 | 5/18/2027 | 250 | 155706.67 | USD | 2 | 1 | CRWD | 622.8267 | 1.07 | 0.747`

(GTD End 6/16/2026 = 2nd Valuation Date; GTD Days 19 = N(1)+N(2) = 9+10.)
