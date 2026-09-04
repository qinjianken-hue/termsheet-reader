# BEN column dictionary & extraction rules

## What a BEN is
A **Bonus Enhanced Note** (dealer product name on the place-order table: `Regular BEN`) is a
worst-of equity-linked note whose redemption at maturity is:

- **all** Reference Price (Final) ≥ their Put Strike Price →
  `Denomination × Max( Worst Underlying Performance (Final), 100% + Coupon Flat )`
- **at least one** below its Put Strike Price → **physical delivery** of the worst performer
  (Denomination ÷ Put Strike Price of the worst underlying, rounded down, cash for the stub).

Two consequences worth stating in every report:
1. **Coupon Flat is a flat maturity payout, not p.a.** A 6M note quoting 35.87% pays 135.87% at
   maturity if the barrier holds. Never annualise it.
2. **The `Max(...)` means the upside is uncapped** — above `100% + Coupon Flat` the note tracks the
   worst performer at the Upside Participation rate (100% on every note seen so far). This is what
   the dealer table's `Upside Participation %` column means.

There is **no autocall / KO and no continuous KI** on a Regular BEN — it is European, observed only
on the Valuation Date. Do not reach for the FCN skill's KO/KI vocabulary.

## Section 1 — the standard layout (one row per underlying)
Worst-of baskets get **one row per underlying** and a **blank row between different notes**, exactly
like the FCN tracker.

| Column | Source on the termsheet | Notes |
|---|---|---|
| Trade Date | `DATES` → Trade Date | dealer table calls it **Strike Date** |
| Issue Date | `DATES` → Issue Date | |
| Tenor (mo) | dealer table `Tenor`, e.g. `6M` | measured **Issue Date → Final Valuation Date**, not from Trade Date. Emit the integer only |
| Final Valuation Date | `DATES` → **Expiration Date** | the termsheet defines `Valuation Date = The Expiration Date`. **Not** the Maturity Date |
| Underlying | `UNDERLYING BASKET` → Bloomberg Ticker (i) | e.g. `AMD UW`, `NVDA UW`. Termsheet prints `AMD UW Equity` — drop the trailing `Equity` |
| Currency | `PRODUCT` → Denomination currency | |
| Notional | `PRODUCT` → Issue Size / Denomination | |
| Coupon Flat | back it out of the Redemption Amount: `Max(…, 135.87%)` → **35.87%** | cross-check against the dealer table `Coupon Flat %` |
| Put Strike = Coupon Barrier | `PRODUCT` → Put Strike Level | on a Regular BEN these are the same number; if the dealer table ever shows them differing, that is a **RED** row, not a typo to smooth over |
| Spot Price | `UNDERLYING BASKET` → **Reference Price (Initial) (i)** | usually a 4-dp VWAP, not a round close — the order email says `Spot reference : VWAP` |

## Section 2 — "Only for Rajiv" (this sender only)
**Emit this second block only when the booking email is from Rajiv Sharma.** Other senders get
Section 1 alone. It is his own sheet's layout: **one row per note** (the basket collapses into one
cell), black header bar, no colour coding.

| Column | Value |
|---|---|
| Order Date | Trade Date, `dd-MMM-yy` |
| Product | dealer table `Product`, e.g. `Regular BEN` |
| CCY / Tenor | `USD` / `6M` (with the `M`) |
| Issue Date / Final Valuation Date / Maturity Date | `dd-MMM-yy` |
| Underlying Name | full company names joined `"; "`, in basket order — use the **dealer email's** spelling (`Advanced Micro Devices Inc; NVIDIA Corp`), not the termsheet's legal name |
| Put Strike % / Coupon Flat % / Coupon Barrier % / Upside Participation % | plain numbers, no `%` sign, no percent format |
| Size | `USD 100,000` — currency prefix, as text |
| Strike Price | one line per underlying, `AMD : USD 443.0555` — **Spot × Put Strike %**, 4 dp |

### The Strike Price ambiguity (confirmed 2026-08-21)
Every note in Rajiv's existing sheet had `Put Strike % = 100`, so his `Strike Price` column was
numerically identical to spot and the intent was unreadable. On the first note with a strike ≠ 100%
(XS3454350854, 95%) we filled it with **Spot × Put Strike %** — matching the termsheet's own
`Put Strike Price (i)` column and the column's name. If the user ever says he wants raw spot there,
change `strike_price()` in `scripts/gen_ben_xlsx.py` and update this note.

## Verification — same rule as the FCN reader, do not widen it
The **termsheet is the source of truth**; 核对 means confirming a term against an **independent
second document** — the dealer place-order email table, the email **subject line**, or the filename.
Only that earns a GREEN cell.

- **Internal arithmetic is NOT a cross-check.** `Strike Price = Spot × Put Strike %` always
  reconciles against the termsheet's printed `Put Strike Price` — the maths cannot be wrong, so it
  proves nothing about whether the number was *read* right. Eyeball it to catch a misread, but log
  it as AMBER, never GREEN.
- A term only the termsheet states is **AMBER — "termsheet only"**. That means "no second source",
  not "wrong".
- `Spot Price` is almost always AMBER: the order email gives the spot *basis* (VWAP) but no levels.

## JSON schema
See `example_input.json` — it is the real XS3454350854 note, kept as a worked reference.

```
sender                 str, the booking email's sender display name
notes[]
  isin, issuer, product          str
  trade_date, issue_date,
  final_val_date, maturity_date  str, "dd-MMM-yyyy"
  tenor                          int, months
  currency                       str
  notional                       int
  coupon_flat, put_strike,
  coupon_barrier,
  upside_participation           float, as percent NUMBERS (35.87, not 0.3587)
  underlyings[]
    ticker  str  Bloomberg, "AMD UW"
    short   str  ticker stub for the Strike Price cell, "AMD"
    name    str  full name for the Rajiv row, "Advanced Micro Devices Inc"
    spot    float  Reference Price (Initial)
  verification { "<Column or extra term>": {ref, source} }   -> GREEN
  mismatch     [ "<Column>" ]                                -> RED
  amber_reason { "<Column>": "why there is no second source" }
  extra_audit  { "<term>": "<value>" }  extra Audit-only rows (ISIN, Issuer,
               Maturity Date, Upside Participation, Strike Price, …)
```

`verification` keys must match the Section-1 column names exactly for the Tracker colouring to fire;
keys naming an `extra_audit` term colour that Audit row only.

## Naming & folders
Same convention as the FCN reader: **one output folder per sender**, the file named
`BEN Termsheet Reader - <sender> <timestamp>.xlsx` inside it —
`…\Downloads\Rajiv Sharma\BEN Termsheet Reader - Rajiv Sharma 2026-08-21 164500.xlsx`.
If a sender has more than one product type in the same batch, give each product its own sub-folder
(`…\Rajiv Sharma\BEN\`, `…\Rajiv Sharma\FCN\`). `--outdir` overrides; the default is
`~\Downloads\<sender>\`.
