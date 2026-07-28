# FCN tracker — column dictionary & extraction rules

Authoritative spec for turning an FCN termsheet into the tracker CSV. Validated 2026-06-03 against
four real termsheets (Vontobel, BBVA, DBS, Nomura).

## Contents
- [Two layouts](#two-layouts)
- [Column dictionary](#column-dictionary)
- [Field-by-field extraction](#field-by-field-extraction)
- [Observation-date placement](#observation-date-placement)
- [Derived price columns](#derived-price-columns)
- [Memory: Yes or No](#memory-yes-or-no)
- [KI Type: EKI vs AKI](#ki-type-eki-vs-aki)
- [Ticker conversion](#ticker-conversion)
- [Currency](#currency)
- [Worked example](#worked-example)

## Two layouts
The sheet has two shapes; pick by **Tenor**. A note's autocall is observed monthly, so an N-month
note has up to N monthly observation dates — the 12-column block can't hold more than 12, so longer
notes use the 24-column block.

- **Tenor ≤ 12M → 12-obs layout.** Fill **A–AD**. Observation columns are C–N (Obs Date 1..12/Final).
- **Tenor > 12M → 24-obs layout.** Fill **A–AP**. Observation columns are C–Z (Obs Date 1..24/Final);
  every field after the observation block shifts right by 12. This matches the user's
  "Reymond Portfolio Monitor - FCN New" template.
- A single CSV has one shape. If a batch mixes tenors, use the 24-obs layout for the whole file
  (`gen_fcn_csv.py` does this automatically when any tenor > 12).

"Over 12 months" means strictly > 12, so a 12M note stays in the 12-obs layout.

## Column dictionary
Fields are identical between layouts; only the observation block width (and therefore the letters
after it) changes. Letters below are the **12-obs** layout, with the 24-obs letter in parentheses
where it differs.

| Field | 12-obs | 24-obs | Notes |
|---|---|---|---|
| Trade date | A | A | |
| Issue date | B | B | |
| Obs Date 1..(12 or 24)/Final Obs | C–N | C–Z | last = Final Obs |
| Maturity Date | O | AA | |
| KO Type | P | AB | `Daily Close` or `End of Period` |
| KI Type | Q | AC | `EKI` / `AKI` / blank |
| Memory | R | AD | `Yes` / `No` |
| Tenor (mo) | S | AE | integer |
| Principal (ccy) | T | AF | note notional |
| Coupon (p.a.) | U | AG | decimal, annualized |
| Stock | V | AH | Yahoo ticker |
| Currency | W | AI | **note** currency |
| Trade Price (ccy) | X | AJ | = Initial Price |
| KO Level (%) | Y | AK | decimal |
| Strike Level (%) | Z | AL | decimal |
| KI Level (%) | AA | AM | decimal / blank |
| KO Price | AB | AN | = X × KO Level |
| Strike Price | AC | AO | = X × Strike Level |
| KI Price | AD | AP | = X × KI Level / blank |

Everything after KI Price — Database Last Close, Underlying KO/KI/Strike Status, Note Status,
DTKO, DTKI, Stock Performance, No./Total Obs, Coupon Received, M2M, Remark, ISIN — is computed by
the sheet from live market data. **Do not fill these.**

A note's terms split into **note-level** fields (A–U / the block up to Coupon) that repeat on every
underlying's row, and **per-underlying** fields (Stock onward) that differ per row.

## Field-by-field extraction
Issuers use different words for the same thing; map by meaning.

- **Trade date** — "Trade Date", "Strike Date", "Initial Fixing", "Initial Valuation Date". These
  usually coincide; if both a trade date and a strike date appear and differ, use the Trade Date.
- **Issue date** — "Issue Date", "Payment Date".
- **Maturity Date** — "Maturity Date", "Repayment Date". Often a couple of business days after the
  final valuation.
- **Obs dates** — the per-period observation / valuation / "Period End" dates from the schedule
  table. The autocall is observed on these (daily-monitored notes still record these monthly dates).
  The final one equals the Final Valuation / Final Fixing date. See placement below.
- **Tenor (mo)** — number of monthly observation periods (the headline "6M / 12M").
- **Principal (ccy)** — the termsheet's stated aggregate amount: "Aggregate Nominal Amount",
  "Aggregate Principal", "Nominal Amount of Series". If the termsheet is a generic Final Terms that
  gives only a per-note denomination / "Nominal Value" (no booked amount), use that and flag it for
  the user — the actual booked notional may differ.
- **Coupon (p.a.)** — see annualization below.
- **KO Type** — `Daily Close` if the autocall/knock-out is observed on every scheduled trading day
  (wording like "each Scheduled Trading Day", "Daily"); `End of Period` if observed only on the
  period-end/valuation date. Strip any "Memory" word here into the Memory field.
- **Trade Price** — the per-underlying "Initial Price" / "Initial Reference Price" / "RI Initial
  Value", copied verbatim (it is in the stock's own currency, typically USD).
- **KO / Strike / KI Level** — the percentages from the basket table, as decimals:
  - KO Level ← "Autocall Trigger Level", "Autocall Level", "Knock-Out Price" % (e.g. 105% → 1.05).
  - Strike Level ← "Strike", "Put Strike", "Forward Price" % (e.g. 62.3% → 0.623).
  - KI Level ← "Knock-In Level/Price" % (e.g. 55% → 0.55); blank if the note has no KI.

## Observation-date placement
The "Final Obs" column always holds the **final/maturity observation**; earlier observations fill
from the first observation column; any leftover middle slots stay blank. `gen_fcn_csv.py` does this
for you — just pass `obs_dates` in chronological order with the final date last.

Examples (12-obs layout, observation columns C–N):
- 6M note, 6 observations → C,D,E,F,G filled, H–M blank, **N = final**.
- 9M note, 9 observations → C–J filled, K–M blank, **N = final**.
- 12M note, 12 observations → C–N all filled (no gap).

24-obs layout (C–Z), e.g. an 18M note (18 observations) → C–S filled, T–Y blank, **Z = final**.

## Derived price columns
Compute, don't copy (the script uses exact decimals):
- KO Price = Trade Price × KO Level
- Strike Price = Trade Price × Strike Level
- KI Price = Trade Price × KI Level (blank if no KI)

These must equal the price levels printed in the termsheet's basket table (e.g. Strike Price column).
If they don't reconcile to ~4 dp, re-check the Initial Price or the level — this is the best
self-check that the extraction is right.

## Memory: Yes or No
"Memory" here is about the **autocall/knock-out**, not the coupon (FCN coupons are fixed). Decide
from the termsheet's mechanics, not just a word in the title:
- **Yes** — each underlying's knock-out is *latched/remembered*: a share that reaches its autocall
  level on any observation date stays "knocked out", and the note autocalls once **every** underlying
  has individually knocked out, **not necessarily on the same date**. Tell-tale wording: "Memorized
  Underlying", "Memory Knock-Out Event", "Individual Knock-Out Event can only occur ... one time",
  "the Highest RI Value ... does not need to occur on the same ... date".
- **No** — the autocall requires **all** underlyings to be at/above their level **on the same
  observation date**, or there is only a single end-of-period observation.

## KI Type: EKI vs AKI
Set by **when** the knock-in barrier is observed:
- **EKI** (European) — observed **only at maturity / the final valuation date**. Wording: "European
  Knock-In", "Knock-In Determination Day: Final Valuation Date", KI tested on "the Redemption
  Valuation Date". A dealer table that lists KI Type as "At Maturity" means EKI.
- **AKI** (American) — observed **continuously / on every scheduled trading day** over a period.
  Wording: "American Knock-In", KI tested on "any" day / "each Scheduled Trading Day".
- Blank if the note has no separate knock-in barrier (the strike is the only downside reference).

## Ticker conversion
Stock (V) is the **Yahoo Finance** ticker. Termsheets quote Reuters RIC and/or Bloomberg codes —
convert:
- **US listings** → bare ticker. RIC `.OQ`/`.O`/`.N` (Nasdaq/NYSE) and Bloomberg `UW`/`UQ`/`UN` all
  drop their suffix: `AMD.OQ` / `AMD UW` → `AMD`; `IBM.N` / `IBM UN` → `IBM`; `TSM.N` → `TSM`.
- **Non-US** → Yahoo exchange suffix. Common ones:

  | Exchange | RIC / BBG hint | Yahoo suffix | Example |
  |---|---|---|---|
  | Hong Kong | `.HK` / `HK` | `.HK` (4-digit, zero-pad) | `700 HK` → `0700.HK` |
  | Tokyo | `.T` / `JT` | `.T` | `7203 JT` → `7203.T` |
  | Korea (KRX) | `.KS` / `KS` | `.KS` | `005930 KS` → `005930.KS` |
  | Taiwan | `.TW` / `TT` | `.TW` | `2330 TT` → `2330.TW` |
  | London | `.L` / `LN` | `.L` | `HSBA LN` → `HSBA.L` |
  | Shanghai / Shenzhen | `.SS` / `.SZ` | `.SS` / `.SZ` | `600519 CH` → `600519.SS` |
  | Australia | `.AX` / `AT` | `.AX` | `BHP AT` → `BHP.AX` |

  If an exchange isn't listed, infer the Yahoo suffix from the listing exchange and note it.

## Currency
Currency (W) = the **note / termsheet currency** — the "Specified Currency", "Settlement Currency",
or "Reference Currency". **Not** the underlying stock's trading currency. A SGD-quanto note on US
stocks has W = SGD even though each Trade Price (X) is a USD share price. The note currency is also
the currency of the Principal.

## Worked example
BBVA, USD, 6M, worst-of CRWV/NVDA/PLTR/QCOM, autocall 105%, strike 62.3%, KI 55% observed at
maturity, memory autocall, coupon 30% p.a. (2.5%/month × 12):

- Note-level: Trade 2026-05-29, Issue 2026-06-12, Maturity 2026-12-16, KO Type `Daily Close`,
  KI Type `EKI`, Memory `Yes`, Tenor `6`, Principal `430000`, Coupon `0.30`, Currency `USD`.
- obs_dates (6): 2026-07-13, 08-12, 09-14, 10-12, 11-12, **2026-12-14 (final)** → C–G + N.
- CRWV: Trade Price 109.53, KO 1.05, Strike 0.623, KI 0.55 →
  KO Price 115.0065, Strike Price 68.23719, KI Price 60.2415 — matches the termsheet's
  Autocall 115.0065 / Strike 68.2372 / Knock-in 60.2415.

`references/example_input.json` holds all four validated notes; run it through `gen_fcn_csv.py` to
see the exact output shape.
