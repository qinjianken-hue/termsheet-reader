---
name: aqdq-termsheet-reader
description: >-
  Convert Accumulator / Decumulator ("AQDQ", AQ/DQ) structured-product termsheets and trade
  recaps into a paste-ready Excel workbook (15-column Tracker rows, columns A–O, plus a
  colour-coded verification/Audit trail) for the user's AQDQ tracker (the "AQDQ Record
  Page" / Portfolio Monitor layout). Use this whenever the user supplies an
  accumulator or decumulator termsheet / trade recap (PDF, or bundled in an Outlook .msg) and
  wants it tracked or extracted — leveraged knock-out forwards, "Eq Buy Below Market" / "Sell
  Above Market" trades, daily share accumulation products, anything with a guaranteed / minimum
  share accrual period. Trigger on "AQDQ", "AQ/DQ", "accumulator", "decumulator", 累计期权/累沽,
  dealer trade recaps (e.g. Goldman/UBS/UOBKH "Trade Recap"), or requests to pull knock-out %,
  forward/strike %, spot, total expiries, guarantee days, shares per day, leverage or notional
  out of such a document into a CSV/spreadsheet row. Prefer this skill even when the user doesn't
  say "AQDQ" but the document is clearly an accumulator/decumulator (daily share purchase/sale,
  knock-out, leverage trigger).
---

# AQDQ Termsheet Reader

## What this does
Reads an Accumulator (AQ) / Decumulator (DQ) termsheet or trade recap and emits an **Excel
workbook** (`.xlsx`) with two sheets whose Tracker rows line up 1:1 with columns A–O of the
user's AQDQ tracker, so rows paste straight in:

- **Tracker** — the paste-ready rows. **One row per trade** (single underlying — unlike FCNs
  there is no basket and **no blank row between trades**). Economic-term cells are **colour-coded**
  by whether the term was cross-checked against the dealer email (green = cross-checked, amber =
  termsheet only, red = email disagrees).
- **Audit** — one row per term: the extracted value, the **email reference** it was checked
  against, whether they matched (✓ / — / ✗), and the **source** of the check. This is the
  accountability trail so whoever reviews the booking sees exactly what was verified.

The deterministic parts — date reformatting, decimal handling, the consistency checks, colouring,
building the Audit, file naming — live in `scripts/gen_aqdq_xlsx.py`; your job is the per-dealer
reading (pulling the right values out of the document) **plus recording which terms the dealer
email corroborates**, so the audit trail is populated. `scripts/gen_aqdq_csv.py` is the
deprecated CSV-only predecessor — don't use it unless the user explicitly asks for a bare CSV.

### What "核对" (cross-check) means here — the user's exact definition, do not widen it
The **termsheet is always the source of truth**; every value is extracted from it. **核对 means
confirming an extracted term against an independent second document** — the dealer's place-order /
booking email table, the email **subject line**, or the filename. That is the *only* thing that
earns a green cell.
- **Internal arithmetic is NOT a cross-check.** The notional reconciliation (Notional = Total
  Expiries × Shares/Day × Trade Price) is a self-check you still run to catch a misread, but it
  never greens a cell — the maths can't be wrong, so it proves nothing about whether the term was
  *read* right. It prints as a WARNING on mismatch only.
- **A term the termsheet alone states** (the struck spot when the order says "Spot ref: VWAP", the
  exact notional when the order only gives a Min/Max range, the effective/observation dates) is
  **amber — "termsheet only"**, even though you're confident it's right.
- Only add a `verification` entry for a term that **genuinely appears in the email/subject/order
  table**. That honesty is the whole point. A dealer place-order table typically corroborates
  AQ/DQ, underlying, strike %, KO %, leverage, shares/day, tenor, guarantee period, total days,
  currency and issuer; the struck spot / exact notional / effective date usually stay amber.

## Inputs
- **Termsheet / trade-recap PDFs** — extract text with `scripts/extract_pdf.py` (the Read tool
  cannot render PDFs here — poppler/pdftoppm is missing — so always go through pypdf).
- **An Outlook `.msg`** bundling PDFs and possibly a summary table: `scripts/parse_msg.py`.
- Any table the user provides (email body, tracker export) is a **validation reference** —
  extract from the document, then cross-check.

## Workflow
1. **Collect documents.** For a `.msg`: `python scripts/parse_msg.py "<file.msg>" "<workdir>"`.
2. **Extract text**: `python scripts/extract_pdf.py "<file.pdf>" "<out.txt>"`. Find the general
   terms (trade/effective dates, spot, forward, knock-out, daily number of shares, guaranteed
   period, notional) and the SCHEDULE table (per-period start/end/settlement dates and the
   per-period trading-day counts N).
3. **Read each trade into structured fields** per `references/rules.md` — the authoritative
   field dictionary and extraction rules (Share/Day vs Leverage split, Total Expiries and
   Guarantee Days from the schedule, AQ/DQ classification, label synonyms). **Read it first.**
4. **Write a JSON file** (schema below) and run `python scripts/gen_aqdq_xlsx.py <data.json>`
   (optionally `--outdir <dir>`; default = the user's Downloads, and `--sender "<name>"` — see
   *Naming & folders* below). It converts dates to
   `="M/D/YYYY"` formula-text, keeps levels as calculable decimals, colours the Tracker cells,
   builds the Audit sheet, runs the consistency self-checks (notional reconciliation, KO/Strike
   direction vs AQ/DQ) as printed WARNINGs, and writes
   `AQDQ Termsheet Reader[ - <sender>] <timestamp>.xlsx`.
   - **Read the dealer email/subject/order table and record what it corroborates.** Per trade,
     add a `verification` block mapping only the terms the email actually confirms to
     `{ref, source}` — `ref` = what the email says (e.g. `"128.80%"`, `"90.00%"`, `"2"`, `"250"`,
     `"4W"`, `"DECU"`), `source` = where (e.g. `"dealer order table (Strike %)"`). Those cells go
     green. Terms the email doesn't state (struck spot, exact notional, effective date) — leave
     out; they stay amber. A numeric `ref` that disagrees goes red.
5. **Cross-check and report.** Report what the email corroborated (the green cells), what is
   termsheet-only (amber), and any red Audit row (email ≠ termsheet). Separately, for your own
   confidence, confirm the notional reconciliation (Notional ≈ Total Expiries × Shares/Day ×
   Trade Price) — if it reconciles, the Shares/Day-vs-Leverage split and Total Expiries are
   almost certainly right (this is a self-check, not a green cross-check).

## Naming & folders (sender attribution)
The user reads these booking emails out of Outlook and wants each output file to say **who sent
it**, so a batch stays sortable.

- **Always pass `--sender "<sender name>"`** — the display name of the person who sent the booking
  email (e.g. `--sender "Zhang Wei"`). The file is written as
  `AQDQ Termsheet Reader - Zhang Wei 2026-07-20 112752.xlsx`. Omitting it falls back to the plain
  `AQDQ Termsheet Reader <timestamp>.xlsx`.
- **Batch / multi-item runs (the user's convention):** when the user marks several booking emails
  **unread** and asks to do them together, group the work **by sender → then by product**:
  - one **output folder per sender** (e.g. `…\Downloads\Zhang Wei\`, `…\Downloads\Rajiv Sharma\`);
  - if one sender has **more than one product type**, give each product its **own sub-folder**
    (e.g. `…\Rajiv Sharma\AQDQ\` and `…\Rajiv Sharma\FCN\`). A single sender + single product =
    just the one sender folder, no product sub-folder needed.
  - Pass that folder as `--outdir` and still pass `--sender` (so the filename carries the name too).
  - A sender with several trades of the **same** product goes into **one** workbook (multiple
    Tracker rows), not one file per trade.

## JSON input schema for gen_aqdq_xlsx.py
Dates in **ISO** (yyyy-mm-dd) — the script emits `="M/D/YYYY"`. Prices/levels as **strings** to
preserve exact decimals. The `trade_ref`, `issuer`, `client`, and `verification` fields are
**optional** but drive the Audit sheet and the green/amber/red colouring — supply `verification`
for every term the dealer email corroborates. One object per trade:

```json
{
  "trades": [
    {
      "aq_dq": "AQ",
      "trade_date": "2026-05-19",
      "initial_obs_date": "2026-05-20",
      "gtd_period_end": "2026-06-16",
      "gtd_days": 19,
      "final_obs_date": "2027-05-18",
      "total_expiries": 250,
      "notional": "155706.67",
      "currency": "USD",
      "leverage": 2,
      "shares_per_day": 1,
      "stock": "CRWD",
      "trade_price": "622.8267",
      "ko_level": "1.07",
      "strike_level": "0.747",
      "trade_ref": "SDB1453800981",
      "issuer": "Goldman Sachs",
      "verification": {
        "aq_dq":         {"ref": "ACCU",     "source": "dealer order table (Product)"},
        "strike_level":  {"ref": "74.70%",   "source": "dealer order table (Strike %)"},
        "ko_level":      {"ref": "107.00%",  "source": "dealer order table (KO %)"},
        "leverage":      {"ref": "2",        "source": "dealer order table (Leverage Factor)"},
        "shares_per_day":{"ref": "1",        "source": "order '1 share/day'"},
        "total_expiries":{"ref": "250",      "source": "dealer order table (# of days)"},
        "currency":      {"ref": "USD",      "source": "dealer order table (CCY)"},
        "stock":         {"ref": "CRWD.OQ",  "source": "dealer order table (Underlying)"},
        "trade_date":    {"ref": "19-May-26","source": "order date + subject"}
      }
    }
  ]
}
```

- `verification`: **only** include a term the dealer email / subject / order table actually
  states. Keys among `aq_dq trade_date initial_obs_date gtd_period_end gtd_days final_obs_date
  total_expiries notional currency leverage shares_per_day stock trade_price ko_level
  strike_level`; each `{ref, source}`. A `%` ref like `"90.00%"` reconciles against the decimal
  `0.9`; a plain number ref (`"250"`, `"2"`) matches numerically; a text ref (`"DECU"`, `"4W"`,
  `"12M"`, `"19-May-26"`) is matched as-is (shown green, never a false red). A numeric ref that
  disagrees colours the field **red**. There is **no** `notional`/`trade_price` internal-recon
  cross-check — notional reconciliation is deliberately a self-check only.
- `trade_ref` / `issuer` / `client`: label the Audit sheet only — never written into Tracker rows.

**Gotchas the generator now guards (all print WARNINGs, none block):**
- A percentage `ref` **must keep its `%`** (`"90.00%"`, not `"90.00"`). Without it the ref is read
  as the literal number 90 and the cell goes falsely **red**. Text refs (`"DECU"`, `"4W"`, `"12M"`,
  `"09-Jul-26"`) are fine — matched as-is, never a false red.
- `notional` / prices tolerate commas and a currency prefix (`"USD 152,271.80"` is fine); a value
  that still isn't a number gives a clear error, not a raw `decimal.InvalidOperation`.
- `ko_level` / `strike_level` must be **decimals** (0.90, 1.288). A percent-number slip (128.80)
  is warned as out-of-range — the AQ/DQ direction test alone would NOT catch a strike > 1.
- An unknown `verification` key (typo like `"strike"` for `"strike_level"`) is warned and ignored.
- The notional reconciliation (Total Expiries × Shares/Day × Trade Price) still only warns; it is a
  self-check, never a green cross-check. It does **not** cover GTD Days, GTD Period End or Initial
  Obs Date — those are read from the schedule and only go green if the dealer email states them.
- `scripts/gen_aqdq_csv.py` (deprecated) has NOT been hardened — it still crashes on a comma in
  notional. Use `gen_aqdq_xlsx.py`.

`references/example_input.json` holds the validated CRWD (GS/UOBKH leveraged accumulator) base
worked example (the verification fields are optional add-ons on top).

## Quick rules (full detail in references/rules.md)
- **KO/Strike levels are decimals** (107% → 1.07, 74.70% → 0.747) — the user needs values that
  always compute in Excel; the tracker's % formatting renders them as 107.00% on paste. The
  tracker's KO Price / Strike Price columns (P/Q) are sheet-computed — don't output them.
- **Shares/Day = the base quantity** (shares accrued when at/above the forward); **Leverage =
  below-forward quantity ÷ base** ("1 Share if ≥ Forward, 2 if <" ⇒ Shares/Day 1, Leverage 2).
- **Total Expiries**: explicit "total Scheduled Trading Days" field → else **sum the SCHEDULE's
  per-period N counts** → only as a flagged last resort count Mon–Fri (it misses holidays).
- **Guarantee Days** = sum of N over the guaranteed periods; **Initial Obs Date = Effective
  Date** (daily accrual start), not the first bi-weekly valuation date.
- **AQ** = accumulator / buys below market (strike < spot, KO > spot); **DQ** = decumulator
  (strike > spot, KO < spot).
- Dates ISO in the JSON (the workbook writes `="M/D/YYYY"` formula-text so the Excel → Google
  Sheets copy never gets reformatted); Stock = Yahoo ticker; Currency = the termsheet's
  settlement currency; Notional = the document's stated amount (no commas), else computed.

## Environment notes
- `pypdf` and `extract_msg` are installed; use them (not the Read tool) for PDFs/.msg.
- Prefix runs with `$env:PYTHONIOENCODING='utf-8'` for clean Unicode output.
