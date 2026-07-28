---
name: fcn-termsheet-reader
description: >-
  Convert FCN (Fixed Coupon Note) structured-product termsheets into a paste-ready Excel workbook
  for the user's FCN tracking sheet (the "FCN New" layout — columns A–AD, or A–AP for tenors over 12
  months), with a colour-coded verification/audit trail showing which values were cross-checked and
  against what. Use this whenever the user supplies one or more FCN termsheets or factsheets (as
  PDFs, or bundled in an Outlook .msg) and wants them tracked, booked, or turned into the
  tracker/CSV/Excel format. Trigger on "FCN tracking", "termsheet reader", autocallable / worst-of / fixed-coupon /
  knock-in (KI) / knock-out (KO / autocall) notes, or any request to pull note terms (strike,
  coupon, KO/KI levels, observation or valuation dates, initial prices, ISIN) out of a
  structured-note termsheet into a CSV or spreadsheet. Also triggers when the user forwards a
  dealer "trade booking" email with several notes plus a summary table to validate against.
  Prefer this skill even when the user doesn't say "FCN" but the document is clearly an
  autocallable worst-of equity-linked note.
---

# FCN Termsheet Reader

## What this does
Reads FCN (Fixed Coupon Note) termsheets and emits an **Excel workbook** (`.xlsx`) with two sheets:

- **Tracker** — columns line up 1:1 with the user's FCN tracking sheet, so rows paste straight in.
  These are **worst-of baskets**, so there is **one row per underlying** and a **blank row between
  different notes**. Economic-term cells are **colour-coded** by whether the term was cross-checked
  against the dealer email (green = cross-checked, amber = termsheet only, red = email disagrees).
- **Audit** — one row per term: the extracted value, the **email reference** it was checked against,
  whether they matched (✓ / — / ✗), and the **source** of the check. This is the accountability
  trail so whoever reviews the booking sees exactly what was verified against the email.

### What "核对" (cross-check) means here — the user's exact definition, do not widen it
The **termsheet is always the source of truth**; every value is extracted from it. **核对 means
confirming an extracted term against an independent second document** — the dealer's place-order /
booking email table, the email **subject line**, or the filename. That is the *only* thing that
earns a green cell. In particular:
- **Internal arithmetic is NOT a cross-check.** Computed KO/Strike Price = Trade Price × level always
  reconciles against the termsheet's own printed basket price — but the maths can't be wrong, so it
  proves nothing about whether the term was *read* correctly. Never green a cell on that basis. (You
  should still eyeball the reconciliation yourself while extracting to catch a misread, but it does
  not go in the audit as a check.)
- **A term the termsheet alone states** (e.g. KO Type inferred from "every Scheduled Trading Day",
  the coupon, the maturity date, memory) is **amber — "termsheet only"**, even though you're
  confident it's right. Amber means "read correctly, but no independent second source", not "wrong".
- Only add a `verification` entry for a term that **genuinely appears in the email/subject/booking
  table**. That honesty is the whole point.

The deterministic parts — computing KO/Strike/KI Price, choosing the layout, colouring, building the
Audit, naming the file — are handled by `scripts/gen_fcn_xlsx.py`. Your job is the per-issuer reading
(pulling the right values out of each termsheet) **plus recording which terms the dealer email
corroborates**, so the audit trail is populated.

Output is **Excel only** (the user copies rows from the Tracker sheet straight into Google Sheets).
`scripts/gen_fcn_csv.py` is the deprecated CSV-only predecessor — don't use it unless asked.

## Inputs
- **Termsheet PDFs** — one per note. Extract text with `scripts/extract_pdf.py`. The Read tool
  cannot render PDFs in this environment (poppler/pdftoppm is missing), so always go through pypdf.
- **An Outlook `.msg`** bundling several termsheet PDFs and (usually) a summary table in the email
  body. Run `scripts/parse_msg.py` to print the body and save the PDF attachments.
- A table in the email body (or one the user pastes) is a **validation reference, not the source of
  truth** — extract from the PDF, then cross-check against the table.

## Workflow
1. **Collect termsheets.** Given a `.msg`, run
   `python scripts/parse_msg.py "<file.msg>" "<workdir>"` — prints the body and saves each PDF
   attachment + the HTML body into `<workdir>`.
2. **Extract text** from each PDF: `python scripts/extract_pdf.py "<file.pdf>" "<out.txt>"`. For
   long termsheets, search the text for the basket table (initial prices, levels), the schedule
   (observation/valuation dates), the coupon, the KI/KO definitions, and the headline dates.
3. **Read each note into structured fields** per `references/columns.md` — the authoritative column
   dictionary and extraction rules (mappings, EKI vs AKI, Memory Yes/No, ticker and currency
   conventions, observation-date placement). **Read it before extracting.**
4. **Write a JSON file** of the notes (schema below) and run `python scripts/gen_fcn_xlsx.py
   <data.json>` (optionally `--outdir <dir>`; default = the user's Downloads, and
   `--sender "<name>"` — see *Naming & folders* below). It computes
   KO/Strike/KI Price with exact decimals, auto-selects the 12- vs 24-observation layout from
   tenor, colours the Tracker cells, builds the Audit sheet, writes
   `FCN Termsheet Reader[ - <sender>] <timestamp>.xlsx`, and prints the path.
   - **Read the dealer email/subject/booking table and record what it corroborates.** Per note, add
     a `verification` block mapping only the terms the email actually confirms to `{ref, source}` —
     `ref` = what the email says (e.g. `"46.04%"`, `"Daily"`, `"400,000"`, `"7 Jul 2026"`),
     `source` = where (e.g. `"dealer order email"`, `"email subject"`, `"booking table"`). Those
     cells go green. Terms the email doesn't mention: leave out — they stay amber ("termsheet
     only"). A numeric `ref` that disagrees with the extracted value goes red. Typical corroboration:
     a UBS place-order email restates strike / KO type / KO / coupon / tenor / notional; a
     "Trade Reporting" email only gives the booking table (notional, trade & settlement dates, ISIN)
     plus the subject line (strike %, tenor, underlyings) — so those notes get far fewer greens.
5. **Cross-check and report.** Compare every extracted term to the dealer email and record the
   matches in `verification`. Separately (for your own confidence, not the audit) eyeball that
   computed KO/Strike Price = Trade Price × level equals the termsheet's printed basket price — if
   that *doesn't* reconcile you misread the initial price or a level, so fix it. Report what the
   email corroborated, what is termsheet-only, and flag any red Audit row (email ≠ termsheet).

## Naming & folders (sender attribution)
The user reads these booking emails out of Outlook and wants each output file to say **who sent
it**, so a batch stays sortable.

- **Always pass `--sender "<sender name>"`** — the display name of the person who sent the booking
  email (e.g. `--sender "Wang Linjing"`). The file is written as
  `FCN Termsheet Reader - Wang Linjing 2026-07-20 095720.xlsx`. Omitting it falls back to the plain
  `FCN Termsheet Reader <timestamp>.xlsx`.
- **Batch / multi-item runs (the user's convention):** when the user marks several booking emails
  **unread** and asks to do them together, group the work **by sender → then by product**:
  - one **output folder per sender** (e.g. `…\Downloads\Wang Linjing\`, `…\Downloads\Zhang Wei\`);
  - if one sender has **more than one product type**, give each product its **own sub-folder**
    (e.g. `…\Zhang Wei\FCN\` and `…\Zhang Wei\AQDQ\`). A single sender + single product = just the
    one sender folder, no product sub-folder needed.
  - Pass that folder as `--outdir` and still pass `--sender` (so the filename carries the name too).
  - A sender with several notes goes into **one** workbook (multiple Tracker rows), not one file
    per note.

## JSON input schema for gen_fcn_xlsx.py
One object per note in `fcns`. Note-level fields apply to every underlying; `underlyings` carries
the per-stock values. Pass prices/levels/coupon as **strings** to preserve exact decimals. The
`isin`, `issuer`, and `verification` fields are **optional** but drive the Audit sheet and the
green/amber/red colouring — supply `verification` for every term the dealer email corroborates.

```json
{
  "fcns": [
    {
      "isin": "XS0000000000",
      "issuer": "UBS",
      "trade_date": "2026-05-29",
      "issue_date": "2026-06-12",
      "obs_dates": ["2026-07-13", "...chronological...", "2026-06-14"],
      "maturity_date": "2027-06-16",
      "ko_type": "Daily Close",
      "ki_type": "",
      "memory": "No",
      "tenor_mo": 12,
      "principal": 50000,
      "coupon_pa": "0.12",
      "ko_level": "1",
      "strike_level": "0.5",
      "ki_level": "",
      "currency": "USD",
      "verification": {
        "ko_type":      {"ref": "Daily",      "source": "dealer order email"},
        "ko_level":     {"ref": "100%",       "source": "dealer order email"},
        "strike_level": {"ref": "50%",        "source": "dealer order email"},
        "coupon_pa":    {"ref": "12%",        "source": "dealer order email"},
        "principal":    {"ref": "50,000",     "source": "booking table + order email"},
        "tenor_mo":     {"ref": "12",         "source": "dealer order email"},
        "trade_date":   {"ref": "29 May 2026","source": "booking table"},
        "issue_date":   {"ref": "12 Jun 2026","source": "booking table (Settlement Date)"}
      },
      "underlyings": [
        {"stock": "AMAT", "trade_price": "450.06"},
        {"stock": "IBM",  "trade_price": "297.80"}
      ]
    }
  ]
}
```

- `verification`: **only** include a term here if the dealer email / subject / booking table
  actually states it — that is what "核对" means (see the definition above). Keys among `ko_type
  ko_level strike_level ki_level coupon_pa principal tenor_mo memory trade_date issue_date
  maturity_date`; each `{ref, source}`. `ref` = the email's wording (a `%` ref like `"50%"`
  reconciles against the decimal `0.5`; a date/text ref like `"29 May 2026"` / `"Daily"` is matched
  as-is, not numerically). A numeric `ref` that disagrees colours the field **red**. There are **no**
  `*_price_ref` fields — internal price reconciliation is deliberately not treated as a cross-check.

- `obs_dates`: every observation/valuation date in chronological order; the **last** entry is the
  final/maturity observation. The script puts the last one in the "Final Obs" column and fills the
  earlier ones from the first observation column, leaving middle slots blank.
- `ki_type`: `"EKI"`, `"AKI"`, or `""`. `ki_level`: decimal string or `""` (blank ⇒ KI Price blank).
- `tenor_mo` drives the layout: ≤ 12 ⇒ 12-obs (A–AD); > 12 ⇒ 24-obs (A–AP). If a batch mixes both,
  the whole file uses the 24-obs layout.
- `isin` / `issuer`: label the Audit sheet only — **never** written into the Tracker rows (the
  tracker's ISIN column is computed by the user's sheet).
- A worked `references/example_input.json` reproduces the four notes this skill was validated on
  (base schema; the verification fields are optional add-ons on top).

## Quick rules (full detail in references/columns.md)
- Trade Price (X) = the termsheet's Initial Price, verbatim. KO/Strike/KI Price = X × level.
- Levels and coupon are stored as **decimals** (105% → 1.05, 50% → 0.5, 12% p.a. → 0.12).
- Coupon (p.a.) = per-period coupon × periods per year (monthly ⇒ ×12); or take the stated p.a.
- Currency = the **note** currency, not the stock's trading currency.
- Stock = Yahoo Finance ticker (US ⇒ bare; non-US ⇒ exchange suffix).
- Memory = Yes when each underlying's knock-out is latched/remembered (needn't all hit on the same
  date); No otherwise.
- KI Type: EKI = observed only at maturity ("European" / "At Maturity"); AKI = continuous/daily
  ("American").
- Dates as yyyy-mm-dd — the script rejects anything else. In the Tracker sheet it writes date cells
  as `="yyyy-mm-dd"` formula-text on purpose: the user copies rows from Excel into Google Sheets,
  and a bare date would get re-displayed by Excel in d/m/yyyy locale form and then flipped to m/d by
  Sheets. The formula-text survives the round trip unchanged. Don't "simplify" this away.
- Status / last-close / ISIN columns are computed by the sheet — don't fill them (ISIN goes in the
  Audit sheet's label column instead, via the note's `isin` field).

## Environment notes
- `pypdf`, `extract_msg`, and `openpyxl` are installed; use them (don't rely on the Read tool for
  PDFs/.msg). Termsheets also arrive as `.docx` (e.g. MS, Citi) — extract those by unzipping
  `word/document.xml` and stripping tags (python-docx is **not** installed).
- Prefix script runs with `$env:PYTHONIOENCODING='utf-8'` so Unicode in termsheets prints cleanly.
