---
name: ben-termsheet-reader
description: >-
  Convert BEN (Bonus Enhanced Note / "Regular BEN") structured-product termsheets into a paste-ready
  Excel workbook for the user's BEN tracker, with a colour-coded verification/audit trail showing
  which values were cross-checked against the dealer place-order email. Use this whenever the user
  supplies a BEN termsheet (PDF, or bundled in an Outlook .msg) and wants it tracked, booked, or
  turned into the tracker/Excel format. Trigger on "BEN", "Bonus Note", "Bonus Enhanced Note",
  "Regular BEN", worst-of bonus notes, or any request to pull note terms (put strike, coupon flat,
  coupon barrier, upside participation, spot, strike price, final valuation date) out of such a
  termsheet into a spreadsheet. Prefer this skill even when the user doesn't say "BEN" but the
  document is clearly a worst-of Bonus Note — redemption of the form
  "Max(Worst Underlying Performance, 1XX%)" with physical delivery below a put strike, and NO
  autocall/KO and no continuous knock-in (that would be an FCN — use fcn-termsheet-reader instead).
---

# BEN Termsheet Reader

## What this does
Reads BEN (Bonus Enhanced Note) termsheets and emits an **Excel workbook** (`.xlsx`) with two sheets:

- **Tracker** —
  - *Section 1*: the standard 10-column layout, **one row per underlying** (these are worst-of
    baskets) with a blank row between different notes. Cells are colour-coded by whether the term
    was cross-checked against the dealer email (green = cross-checked, amber = termsheet only,
    red = email disagrees).
  - *Section 2*: **"Only for Rajiv"** — his own 14-column layout, one row per note, black header,
    no colour coding. **Emitted only when the sender is Rajiv.**
- **Audit** — one row per term: the extracted value, the **email reference** it was checked against,
  whether they matched (OK / - / MISMATCH), and the **source** of the check.

## Is it actually a BEN?
A Regular BEN is **European** — no autocall, no KO, no continuous KI. Tell-tale on the termsheet:

```
Redemption Amount
  a) if all Reference Price (Final) (i) are equal to or above their respective Put Strike Price (i)
     Denomination x [ Max( Worst Underlying Performance (Final), 135.87% ) ]
  b) if at least one Reference Price (Final) (i) is below its Put Strike Price (i)
     the Deliverable Assets
```

If instead you see observation dates, a KO/autocall level, a coupon per period or a knock-in
observed daily → that is an FCN, use `fcn-termsheet-reader`.

## Inputs
- **Termsheet PDF** — extract text with `scripts/extract_pdf.py`. The Read tool cannot render PDFs
  in this environment (poppler/pdftoppm is missing), so always go through pypdf.
- **The dealer place-order email** — Rajiv's "Place Order: BEN on …" mail carries a full quote table
  (Quote ID, Product, CCY, Tenor, Strike/Issue/Final Valuation/Maturity Date, underlyings, Put
  Strike %, Coupon Flat %, Coupon Barrier %, Upside Participation %, Min/Max Notional, Issuer,
  Interbank Price %). **This is the cross-check source and is what makes cells green** — always go
  find it, it is a separate email from the one carrying the termsheet.
- **An Outlook `.msg`** bundling the PDF — run `scripts/parse_msg.py` to print the body and save the
  attachments.

Reading Outlook directly: the M365 connector is blocked on this machine — use local Outlook COM
(`win32com`), and save attachments with `att.SaveAsFile` (never `OpenSharedItem`, it hangs).
The termsheet usually arrives in **Inbox\Trade Booking**; the place-order email sits in **Inbox**.

## Workflow
1. **Collect the termsheet** (and, given a `.msg`, run
   `python scripts/parse_msg.py "<file.msg>" "<workdir>"`).
2. **Find the matching place-order email.** Match on the client account, the subject's trade date,
   and the underlyings. Without it every cell is amber and the audit is worthless.
3. **Extract text**: `python scripts/extract_pdf.py "<file.pdf>" "<out.txt>"`. The fields you need
   are all on pages 1–2: the `PRODUCT`, `DATES`, `UNDERLYING BASKET` and `Redemption Amount` blocks.
4. **Read the note into structured fields** per `references/columns.md` — the authoritative column
   dictionary and extraction rules. **Read it before extracting.**
5. **Write a JSON file** (schema in `references/columns.md`, worked example in
   `references/example_input.json`) and run
   `python scripts/gen_ben_xlsx.py <data.json> --sender "<sender name>"`
   (optionally `--outdir <dir>`; default is `~\Downloads\<sender>\`). It computes the Strike Price,
   collapses the basket for the Rajiv block, colours the Tracker, builds the Audit sheet, writes
   `BEN Termsheet Reader - <sender> <timestamp>.xlsx` **inside a per-sender folder**, and prints the
   path.
   - **Record only what the email genuinely corroborates** in `verification` — see below.
6. **Report** what the email corroborated, what is termsheet-only, and flag any red row.

## What "核对" (cross-check) means here — the user's exact definition, do not widen it
The **termsheet is always the source of truth**; every value is extracted from it. **核对 means
confirming an extracted term against an independent second document** — the dealer place-order
table, the email **subject line**, or the filename. That is the *only* thing that earns a green cell.

- **Internal arithmetic is NOT a cross-check.** `Strike Price = Spot × Put Strike %` always
  reconciles against the termsheet's printed `Put Strike Price` — the maths can't be wrong, so it
  proves nothing about whether the number was *read* correctly. Eyeball it yourself to catch a
  misread, but never green a cell on that basis.
- A term the termsheet alone states is **amber — "termsheet only"**, even when you're confident.
- `Spot Price` is normally amber: the order email gives the basis (`Spot reference : VWAP`) but no
  levels.

## Two things to say in every report
1. **Coupon Flat is a flat maturity payout, not p.a.** A 6M note quoting 35.87% pays 135.87% at
   maturity if the barrier holds. Never annualise it.
2. **The upside is uncapped** — `Max(worst performance, 100% + coupon flat)` means above the flat
   coupon the note tracks the worst performer at the Upside Participation rate (100% so far).

Also check the tenor: it runs **Issue Date → Final Valuation Date**, so a "6M" note traded 20-Aug
with a 03-Mar Final Valuation Date is 6 months from the 03-Sep issue, ~6.4 months from trade.

## Naming & folders
**One output folder per sender**, the workbook inside it —
`…\Downloads\Rajiv Sharma\BEN Termsheet Reader - Rajiv Sharma 2026-08-21 164500.xlsx`.
If a sender has several product types in one batch, give each product its own sub-folder
(`…\Rajiv Sharma\BEN\`, `…\Rajiv Sharma\FCN\`). A sender with several notes goes into **one**
workbook (multiple rows), not one file per note.
