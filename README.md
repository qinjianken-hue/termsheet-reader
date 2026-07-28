# Termsheet Reader

Two [Claude Code](https://claude.com/claude-code) skills that turn structured-product
termsheets into paste-ready tracker rows, each with a colour-coded verification audit trail.

| Skill | Reads | Produces |
|---|---|---|
| [`fcn-termsheet-reader`](fcn-termsheet-reader/) | FCN / autocallable worst-of equity-linked notes | Excel workbook — **Tracker** (one row per underlying, blank row between notes, columns A–AD or A–AP) + **Audit** |
| [`aqdq-termsheet-reader`](aqdq-termsheet-reader/) | Accumulator / Decumulator ("AQ/DQ") leveraged knock-out forwards | Excel workbook — **Tracker** (one row per trade, columns A–O) + **Audit** |

## The audit trail is the point

The termsheet is always the source of truth; every value is extracted from it. A cell only
goes **green** when the term is confirmed against an *independent second document* — the
dealer's booking / place-order email table, the email subject line, or the filename.

- **green** — cross-checked against the dealer email
- **amber** — read from the termsheet only, no independent second source
- **red** — the email disagrees with the termsheet (investigate before booking)

Internal arithmetic is deliberately **not** treated as a cross-check. Computed
`KO Price = Trade Price × KO Level` always reconciles against the termsheet's own printed
basket price, so it proves nothing about whether the term was *read* correctly. The same
applies to the AQDQ notional reconciliation
(`Notional ≈ Total Expiries × Shares/Day × Trade Price`) — it runs as a self-check and prints
a warning on mismatch, but never colours a cell green.

## Layout

```
fcn-termsheet-reader/
  SKILL.md                     # trigger description + workflow
  references/columns.md        # column dictionary & extraction rules
  references/example_input.json
  scripts/gen_fcn_xlsx.py      # JSON -> Tracker + Audit workbook
  scripts/gen_fcn_csv.py       # deprecated CSV-only predecessor
  scripts/extract_pdf.py       # pypdf text extraction
  scripts/parse_msg.py         # unpack an Outlook .msg

aqdq-termsheet-reader/
  SKILL.md
  references/rules.md          # field dictionary & extraction rules
  references/example_input.json
  scripts/gen_aqdq_xlsx.py
  scripts/gen_aqdq_csv.py      # deprecated
  scripts/extract_pdf.py
  scripts/parse_msg.py
```

## Install

Copy either directory into `~/.claude/skills/` (Windows: `C:\Users\<you>\.claude\skills\`).
Claude Code picks the skill up on the next session.

## Usage

The skills are model-invoked — hand over the termsheet PDFs (or the Outlook `.msg`) and ask
for the tracker rows. Under the hood each generator takes a JSON file:

```bash
python scripts/gen_fcn_xlsx.py  data.json  --outdir <dir> --sender "<booking email sender>"
python scripts/gen_aqdq_xlsx.py data.json  --outdir <dir> --sender "<booking email sender>"
```

Both write `<FCN|AQDQ> Termsheet Reader - <sender> <timestamp>.xlsx`. The JSON schema and the
per-field extraction rules live in each skill's `SKILL.md` and `references/`.

## Requirements

Python 3 with `pypdf`, `extract_msg`, `openpyxl`.

## Notes

Dates are written into the Tracker as `="M/D/YYYY"` formula-text on purpose: these rows get
copied from Excel into Google Sheets, and a bare date would be re-rendered in the local
d/m/yyyy locale and then flipped to m/d by Sheets. The formula-text survives the round trip.

The worked examples under `references/` are real trades, retained because they are what each
skill was validated against.
