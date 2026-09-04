# -*- coding: utf-8 -*-
"""Generate the FCN tracker WORKBOOK (.xlsx) from structured note data.

Usage:
    python gen_fcn_xlsx.py <data.json> [--outdir DIR]

Two sheets:
  * "Tracker" - the paste-ready rows (one row per underlying, blank row between notes,
                dates as ="yyyy-mm-dd" formula-text so the Excel -> Google Sheets copy
                never reformats). Economic-term cells are colour-coded by whether the term
                was CROSS-CHECKED AGAINST THE DEALER EMAIL.
  * "Audit"   - one row per term: the extracted value, the email reference it was checked
                against, whether they matched, and the SOURCE of the check. The
                accountability trail - who checked what, against what.

WHAT "核对" (cross-check) MEANS HERE  (this is the user's definition - do not widen it):
  The termsheet is always the SOURCE OF TRUTH; every value is extracted from it.
  "核对" means confirming an extracted term against an INDEPENDENT SECOND DOCUMENT - the
  dealer's place-order / booking email table, the email subject line, or the filename.
  Internal arithmetic (computed KO/Strike price = X x level) is NOT a cross-check: the
  maths cannot be wrong, so it proves nothing about whether the term was read correctly.
  A term is only green if a real email/booking/subject reference corroborates it.

Colour key (Tracker + Audit "Match"):
    GREEN  ✓  = cross-checked against the dealer email / booking table / subject
    AMBER  —  = extracted from the termsheet only (incl. computed prices) - NOT cross-checked
    RED    ✗  = the email says something different from the termsheet - investigate

Driven by optional JSON fields:
  note-level:
     "isin", "issuer"            -> labels for the Audit sheet (never written into Tracker)
     "investor": "EM4 Xu Chaoping"
         -> the tracker's Investor column (AQ in the 12-obs layout, BC in the 24-obs), read off
            the booking table's TR Code + Client Name. Written on the FIRST underlying row of the
            note only. AE-AP are emitted as empty grey spacer columns so Investor lands on its
            real letter; those spacers are the sheet's own formulas, so paste A-AD and AQ only.
     "verification": {           -> ONLY email-corroborated terms go here; field -> {ref, source}
         "ko_type":     {"ref": "Daily",   "source": "dealer order email"},
         "strike_level":{"ref": "46.04%",  "source": "email subject"},
         "principal":   {"ref": "400,000", "source": "booking table"},
         "trade_date":  {"ref": "7 Jul 2026", "source": "booking table"},
         ...any of: ko_type ko_level strike_level ki_level coupon_pa principal
                    tenor_mo memory trade_date issue_date maturity_date
     }
  A field WITH a verification entry is green (email source shown); WITHOUT one it is amber
  ("termsheet only"). A numeric ref that disagrees with the extracted value goes RED.
  DO NOT add a verification entry unless the term genuinely appears in the email/subject/
  booking table - that is the whole point of the accountability trail.
"""
import os, re, sys, json, argparse
from decimal import Decimal as D
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

FIELDS_AFTER_OBS = ["Maturity Date", "KO Type", "KI Type", "Memory", "Tenor (mo)",
    "Principal (ccy)", "Coupon (p.a.)", "Stock", "Currency", "Trade Price (ccy)",
    "KO Level (%)", "Strike Level (%)", "KI Level (%)", "KO Price", "Strike Price", "KI Price"]

# which economic-term columns get a green/amber/red status (Stock/Currency stay neutral)
PAINT_COLS = ["Trade date", "Issue date", "Maturity Date", "KO Type", "KI Type", "Memory",
    "Tenor (mo)", "Principal (ccy)", "Coupon (p.a.)", "Trade Price (ccy)",
    "KO Level (%)", "Strike Level (%)", "KI Level (%)", "KO Price", "Strike Price", "KI Price"]

# column name -> verification key (columns absent here are always amber: computed / determined)
VER_KEY = {"Trade date": "trade_date", "Issue date": "issue_date", "Maturity Date": "maturity_date",
    "KO Type": "ko_type", "Tenor (mo)": "tenor_mo", "Principal (ccy)": "principal",
    "Coupon (p.a.)": "coupon_pa", "KO Level (%)": "ko_level", "Strike Level (%)": "strike_level",
    "KI Level (%)": "ki_level", "Memory": "memory", "KI Type": "ki_type"}

# ---- styling ---------------------------------------------------------------
GREEN = PatternFill("solid", fgColor="C6EFCE")
AMBER = PatternFill("solid", fgColor="FFEB9C")
RED   = PatternFill("solid", fgColor="FFC7CE")
HDR   = PatternFill("solid", fgColor="1F4E78")
INVHDR = PatternFill("solid", fgColor="7F6000")   # Investor header (col AQ / BC)
SPACERHDR = PatternFill("solid", fgColor="A6A6A6")  # the sheet-computed columns, emitted blank
NOTEBAR = PatternFill("solid", fgColor="DDEBF7")
GREEN_FONT = Font(color="006100")
AMBER_FONT = Font(color="9C6500")
RED_FONT   = Font(color="9C0006")
HDR_FONT   = Font(color="FFFFFF", bold=True)
BOLD       = Font(bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

OK, EXTRACTED, FAIL = "ok", "extracted", "fail"
FILL = {OK: GREEN, EXTRACTED: AMBER, FAIL: RED}
FONT = {OK: GREEN_FONT, EXTRACTED: AMBER_FONT, FAIL: RED_FONT}
MARK = {OK: "✓", EXTRACTED: "—", FAIL: "✗"}


def dec(x):
    if x is None or x == "":
        return None
    return D(str(x))


def dec_ref(x):
    """Parse a numeric ref (strips %, commas, currency). Returns None for anything with
    letters - a date ('7 Jul 2026') or a label ('Daily') is compared textually, never
    numerically (else '7 Jul 2026' -> 72026 would spuriously mismatch the ISO date)."""
    if x is None or x == "":
        return None
    raw = str(x)
    if re.search(r"[A-Za-z]", raw):
        return None
    s = re.sub(r"[^0-9.\-]", "", raw)
    if s in ("", "-", ".", "-."):
        return None
    try:
        return D(s)
    except Exception:
        return None


def num(x):
    if isinstance(x, D):
        return float(x)
    return x


def norm(x):
    if isinstance(x, D):
        return format(x.normalize(), "f")
    return "" if x is None else str(x)


def cmp_nums(extracted, ref):
    """Numeric forms of (extracted, ref), handling '%' refs (90.00% -> 0.9)."""
    ex_n, ref_n = dec_ref(extracted), dec_ref(ref)
    if ref_n is not None and "%" in str(ref):
        ref_n = ref_n / D(100)
    if ex_n is not None and "%" in str(extracted):
        ex_n = ex_n / D(100)
    return ex_n, ref_n


def field_status(extracted, ver):
    """Green iff an email verification entry exists; red if a numeric ref disagrees."""
    if not ver:
        return EXTRACTED, None, "termsheet only (not in email)"
    ref = ver.get("ref", "")
    source = ver.get("source", "") or "email"
    ex_n, ref_n = cmp_nums(extracted, ref)
    if ex_n is not None and ref_n is not None and abs(ex_n - ref_n) > D("0.0001"):
        return FAIL, ref, source
    return OK, ref, source


def date_formula(d, field="date"):
    if d is None or d == "":
        return ""
    if not ISO_DATE.match(str(d)):
        raise ValueError(f"{field} {d!r} is not yyyy-mm-dd - fix the JSON, never pass "
                         f"locale-ambiguous formats like 4/6/2026")
    return f'="{d}"'


def place_obs(obs_dates, n_slots):
    slots = [""] * n_slots
    if not obs_dates:
        return slots
    if len(obs_dates) > n_slots:
        raise ValueError(f"{len(obs_dates)} observation dates exceed {n_slots} slots")
    *earlier, final = obs_dates
    for i, d in enumerate(earlier):
        slots[i] = d
    slots[-1] = final
    return slots


def header_row(n_obs):
    obs = [f"Obs Date {i}" for i in range(1, n_obs)] + [f"Obs Date {n_obs}/Final Obs"]
    return ["Trade date", "Issue date"] + obs + FIELDS_AFTER_OBS


# Between the data block and Investor the tracker has 12 columns the SHEET computes itself.
# They are emitted as EMPTY spacer columns purely so Investor lands on its real tracker
# letter (AQ in the 12-obs layout, BC in the 24-obs) - the user asked for that alignment.
# They are left blank on purpose: paste them over a live tracker and you wipe its formulas.
SHEET_COMPUTED_HEADERS = ["Database Last Close (LC)", "Underlying KO Status",
    "Underlying KI Status", "Underlying Strike Status", "Note Status", "DTKO", "DTKI",
    "Stock Performance", "No. of Obs", "Total Obs", "Coupon Received", "M2M"]
SHEET_COMPUTED_COLS = len(SHEET_COMPUTED_HEADERS)


def investor_col_letter(n_data_cols):
    """Tracker letter the Investor value belongs in: 12-obs -> AQ, 24-obs -> BC."""
    return openpyxl.utils.get_column_letter(n_data_cols + SHEET_COMPUTED_COLS + 1)


def isin_col_letter(n_data_cols):
    """Tracker letter for the optional ISIN column, one right of Investor:
    12-obs -> AR, 24-obs -> BD. Only emitted with --tracker-isin (Shaun's layout)."""
    return openpyxl.utils.get_column_letter(n_data_cols + SHEET_COMPUTED_COLS + 2)


# ---- Tracker sheet ---------------------------------------------------------

def build_tracker(wb, fcns, n_obs, tracker_isin=False):
    ws = wb.active
    ws.title = "Tracker"
    hdr = header_row(n_obs)
    inv_letter = investor_col_letter(len(hdr))
    isin_letter = isin_col_letter(len(hdr))
    ws.append(hdr + [f"({h}) - sheet-computed, left blank" for h in SHEET_COMPUTED_HEADERS]
                  + [f"Investor (col {inv_letter})"]
                  + ([f"ISIN (col {isin_letter})"] if tracker_isin else []))
    for c in ws[1]:
        c.fill = HDR; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    col = {name: hdr.index(name) + 1 for name in hdr}
    # empty spacer columns so Investor sits on its real tracker letter
    for i in range(len(hdr) + 1, len(hdr) + 1 + SHEET_COMPUTED_COLS):
        ws.cell(row=1, column=i).fill = SPACERHDR
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 11
    inv_col = len(hdr) + SHEET_COMPUTED_COLS + 1
    ws.cell(row=1, column=inv_col).fill = INVHDR
    isin_col = inv_col + 1
    if tracker_isin:
        ws.cell(row=1, column=isin_col).fill = INVHDR

    r = 2
    for f in fcns:
        ko, st, ki = dec(f["ko_level"]), dec(f["strike_level"]), dec(f.get("ki_level"))
        ver = f.get("verification", {}) or {}
        obs = [date_formula(d, "obs_date") for d in place_obs(f.get("obs_dates", []), n_obs)]
        tail = [date_formula(f["maturity_date"], "maturity_date"), f["ko_type"],
                f.get("ki_type", ""), f["memory"], f["tenor_mo"],
                num(f["principal"]), num(dec(f["coupon_pa"]))]
        # status per paint-column (same for every underlying row of this note)
        extracted_by_col = {"Trade date": f["trade_date"], "Issue date": f["issue_date"],
            "Maturity Date": f["maturity_date"], "KO Type": f["ko_type"], "KI Type": f.get("ki_type", ""),
            "Memory": f["memory"], "Tenor (mo)": f["tenor_mo"], "Principal (ccy)": f["principal"],
            "Coupon (p.a.)": f["coupon_pa"], "KO Level (%)": f["ko_level"],
            "Strike Level (%)": f["strike_level"], "KI Level (%)": f.get("ki_level", "")}
        status_by_col = {}
        for cn in PAINT_COLS:
            if cn in ("Trade Price (ccy)", "KO Price", "Strike Price", "KI Price"):
                status_by_col[cn] = EXTRACTED           # computed / extracted -> never email-checked
            else:
                status_by_col[cn] = field_status(extracted_by_col[cn], ver.get(VER_KEY[cn]))[0]

        investor = (f.get("investor") or "").strip()
        for u_idx, u in enumerate(f["underlyings"]):
            X = dec(u["trade_price"])
            AB = X * ko if (X is not None and ko is not None) else None
            AC = X * st if (X is not None and st is not None) else None
            AD = X * ki if (X is not None and ki is not None) else None
            row = [date_formula(f["trade_date"], "trade_date"),
                   date_formula(f["issue_date"], "issue_date"), *obs, *tail,
                   u["stock"], f["currency"], num(X), num(ko), num(st), num(ki),
                   num(AB), num(AC), num(AD)]
            for cidx, val in enumerate(row, start=1):
                ws.cell(row=r, column=cidx, value=val)
            for cn in PAINT_COLS:
                s = status_by_col[cn]
                if s is None or (AD is None and cn in ("KI Level (%)", "KI Price")):
                    continue
                cell = ws.cell(row=r, column=col[cn])
                cell.fill = FILL[s]; cell.font = FONT[s]
            # Investor sits on the FIRST underlying row of the note only - that is how the
            # tracker records it (one entry per note, not per basket line)
            if investor and u_idx == 0:
                ic = ws.cell(row=r, column=inv_col, value=investor)
                ic.fill = FILL[OK]; ic.font = FONT[OK]
            # ISIN sits one column right of Investor, same first-row-only rule. Like
            # Investor it is booking-email sourced, so it paints green.
            if tracker_isin and u_idx == 0 and f.get("isin"):
                sc = ws.cell(row=r, column=isin_col, value=f["isin"])
                sc.fill = FILL[OK]; sc.font = FONT[OK]
            r += 1
        r += 1  # blank separator between notes

    for i, name in enumerate(hdr, start=1):
        letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[letter].width = 13 if name.startswith(("Trade", "Issue", "Obs", "Maturity")) else 11
    # multi-client Investor strings ("NAME (200,000), NAME (200,000)") need real width
    inv_width = max([26] + [len(f.get("investor") or "") for f in fcns])
    ws.column_dimensions[openpyxl.utils.get_column_letter(inv_col)].width = min(inv_width, 90)
    if tracker_isin:
        ws.column_dimensions[openpyxl.utils.get_column_letter(isin_col)].width = 16
    _legend(ws, r + 1, 1)
    ws.cell(row=r + 1, column=inv_col,
            value=f"Investor sits on its real tracker letter ({inv_letter}). The "
                  f"{SHEET_COMPUTED_COLS} grey columns before it are the sheet's own formulas "
                  f"and are emitted BLANK - don't paste them over a live tracker.").font = BOLD
    return ws


def _legend(ws, r, c):
    ws.cell(row=r, column=c, value="Colour key (核对 = cross-checked vs the dealer email):").font = BOLD
    for label, status in [("cross-checked vs email / booking table / subject", OK),
                          ("termsheet only - NOT cross-checked (incl. computed prices)", EXTRACTED),
                          ("email disagrees with termsheet - investigate", FAIL)]:
        r += 1
        m = ws.cell(row=r, column=c, value=MARK[status]); m.fill = FILL[status]; m.font = FONT[status]
        m.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=c + 1, value=label)


# ---- Audit sheet -----------------------------------------------------------

AUDIT_HDR = ["Note", "ISIN", "Issuer", "Field", "Underlying", "Extracted (termsheet)",
             "Reference (email)", "Match", "Source"]

# note-level fields listed in the Audit, in order
NOTE_FIELDS = [("KO Type", "ko_type"), ("KO Level (%)", "ko_level"),
    ("Strike Level (%)", "strike_level"), ("KI Level (%)", "ki_level"),
    ("Coupon (p.a.)", "coupon_pa"), ("Principal (ccy)", "principal"), ("Tenor (mo)", "tenor_mo"),
    ("Memory", "memory"), ("KI Type", "ki_type"), ("Trade date", "trade_date"),
    ("Issue date", "issue_date"), ("Maturity Date", "maturity_date")]


def build_audit(wb, fcns):
    ws = wb.create_sheet("Audit")
    ws.append(AUDIT_HDR)
    for c in ws[1]:
        c.fill = HDR; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "A2"
    r = 2

    for idx, f in enumerate(fcns, start=1):
        note = f"#{idx}"
        isin = f.get("isin", "")
        issuer = f.get("issuer", "")
        ver = f.get("verification", {}) or {}
        block_start = r

        def add(field, underlying, extracted, ref, status, source):
            nonlocal r
            ws.append([note, isin, issuer, field, underlying,
                       "" if extracted is None else str(extracted),
                       "" if ref is None else str(ref), MARK[status], source])
            mc = ws.cell(row=r, column=8)
            mc.fill = FILL[status]; mc.font = FONT[status]
            mc.alignment = Alignment(horizontal="center")
            r += 1

        # per-underlying: the initial/trade price (a termsheet figure, not in the email)
        for u in f["underlyings"]:
            add("Initial/Trade Price", u["stock"], u["trade_price"], None, EXTRACTED,
                "termsheet only (not in email)")

        # note-level terms
        vals = {"KO Type": f["ko_type"], "KO Level (%)": f["ko_level"],
                "Strike Level (%)": f["strike_level"], "KI Level (%)": f.get("ki_level", ""),
                "Coupon (p.a.)": f["coupon_pa"], "Principal (ccy)": f["principal"],
                "Tenor (mo)": f["tenor_mo"], "Memory": f["memory"], "KI Type": f.get("ki_type", ""),
                "Trade date": f["trade_date"], "Issue date": f["issue_date"],
                "Maturity Date": f["maturity_date"]}
        for label, key in NOTE_FIELDS:
            status, ref, source = field_status(vals[label], ver.get(key))
            add(label, "", vals[label], ref, status, source)

        # Investor never appears in the termsheet - it comes from the booking email only,
        # so it is green by construction (or amber-blank when the email didn't name a client)
        investor = (f.get("investor") or "").strip()
        if investor:
            src = (ver.get("investor", {}) or {}).get(
                "source", "booking table (TR Code + Client Name)")
            add("Investor", "", investor, investor, OK, src)
        else:
            add("Investor", "", "", None, EXTRACTED, "not stated in the booking email")

        for rr in range(block_start, r):
            ws.cell(row=rr, column=1).fill = NOTEBAR
        r += 1  # blank row between notes

    widths = [7, 15, 10, 20, 12, 20, 20, 7, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(AUDIT_HDR)):
        for cell in row:
            if cell.value not in (None, ""):
                cell.border = BORDER
    return ws


def clean_sender(name):
    """Sanitise a sender name for use in a filename (drop chars illegal on Windows,
    collapse whitespace). Returns '' if nothing usable is left."""
    if not name:
        return ""
    name = re.sub(r'[\\/:*?"<>|]', " ", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", help="path to the FCN JSON file")
    ap.add_argument("--outdir", default=os.path.join(os.path.expanduser("~"), "Downloads"))
    ap.add_argument("--sender", default="",
                    help="dealer/sender name to tag into the output filename so batches "
                         "from different senders stay distinguishable")
    ap.add_argument("--tracker-isin", action="store_true",
                    help="also write each note's ISIN one column right of Investor "
                         "(AR in the 12-obs layout, BD in the 24-obs). Shaun Lee Wei Qing's "
                         "layout only - see SKILL.md 'Shaun Lee Wei Qing: one block per note'")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8-sig") as fh:
        fcns = json.load(fh)["fcns"]

    n_obs = 24 if any(int(f["tenor_mo"]) > 12 for f in fcns) else 12

    wb = openpyxl.Workbook()
    build_tracker(wb, fcns, n_obs, tracker_isin=args.tracker_isin)
    build_audit(wb, fcns)

    now = datetime.now()
    sender = clean_sender(args.sender)
    tag = f" - {sender}" if sender else ""
    path = os.path.join(args.outdir,
                        "FCN Termsheet Reader{} {:%Y-%m-%d %H%M%S}.xlsx".format(tag, now))
    # per-sender folders (see SKILL.md "Naming & folders") usually don't exist yet
    os.makedirs(args.outdir, exist_ok=True)
    wb.save(path)

    n_rows = sum(len(f["underlyings"]) for f in fcns)
    print(f"WROTE: {path}")
    print(f"layout: {n_obs}-obs | FCNs: {len(fcns)} | data rows: {n_rows}")


if __name__ == "__main__":
    main()
