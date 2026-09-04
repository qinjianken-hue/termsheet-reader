# -*- coding: utf-8 -*-
"""Generate the AQDQ tracker WORKBOOK (.xlsx) from structured trade data.

Usage:
    python gen_aqdq_xlsx.py <data.json> [--outdir DIR]

Two sheets (parallel to the FCN skill's gen_fcn_xlsx.py):
  * "Tracker" - the paste-ready rows (columns A-O, one row per trade, no blank
                separator rows; dates as ="M/D/YYYY" formula-text so the
                Excel -> Google Sheets copy never gets reformatted). Economic-term
                cells are colour-coded by whether the term was CROSS-CHECKED AGAINST
                THE DEALER EMAIL.
  * "Audit"   - one row per term: the extracted value, the email reference it was
                checked against, whether they matched, and the SOURCE of the check.

WHAT "核对" (cross-check) MEANS HERE (the user's definition - do not widen it):
  The termsheet is always the SOURCE OF TRUTH; every value is extracted from it.
  "核对" means confirming an extracted term against an INDEPENDENT SECOND DOCUMENT -
  the dealer's place-order / booking email table, the email subject line, or the
  filename. Internal arithmetic (e.g. notional = expiries x shares x spot) is NOT a
  cross-check: the maths cannot be wrong, so it proves nothing about whether the term
  was read correctly. It still runs as a self-check and prints a WARNING on mismatch,
  but it never colours a cell green. A term is only green if a real email reference
  corroborates it.

Colour key (Tracker + Audit "Match"):
    GREEN  ✓  = cross-checked against the dealer email / order table / subject
    AMBER  —  = extracted from the termsheet only - NOT cross-checked
    RED    ✗  = the email says something different from the termsheet - investigate

Driven by optional JSON fields per trade:
     "trade_ref", "issuer", "client"  -> labels for the Audit sheet (never in Tracker)
     "verification": {                -> ONLY email-corroborated terms go here
         "strike_level": {"ref": "128.80%", "source": "dealer order table (Strike %)"},
         "ko_level":     {"ref": "90.00%",  "source": "dealer order table (KO %)"},
         "leverage":     {"ref": "2",       "source": "dealer order table"},
         ...any of: aq_dq trade_date initial_obs_date gtd_period_end gtd_days
                    final_obs_date total_expiries notional currency leverage
                    shares_per_day stock trade_price ko_level strike_level
     }
  A field WITH a verification entry is green (email source shown); WITHOUT one it is
  amber ("termsheet only"). A numeric ref that disagrees goes RED. Do NOT add a
  verification entry unless the term genuinely appears in the email/subject/order table.
"""
import os, re, json, argparse
from decimal import Decimal as D
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# the 15 tracker columns (A-O), in order
HEADERS = ["AQ/DQ", "Trade Date (M/D/YYYY)", "Initial Obs Date", "GTD Period End",
           "GTD Days", "Final Obs Date", "Total Expiries", "Notional", "Currency",
           "Leverage (1 or 2)", "Shares per Day", "Stock", "Trade Price (ccy)",
           "KO Level (%)", "Strike Level (%)"]

# columns that hold ISO dates in the JSON and are written as ="M/D/YYYY" formula-text
DATE_COLS = {"Trade Date (M/D/YYYY)", "Initial Obs Date", "GTD Period End", "Final Obs Date"}

# column name -> verification key. Every economic/identity column is paintable; a
# column with no verification entry stays amber ("termsheet only").
VER_KEY = {
    "AQ/DQ": "aq_dq", "Trade Date (M/D/YYYY)": "trade_date", "Initial Obs Date": "initial_obs_date",
    "GTD Period End": "gtd_period_end", "GTD Days": "gtd_days", "Final Obs Date": "final_obs_date",
    "Total Expiries": "total_expiries", "Notional": "notional", "Currency": "currency",
    "Leverage (1 or 2)": "leverage", "Shares per Day": "shares_per_day", "Stock": "stock",
    "Trade Price (ccy)": "trade_price", "KO Level (%)": "ko_level", "Strike Level (%)": "strike_level"}

# ---- styling (identical palette to the FCN workbook) -----------------------
GREEN = PatternFill("solid", fgColor="C6EFCE")
AMBER = PatternFill("solid", fgColor="FFEB9C")
RED   = PatternFill("solid", fgColor="FFC7CE")
HDR   = PatternFill("solid", fgColor="1F4E78")
TRADEBAR = PatternFill("solid", fgColor="DDEBF7")
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


def dec_ref(x):
    """Parse a numeric ref (strips %, commas, currency). Returns None for anything with
    letters - a date ('09-Jul-26') or a label ('DECU', '4W', '12M') is compared textually,
    never numerically."""
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


def mdy(iso):
    if not ISO_DATE.match(str(iso)):
        raise ValueError(f"date {iso!r} is not yyyy-mm-dd - fix the JSON, never pass "
                         f"locale-ambiguous formats like 4/6/2026")
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.month}/{d.day}/{d.year}"


def num(x):
    if x is None or x == "":
        return None
    if isinstance(x, D):
        return float(x)
    return x


def to_dec(x, field="value"):
    """Decimal from a possibly-messy numeric string: tolerates commas, whitespace and a
    currency prefix (e.g. 'USD 152,271.80' -> 152271.80). Raises a clear error instead of a
    bare decimal.InvalidOperation so a stray comma never crashes with a raw stack trace."""
    if x is None or str(x).strip() == "":
        raise ValueError(f"{field} is empty - every trade needs a numeric {field}")
    s = re.sub(r"[^0-9.\-]", "", str(x))
    try:
        return D(s)
    except Exception:
        raise ValueError(f"{field} {x!r} is not a number (got {s!r} after cleaning)")


VALID_VER_KEYS = set(VER_KEY.values())


def self_checks(t):
    """Internal consistency checks (NOT email cross-checks). Returns list of warnings."""
    warns = []
    stock = t.get("stock", "?")
    ko, st = to_dec(t["ko_level"], "ko_level"), to_dec(t["strike_level"], "strike_level")
    px, notional = to_dec(t["trade_price"], "trade_price"), to_dec(t["notional"], "notional")
    calc = D(int(t["total_expiries"])) * D(int(t["shares_per_day"])) * px
    if calc and abs(calc - notional) / calc > D("0.005"):
        warns.append(f"{stock}: notional {notional} vs TotalExpiries x Shares/Day x "
                     f"TradePrice = {calc} (>0.5% off) - re-check Shares/Day vs Leverage "
                     f"and Total Expiries")
    # levels must be decimals (0.90, 1.288), never percent-numbers (90, 128.80) - a units
    # slip like strike=128.80 would slide past the direction test below, so band-check first
    for lname, lvl in [("KO", ko), ("Strike", st)]:
        if not (D("0.2") <= lvl <= D("5")):
            warns.append(f"{stock}: {lname} level {lvl} out of range - expected a decimal like "
                         f"0.90 or 1.288, not a percent number; check units")
    if t["aq_dq"] == "AQ" and not (ko >= 1 and st <= 1):
        warns.append(f"{stock}: AQ but KO={ko}/Strike={st} - direction looks off")
    if t["aq_dq"] == "DQ" and not (ko <= 1 and st >= 1):
        warns.append(f"{stock}: DQ but KO={ko}/Strike={st} - direction looks off")
    ver = t.get("verification", {}) or {}
    for k in ver:
        if k not in VALID_VER_KEYS:
            warns.append(f"{stock}: verification key '{k}' is not a tracker field - it will be "
                         f"IGNORED (typo? valid keys: {', '.join(sorted(VALID_VER_KEYS))})")
    return warns


# ---- Tracker sheet ---------------------------------------------------------

def cell_value(t, name):
    """The raw value for tracker column `name` (dates -> ISO for now)."""
    key = VER_KEY[name]
    v = t.get(key, "")
    if name in DATE_COLS:
        return v  # ISO string; formatted at write time
    if name in ("Notional", "Trade Price (ccy)", "KO Level (%)", "Strike Level (%)"):
        return num(to_dec(v, key)) if v not in (None, "") else None
    return v


def build_tracker(wb, trades):
    ws = wb.active
    ws.title = "Tracker"
    ws.append(HEADERS)
    for c in ws[1]:
        c.fill = HDR; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    col = {name: HEADERS.index(name) + 1 for name in HEADERS}

    r = 2
    for t in trades:
        ver = t.get("verification", {}) or {}
        for name in HEADERS:
            c = ws.cell(row=r, column=col[name])
            if name in DATE_COLS:
                # a 0W-guarantee trade has no guaranteed period at all - leave the date blank
                # rather than inventing one (GTD Days is then 0)
                raw = t.get(VER_KEY[name], "")
                c.value = f'="{mdy(raw)}"' if raw not in (None, "") else None
            else:
                c.value = cell_value(t, name)
            status = field_status(t.get(VER_KEY[name], ""), ver.get(VER_KEY[name]))[0]
            c.fill = FILL[status]; c.font = FONT[status]
        r += 1

    for i, name in enumerate(HEADERS, start=1):
        letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[letter].width = 15 if (name in DATE_COLS or name == "AQ/DQ") else 12
    _legend(ws, r + 1, 1)
    return ws


def _legend(ws, r, c):
    ws.cell(row=r, column=c, value="Colour key (核对 = cross-checked vs the dealer email):").font = BOLD
    for label, status in [("cross-checked vs dealer order table / booking / subject", OK),
                          ("termsheet only - NOT cross-checked (incl. spot/notional)", EXTRACTED),
                          ("email disagrees with termsheet - investigate", FAIL)]:
        r += 1
        m = ws.cell(row=r, column=c, value=MARK[status]); m.fill = FILL[status]; m.font = FONT[status]
        m.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=c + 1, value=label)


# ---- Audit sheet -----------------------------------------------------------

AUDIT_HDR = ["Trade", "Trade Ref", "Stock", "Field", "Extracted (termsheet)",
             "Reference (email)", "Match", "Source"]

# fields listed in the Audit, in tracker order (label -> verification key)
AUDIT_FIELDS = [("AQ/DQ", "aq_dq"), ("Trade Date", "trade_date"),
    ("Initial Obs Date", "initial_obs_date"), ("GTD Period End", "gtd_period_end"),
    ("GTD Days", "gtd_days"), ("Final Obs Date", "final_obs_date"),
    ("Total Expiries", "total_expiries"), ("Notional", "notional"), ("Currency", "currency"),
    ("Leverage", "leverage"), ("Shares per Day", "shares_per_day"), ("Stock", "stock"),
    ("Trade Price", "trade_price"), ("KO Level (%)", "ko_level"), ("Strike Level (%)", "strike_level")]

DATE_KEYS = {"trade_date", "initial_obs_date", "gtd_period_end", "final_obs_date"}


def build_audit(wb, trades):
    ws = wb.create_sheet("Audit")
    ws.append(AUDIT_HDR)
    for c in ws[1]:
        c.fill = HDR; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "A2"
    r = 2
    for idx, t in enumerate(trades, start=1):
        trade = f"#{idx}"
        ref_lbl = t.get("trade_ref", "")
        stock = t.get("stock", "")
        ver = t.get("verification", {}) or {}
        block_start = r
        for label, key in AUDIT_FIELDS:
            raw = t.get(key, "")
            shown = mdy(raw) if (key in DATE_KEYS and raw) else raw
            status, ref, source = field_status(raw, ver.get(key))
            ws.append([trade, ref_lbl, stock, label,
                       "" if shown in (None, "") else str(shown),
                       "" if ref is None else str(ref), MARK[status], source])
            mc = ws.cell(row=r, column=7)
            mc.fill = FILL[status]; mc.font = FONT[status]
            mc.alignment = Alignment(horizontal="center")
            r += 1
        for rr in range(block_start, r):
            ws.cell(row=rr, column=1).fill = TRADEBAR
        r += 1  # blank row between trades

    widths = [7, 20, 8, 18, 20, 18, 7, 34]
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
    ap.add_argument("data", help="path to the AQDQ JSON file")
    ap.add_argument("--outdir", default=os.path.join(os.path.expanduser("~"), "Downloads"))
    ap.add_argument("--sender", default="",
                    help="dealer/sender name to tag into the output filename so batches "
                         "from different senders stay distinguishable")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8-sig") as fh:
        trades = json.load(fh)["trades"]

    wb = openpyxl.Workbook()
    build_tracker(wb, trades)
    build_audit(wb, trades)

    now = datetime.now()
    sender = clean_sender(args.sender)
    tag = f" - {sender}" if sender else ""
    path = os.path.join(args.outdir,
                        "AQDQ Termsheet Reader{} {:%Y-%m-%d %H%M%S}.xlsx".format(tag, now))
    wb.save(path)

    print(f"WROTE: {path}")
    print(f"trades: {len(trades)}")
    for t in trades:
        for msg in self_checks(t):
            print("WARNING:", msg)


if __name__ == "__main__":
    main()
