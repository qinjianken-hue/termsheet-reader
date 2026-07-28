# -*- coding: utf-8 -*-
"""Generate the AQDQ tracker CSV (columns A-O) from structured trade data.

Usage:
    python gen_aqdq_csv.py <data.json> [--outdir DIR]

<data.json> follows the schema in SKILL.md; field rules live in references/rules.md.
Dates in the JSON are ISO (yyyy-mm-dd); the CSV emits M/D/YYYY (no leading zeros) to
match the tracker. KO/Strike levels stay decimals (1.07 = 107%) so they always land in
Excel as calculable numbers. One row per trade, no blank separator rows.

Consistency checks (printed as WARNINGs, never block):
- Notional vs Total Expiries x Shares/Day x Trade Price (catches a wrong
  Shares/Day-vs-Leverage split or a wrong expiry count).
- KO/Strike direction vs AQ/DQ (AQ: KO >= 1 >= Strike; DQ reversed).
"""
import csv, os, json, argparse
from decimal import Decimal as D
from datetime import datetime

HEADERS = ["AQ/DQ", "Trade Date (M/D/YYYY)", "Initial Obs Date", "GTD Period End",
           "GTD Days", "Final Obs Date", "Total Expiries", "Notional", "Currency",
           "Leverage (1 or 2)", "Shares per Day", "Stock", "Trade Price (ccy)",
           "KO Level (%)", "Strike Level (%)"]


def mdY(iso):
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.month}/{d.day}/{d.year}"


def fmt(x):
    if isinstance(x, D):
        return format(x.normalize(), "f")
    return str(x)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", help="path to the AQDQ JSON file")
    ap.add_argument("--outdir", default=os.path.join(os.path.expanduser("~"), "Downloads"))
    args = ap.parse_args()

    with open(args.data, encoding="utf-8-sig") as fh:  # utf-8-sig tolerates a BOM
        trades = json.load(fh)["trades"]

    rows, warnings = [], []
    for t in trades:
        ko, st = D(str(t["ko_level"])), D(str(t["strike_level"]))
        px, notional = D(str(t["trade_price"])), D(str(t["notional"]))

        calc = D(int(t["total_expiries"])) * D(int(t["shares_per_day"])) * px
        if calc and abs(calc - notional) / calc > D("0.005"):
            warnings.append(f"{t['stock']}: notional {fmt(notional)} vs TotalExpiries x Shares/Day x "
                            f"TradePrice = {fmt(calc)} (>0.5% off) - re-check Shares/Day vs Leverage "
                            f"and Total Expiries")
        if t["aq_dq"] == "AQ" and not (ko >= 1 and st <= 1):
            warnings.append(f"{t['stock']}: AQ but KO={fmt(ko)}/Strike={fmt(st)} - direction looks off")
        if t["aq_dq"] == "DQ" and not (ko <= 1 and st >= 1):
            warnings.append(f"{t['stock']}: DQ but KO={fmt(ko)}/Strike={fmt(st)} - direction looks off")

        rows.append([t["aq_dq"], mdY(t["trade_date"]), mdY(t["initial_obs_date"]),
                     mdY(t["gtd_period_end"]), t["gtd_days"], mdY(t["final_obs_date"]),
                     t["total_expiries"], fmt(notional), t["currency"], t["leverage"],
                     t["shares_per_day"], t["stock"], fmt(px), fmt(ko), fmt(st)])

    path = os.path.join(args.outdir, "AQDQ Termsheet Reader {:%Y-%m-%d %H%M%S}.csv".format(datetime.now()))
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(HEADERS)
        w.writerows(rows)

    print(f"WROTE: {path}")
    print(f"trades: {len(rows)}")
    for msg in warnings:
        print("WARNING:", msg)


if __name__ == "__main__":
    main()
