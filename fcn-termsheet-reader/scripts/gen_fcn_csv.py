# -*- coding: utf-8 -*-
"""Generate the FCN tracker CSV from structured note data.

Usage:
    python gen_fcn_csv.py <data.json> [--outdir DIR]

<data.json> follows the schema in SKILL.md. The column spec and rules are in
references/columns.md. This script only does the deterministic work: compute the
KO/Strike/KI Price columns with exact decimals, auto-select the 12- vs 24-observation
layout from tenor, place observation dates, and write a timestamped CSV
("FCN Termsheet Reader <YYYY-MM-DD HHMMSS>.csv") with one row per underlying and a
blank row between notes.
"""
import csv, os, re, sys, json, argparse
from decimal import Decimal as D
from datetime import datetime

ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

FIELDS_AFTER_OBS = ["Maturity Date", "KO Type", "KI Type", "Memory", "Tenor (mo)",
    "Principal (ccy)", "Coupon (p.a.)", "Stock", "Currency", "Trade Price (ccy)",
    "KO Level (%)", "Strike Level (%)", "KI Level (%)", "KO Price", "Strike Price", "KI Price"]


def header_row(n_obs):
    obs = [f"Obs Date {i}" for i in range(1, n_obs)] + [f"Obs Date {n_obs}/Final Obs"]
    return ["Trade date", "Issue date"] + obs + FIELDS_AFTER_OBS


def fmt(x):
    if x is None or x == "":
        return ""
    if isinstance(x, D):
        return format(x.normalize(), "f")
    return str(x)


def dec(x):
    if x is None or x == "":
        return None
    return D(str(x))


def date_cell(d):
    """Emit a date as ="yyyy-mm-dd" so Excel keeps it as literal text.

    A bare 2026-06-04 gets parsed by Excel into a date and *displayed* per the
    system locale (e.g. 4/6/2026); copying that into Google Sheets then
    re-parses it under a different locale and flips day/month. The formula-text
    wrapper survives the Excel -> copy -> Google Sheets round trip unchanged.
    """
    if d is None or d == "":
        return ""
    if not ISO_DATE.match(str(d)):
        raise ValueError(f"date {d!r} is not yyyy-mm-dd - fix the JSON, never pass "
                         f"locale-ambiguous formats like 4/6/2026")
    return f'="{d}"'


def place_obs(obs_dates, n_slots):
    """Final observation -> last slot; earlier ones fill from the start; middle blank."""
    slots = [""] * n_slots
    if not obs_dates:
        return slots
    if len(obs_dates) > n_slots:
        raise ValueError(f"{len(obs_dates)} observation dates exceed {n_slots} slots "
                         f"- tenor too long for this layout")
    *earlier, final = obs_dates
    for i, d in enumerate(earlier):
        slots[i] = d
    slots[-1] = final
    return slots


def build_rows(fcns, n_obs):
    rows = []
    for f in fcns:
        ko, st, ki = dec(f["ko_level"]), dec(f["strike_level"]), dec(f.get("ki_level"))
        obs = [date_cell(d) for d in place_obs(f.get("obs_dates", []), n_obs)]
        note_tail = [date_cell(f["maturity_date"]), f["ko_type"], f.get("ki_type", ""),
                     f["memory"], f["tenor_mo"], f["principal"], dec(f["coupon_pa"])]
        for u in f["underlyings"]:
            X = dec(u["trade_price"])
            AB = X * ko if (X is not None and ko is not None) else None
            AC = X * st if (X is not None and st is not None) else None
            AD = X * ki if (X is not None and ki is not None) else None
            row = [date_cell(f["trade_date"]), date_cell(f["issue_date"]), *obs, *note_tail,
                   u["stock"], f["currency"], X, ko, st, ki, AB, AC, AD]
            rows.append([fmt(c) for c in row])
        rows.append(None)  # separator between FCNs
    if rows and rows[-1] is None:
        rows.pop()
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", help="path to the FCN JSON file")
    ap.add_argument("--outdir", default=os.path.join(os.path.expanduser("~"), "Downloads"))
    args = ap.parse_args()

    with open(args.data, encoding="utf-8-sig") as fh:  # utf-8-sig tolerates a BOM if present
        fcns = json.load(fh)["fcns"]

    # 24-obs layout if any note runs longer than 12 months, else 12-obs
    n_obs = 24 if any(int(f["tenor_mo"]) > 12 for f in fcns) else 12
    rows = build_rows(fcns, n_obs)

    now = datetime.now()
    path = os.path.join(args.outdir, "FCN Termsheet Reader {:%Y-%m-%d %H%M%S}.csv".format(now))
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(header_row(n_obs))
        for r in rows:
            if r is None:
                fh.write("\r\n")
            else:
                w.writerow(r)

    n_rows = sum(1 for r in rows if r is not None)
    print(f"WROTE: {path}")
    print(f"layout: {n_obs}-obs | FCNs: {len(fcns)} | data rows: {n_rows}")


if __name__ == "__main__":
    main()
