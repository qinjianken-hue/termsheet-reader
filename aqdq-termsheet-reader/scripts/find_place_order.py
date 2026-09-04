"""Find the dealer Place Order / trade-confirmation email that corroborates a Rajiv Sharma AQDQ trade.

Rajiv (TR 601) books his AQ/DQ trades by emailing a "Place Order: Accu/Decu on <TICKER>"
message to the structured-products desk on the TRADE DATE, containing a full order table
(Product, CCY, Tenor, # of days, Underlying, Strike %, Guarantee Period, KO %, Leverage
Factor, Issuer). The termsheet itself is only forwarded 1-2 business days LATER with a bare
"FYI" body -- so the order email is the independent second document needed for a real
cross-check. Occasionally the order is sent by Bhart Kishanchand Sheri instead.

The match often lands on a REPLY (e.g. Credit's "Pls proceed") that quotes Rajiv's original
table, which is why sender matching also looks inside the body.

Usage:
    python find_place_order.py --underlying MSFT --trade-date 2026-07-31 [--days-before 2]
                               [--days-after 3] [--chars 6000]

Prints every candidate's sender, timestamp, subject and body. Read the order table out of
the newest one that actually contains it.
"""
import argparse
import datetime
import sys

import win32com.client

SENDER_HINTS = ("rajiv", "bhart")
MAILBOX = "kenqinj@uobkh.com"


def iter_folder(folder, lo, hi):
    """Yield mail items received in [lo, hi], newest first, then recurse into subfolders."""
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
    except Exception:
        items = []
    for m in items:
        try:
            rt = m.ReceivedTime
            day = datetime.datetime(rt.year, rt.month, rt.day)
        except Exception:
            continue
        if day > hi:
            continue
        if day < lo:
            break
        yield m
    try:
        for f in folder.Folders:
            for m in iter_folder(f, lo, hi):
                yield m
    except Exception:
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--underlying", required=True,
                    help="ticker as it appears in the subject, e.g. MSFT, NFLX, AMD")
    ap.add_argument("--trade-date", required=True, help="ISO yyyy-mm-dd from the termsheet")
    ap.add_argument("--days-before", type=int, default=2)
    ap.add_argument("--days-after", type=int, default=3)
    ap.add_argument("--chars", type=int, default=6000, help="body characters to print")
    args = ap.parse_args()

    td = datetime.datetime.strptime(args.trade_date, "%Y-%m-%d")
    lo = td - datetime.timedelta(days=args.days_before)
    hi = td + datetime.timedelta(days=args.days_after)
    tick = args.underlying.upper()

    ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = ol.Folders[MAILBOX].Folders["Inbox"]

    hits = 0
    for m in iter_folder(inbox, lo, hi):
        try:
            subj = m.Subject or ""
            sender = m.SenderName or ""
        except Exception:
            continue
        if "place order" not in subj.lower():
            continue
        if tick not in subj.upper():
            continue
        try:
            body = m.Body or ""
        except Exception:
            body = ""
        who = (sender + " " + body[:4000]).lower()
        if not any(h in who for h in SENDER_HINTS):
            continue
        hits += 1
        print("=" * 78)
        print(f"{m.ReceivedTime} | from: {sender}")
        print(subj)
        print("=" * 78)
        print(body[:args.chars])
        print()

    if not hits:
        print(f"NO Place Order email found for {tick} in "
              f"{lo:%Y-%m-%d}..{hi:%Y-%m-%d} -- widen the window or check the ticker "
              f"spelling in the subject. Leave the terms amber (termsheet only).",
              file=sys.stderr)


if __name__ == "__main__":
    main()
