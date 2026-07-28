# -*- coding: utf-8 -*-
"""Dump an Outlook .msg: print the body text + attachment list, and save PDF
attachments and the HTML body into a workdir.

Usage:
    python parse_msg.py <file.msg> [workdir]

Run with PYTHONIOENCODING=utf-8 so Unicode in the body prints cleanly.
"""
import extract_msg, os, re, sys


def main():
    if len(sys.argv) < 2:
        sys.exit("usage: python parse_msg.py <file.msg> [workdir]")
    path = sys.argv[1]
    workdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(path) or ".", "_fcn_ts")
    os.makedirs(workdir, exist_ok=True)

    m = extract_msg.Message(path)
    print("SUBJECT:", m.subject)
    print("FROM:", m.sender)
    print("DATE:", m.date)
    print("\n===== BODY =====")
    print(m.body)

    print("\n===== ATTACHMENTS =====")
    for i, a in enumerate(m.attachments):
        fn = a.longFilename or a.shortFilename or f"att{i}"
        data = a.data
        size = len(data) if isinstance(data, (bytes, bytearray)) else "?"
        print(f"[{i}] {fn!r} ({size} bytes)")
        if isinstance(fn, str) and fn.lower().endswith(".pdf") and isinstance(data, (bytes, bytearray)):
            safe = re.sub(r'[\\/:*?"<>|]', "_", fn)
            dest = os.path.join(workdir, safe)
            with open(dest, "wb") as f:
                f.write(data)
            print(f"     saved -> {dest}")

    try:
        h = m.htmlBody
        if isinstance(h, (bytes, bytearray)):
            h = h.decode("utf-8", "ignore")
        if h:
            dest = os.path.join(workdir, "_body.html")
            with open(dest, "w", encoding="utf-8") as f:
                f.write(h)
            print(f"\nHTML body -> {dest} ({len(h)} chars)")
    except Exception as e:
        print("HTML body error:", e)

    print(f"\nWORKDIR: {workdir}")


if __name__ == "__main__":
    main()
