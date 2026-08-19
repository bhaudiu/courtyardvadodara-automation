#!/usr/bin/env python3
"""fetch_email.py — download EVERY complete daily report set from the mailbox.

Each email that contains all 3 reports (Manager Flash, Trial Balance, Daily
Operations) becomes one folder under inbox/, which build_site.py then parses.
This lets a single run backfill many days at once (e.g. a whole week forwarded
together), not just the newest one.

GitHub secrets: IMAP_HOST (default imap.gmail.com), IMAP_USER, IMAP_PASS (app password).
Optional: IMAP_LOOKBACK (recent emails to scan, default 150).
"""
import imaplib
import email
import os
import shutil

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]
LOOKBACK = int(os.environ.get("IMAP_LOOKBACK", "150"))
INBOX = "inbox"

# output filename -> (keywords matched in attachment name, allowed extensions)
TARGETS = {
    "manager_flash.pdf":     (("manager", "flash", "e106"), (".pdf",)),
    "trial_balance.pdf":     (("trial", "balance", "e100"), (".pdf",)),
    "daily_operations.xlsx": (("daily operation", "daily_operation", "operations", "operation", "simphony"), (".xlsx", ".xls")),
}


def run():
    if os.path.isdir(INBOX):
        shutil.rmtree(INBOX)
    os.makedirs(INBOX, exist_ok=True)

    M = imaplib.IMAP4_SSL(HOST)
    M.login(USER, PASS)
    M.select("INBOX")
    _, data = M.search(None, "ALL")
    ids = data[0].split()[-LOOKBACK:]
    sets = 0
    for n, i in enumerate(reversed(ids)):  # newest first
        _, md = M.fetch(i, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        found = {}
        for part in msg.walk():
            fn = part.get_filename()
            if not fn:
                continue
            low = fn.lower()
            for out, (keys, exts) in TARGETS.items():
                if out in found:
                    continue
                if low.endswith(exts) and any(k in low for k in keys):
                    found[out] = part.get_payload(decode=True)
        if len(found) == len(TARGETS):
            folder = os.path.join(INBOX, f"set_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for out, payload in found.items():
                with open(os.path.join(folder, out), "wb") as f:
                    f.write(payload)
            sets += 1
    M.logout()
    print(f"downloaded {sets} complete report set(s) into {INBOX}/")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error: {e}")
