#!/usr/bin/env python3
"""fetch_email.py — pull the 3 latest Courtyard Vadodara daily reports from an
IMAP mailbox (the Gmail the property team forwards reports to) and save them
where build_site.py expects. Non-fatal: keeps existing files if anything fails.

GitHub secrets: IMAP_HOST (default imap.gmail.com), IMAP_USER, IMAP_PASS (app password).
Optional: IMAP_LOOKBACK (recent emails to scan, default 60).
"""
import imaplib
import email
import os

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]
LOOKBACK = int(os.environ.get("IMAP_LOOKBACK", "60"))

# output filename -> (keywords matched in attachment name, allowed extensions)
TARGETS = {
    "manager_flash.pdf":     (("manager", "flash", "e106"), (".pdf",)),
    "trial_balance.pdf":     (("trial", "balance", "e100"), (".pdf",)),
    "daily_operations.xlsx": (("daily operation", "daily_operation", "operations", "operation", "simphony"), (".xlsx", ".xls")),
}


def run():
    M = imaplib.IMAP4_SSL(HOST)
    M.login(USER, PASS)
    M.select("INBOX")
    _, data = M.search(None, "ALL")
    ids = data[0].split()[-LOOKBACK:]
    got = {}
    for i in reversed(ids):  # newest first
        _, md = M.fetch(i, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        for part in msg.walk():
            fn = part.get_filename()
            if not fn:
                continue
            low = fn.lower()
            for out, (keys, exts) in TARGETS.items():
                if out in got:
                    continue
                if low.endswith(exts) and any(k in low for k in keys):
                    with open(out, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    got[out] = fn
                    print(f"  saved {out} <- {fn}")
        if len(got) == len(TARGETS):
            break
    M.logout()
    missing = [t for t in TARGETS if t not in got]
    if missing:
        print(f"  waiting on: {missing} (not in last {LOOKBACK} emails yet)")
    print(f"fetched {len(got)}/{len(TARGETS)} reports")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error (keeping existing files): {e}")
