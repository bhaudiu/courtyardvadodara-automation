#!/usr/bin/env python3
"""fetch_email.py — download every daily report set + H&F snapshots + STAR from the mailbox.

Three kinds of mail are harvested into inbox/ for build_site.py:

1. DBR set — an email carrying all 3 reports (Manager Flash, Trial Balance,
   Daily Operations) becomes inbox/set_NNNN/ . One run can backfill many days.

2. H&F snapshot — an email carrying R106 "History and Forecast" PDFs (one per
   forward month) becomes inbox/hnf_NNNN/ . build_site keys each snapshot by the
   RUN date printed on the PDF, so numbering here is irrelevant.

3. STAR / STR — a Monthly (or Weekly) STAR "Glance" workbook is saved as
   inbox/star_monthly.xlsx / inbox/star_weekly.xlsx (newest kept).

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


def _is_hnf(low):
    return low.endswith(".pdf") and ("hnf" in low or "history and forecast" in low or "history & forecast" in low)


def _is_star(low):
    return low.endswith((".xlsx", ".xls")) and "star" in low


def run():
    if os.path.isdir(INBOX):
        shutil.rmtree(INBOX)
    os.makedirs(INBOX, exist_ok=True)

    M = imaplib.IMAP4_SSL(HOST)
    M.login(USER, PASS)
    M.select("INBOX")
    _, data = M.search(None, "ALL")
    ids = data[0].split()[-LOOKBACK:]
    sets = hnf_snaps = 0
    star_saved = {"monthly": False, "weekly": False}   # first hit = newest, kept
    for n, i in enumerate(reversed(ids)):  # newest first
        _, md = M.fetch(i, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        found = {}       # DBR set for this email
        hnf_parts = []   # (filename, payload) of every H&F pdf in this email
        for part in msg.walk():
            fn = part.get_filename()
            if not fn:
                continue
            low = fn.lower()
            # DBR reports
            for out, (keys, exts) in TARGETS.items():
                if out in found:
                    continue
                if low.endswith(exts) and any(k in low for k in keys):
                    found[out] = part.get_payload(decode=True)
            # H&F pdfs
            if _is_hnf(low):
                hnf_parts.append((fn, part.get_payload(decode=True)))
            # STAR workbook
            if _is_star(low):
                kind = "weekly" if "weekly" in low else "monthly"
                if not star_saved[kind]:
                    with open(os.path.join(INBOX, f"star_{kind}.xlsx"), "wb") as f:
                        f.write(part.get_payload(decode=True))
                    star_saved[kind] = True
                    print(f"  saved star_{kind}.xlsx  ({fn})")

        if len(found) == len(TARGETS):
            folder = os.path.join(INBOX, f"set_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for out, payload in found.items():
                with open(os.path.join(folder, out), "wb") as f:
                    f.write(payload)
            sets += 1
        if hnf_parts:
            folder = os.path.join(INBOX, f"hnf_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for fn, payload in hnf_parts:
                with open(os.path.join(folder, os.path.basename(fn)), "wb") as f:
                    f.write(payload)
            hnf_snaps += 1
    M.logout()
    print(f"downloaded {sets} DBR set(s), {hnf_snaps} H&F snapshot email(s), "
          f"STAR monthly={star_saved['monthly']} weekly={star_saved['weekly']} into {INBOX}/")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error: {e}")
