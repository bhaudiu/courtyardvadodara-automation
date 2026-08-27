#!/usr/bin/env python3
"""
fetch_email.py — optimized version: download every daily report set + H&F snapshots + STAR.

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
from email.header import decode_header
from email.utils import parseaddr

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


def _is_hnf(low: str) -> bool:
    return low.endswith(".pdf") and ("hnf" in low or "history and forecast" in low or "history & forecast" in low)


def _is_star(low: str) -> bool:
    return low.endswith((".xlsx", ".xls")) and "star" in low


def _normalize_filename(fn: str) -> str:
    """
    Decode MIME-encoded filenames (e.g. =?UTF-8?B?...?=) and return a clean string.
    """
    if not fn:
        return ""
    decoded_parts = decode_header(fn)
    out = []
    for text, enc in decoded_parts:
        if isinstance(text, bytes):
            enc = enc or "utf-8"
            try:
                out.append(text.decode(enc, errors="replace"))
            except Exception:
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def _scan_message_for_attachments(msg: email.message.Message):
    """
    Efficiently scan a message for relevant attachments.

    Returns:
        found_dbr: dict[out_filename] -> payload bytes (only if all 3 are found)
        hnf_parts: list[(original_filename, payload_bytes)]
        star_info: dict[kind] -> (original_filename, payload_bytes) for first hit per kind
    """
    found_dbr = {}
    hnf_parts = []
    star_info = {}  # "monthly" or "weekly"

    # Precompute keys for speed
    target_keys = {out: keys for out, (keys, _) in TARGETS.items()}
    target_exts = {out: exts for out, (_, exts) in TARGETS.items()}

    # We only care about multipart/* and application/* parts with filenames
    for part in msg.walk():
        # Skip non-attachments quickly
        disp = part.get("Content-Disposition", "")
        if not disp or ("attachment" not in disp.lower() and "inline" not in disp.lower()):
            # Some servers omit Content-Disposition; still check filename
            if not part.get_filename():
                continue

        fn_raw = part.get_filename()
        if not fn_raw:
            continue

        fn = _normalize_filename(fn_raw)
        if not fn:
            continue

        low = fn.lower()

        # DBR reports
        if len(found_dbr) < len(TARGETS):
            for out, keys in target_keys.items():
                if out in found_dbr:
                    continue
                exts = target_exts[out]
                if low.endswith(exts) and any(k in low for k in keys):
                    payload = part.get_payload(decode=True)
                    if payload:
                        found_dbr[out] = payload

        # H&F pdfs
        if _is_hnf(low):
            payload = part.get_payload(decode=True)
            if payload:
                hnf_parts.append((fn, payload))

        # STAR workbook
        if _is_star(low):
            kind = "weekly" if "weekly" in low else "monthly"
            if kind not in star_info:
                payload = part.get_payload(decode=True)
                if payload:
                    star_info[kind] = (fn, payload)

        # Early exit: if we already have all DBR + both STAR types, we can still
        # continue to collect all H&F (they may be many), but we can skip DBR checks.
        # We keep walking to ensure all H&F are captured.

    return found_dbr, hnf_parts, star_info


def run():
    if os.path.isdir(INBOX):
        shutil.rmtree(INBOX)
    os.makedirs(INBOX, exist_ok=True)

    M = imaplib.IMAP4_SSL(HOST)
    M.login(USER, PASS)
    M.select("INBOX")

    # Search ALL, then slice last LOOKBACK
    _, data = M.search(None, "ALL")
    ids = data[0].split()[-LOOKBACK:]

    sets = 0
    hnf_snaps = 0
    star_saved = {"monthly": False, "weekly": False}

    # Process newest first
    for n, i in enumerate(reversed(ids)):
        # Fetch only headers + structure first? For simplicity and speed,
        # we still fetch full RFC822 but avoid extra parsing overhead.
        _, md = M.fetch(i, "(RFC822)")
        raw = md[0][1]
        if not raw:
            continue

        msg = email.message_from_bytes(raw)

        found_dbr, hnf_parts, star_info = _scan_message_for_attachments(msg)

        # Save DBR set if complete
        if len(found_dbr) == len(TARGETS):
            folder = os.path.join(INBOX, f"set_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for out, payload in found_dbr.items():
                with open(os.path.join(folder, out), "wb") as f:
                    f.write(payload)
            sets += 1

        # Save H&F snapshot if any
        if hnf_parts:
            folder = os.path.join(INBOX, f"hnf_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for fn, payload in hnf_parts:
                with open(os.path.join(folder, os.path.basename(fn)), "wb") as f:
                    f.write(payload)
            hnf_snaps += 1

        # Save STAR (first hit per kind = newest because we iterate newest-first)
        for kind, (fn, payload) in star_info.items():
            if not star_saved[kind]:
                path = os.path.join(INBOX, f"star_{kind}.xlsx")
                with open(path, "wb") as f:
                    f.write(payload)
                star_saved[kind] = True
                print(f"  saved star_{kind}.xlsx  ({fn})")

        # Optional progress log every 20 messages
        if (n + 1) % 20 == 0:
            print(f"processed {n + 1}/{len(ids)} messages...")

    M.logout()
    print(
        f"downloaded {sets} DBR set(s), {hnf_snaps} H&F snapshot email(s), "
        f"STAR monthly={star_saved['monthly']} weekly={star_saved['weekly']} into {INBOX}/"
    )


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error: {e}")
