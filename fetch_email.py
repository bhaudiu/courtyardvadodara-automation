#!/usr/bin/env python3
"""
fetch_email.py — download every daily report set + H&F snapshots + STAR.

Three kinds of mail are harvested into inbox/ for build_site.py:

1. DBR set — Manager Flash + Trial Balance + Daily Operations become
   inbox/set_<key>/ . The three reports may arrive TOGETHER in one email
   (the old "Income Auditor" forward) or SPLIT across separate emails
   (the Opera / Simphony scheduled reports that started 2026-08-28).
   Attachments are bucketed by the date in the filename when there is one,
   otherwise by the date the email was received, so both styles work.

2. H&F snapshot — R106 "History and Forecast" PDFs become inbox/hnf_NNNN/ .
   build_site keys each snapshot by the RUN date printed on the PDF, so the
   numbering here is irrelevant.

3. STAR / STR — a Monthly (or Weekly) STAR "Glance" workbook is saved as
   inbox/star_monthly.xlsx / inbox/star_weekly.xlsx (newest kept).

Attachments inside a .zip are expanded and matched like any other attachment
(Simphony mails Daily Operations as DailyReportScheduling.zip).

GitHub secrets: IMAP_HOST (default imap.gmail.com), IMAP_USER, IMAP_PASS (app password).
Optional: IMAP_LOOKBACK (max messages to scan, default 150)
          IMAP_DAYS_BACK (how far back to search, default 45)
          IMAP_MAILBOX  (default: Gmail's All Mail, so archived mail is still seen)
"""

import imaplib
import email
import email.message
import datetime as dt
import io
import os
import re
import shutil
import zipfile
from email.header import decode_header
from email.utils import parsedate_to_datetime

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]
LOOKBACK = int(os.environ.get("IMAP_LOOKBACK", "150"))
DAYS_BACK = int(os.environ.get("IMAP_DAYS_BACK", "45"))
INBOX = "inbox"

# output filename -> (keywords matched in attachment name, allowed extensions)
TARGETS = {
    "manager_flash.pdf": (
        ("manager", "flash", "e106"),
        (".pdf",),
    ),
    "trial_balance.pdf": (
        ("trial", "balance", "e100", "d156"),
        (".pdf",),
    ),
    "daily_operations.xlsx": (
        ("daily operation", "daily_operation", "dailyoperation",
         "operations", "operation", "simphony", "dailyreport"),
        (".xlsx", ".xls", ".csv"),
    ),
}

# 26.08.2026 / 26-08-2026 / 26_08_2026 / 26 08 2026
DATE_IN_NAME = re.compile(r"(\d{1,2})[.\-_ ](\d{1,2})[.\-_ ](20\d{2})")


def _is_hnf(low: str) -> bool:
    if not low.endswith(".pdf"):
        return False
    return (
        "hnf" in low
        or "r106" in low
        or ("history" in low and "forecast" in low)
    )


def _is_star(low: str) -> bool:
    return low.endswith((".xlsx", ".xls")) and "star" in low

def _normalize_filename(fn: str) -> str:
    """Decode MIME-encoded filenames (e.g. =?UTF-8?B?...?=) into a clean string."""
    if not fn:
        return ""
    out = []
    for text, enc in decode_header(fn):
        if isinstance(text, bytes):
            try:
                out.append(text.decode(enc or "utf-8", errors="replace"))
            except Exception:
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def _date_from_name(fn: str):
    """Return an ISO date if the filename carries one, else None."""
    m = DATE_IN_NAME.search(fn)
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    try:
        return dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


def _message_date(msg) -> str:
    """ISO date the message was sent, or 'unknown'."""
    raw = msg.get("Date")
    if not raw:
        return "unknown"
    try:
        return parsedate_to_datetime(raw).date().isoformat()
    except Exception:
        return "unknown"


def _iter_attachments(msg):
    """Yield (filename, payload) for every attachment, expanding .zip members."""
    for part in msg.walk():
        fn_raw = part.get_filename()
        if not fn_raw:
            continue
        fn = _normalize_filename(fn_raw)
        if not fn:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue

        if fn.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(payload)) as z:
                    names = [n for n in z.namelist() if not n.endswith("/")]
                    print(f"    zip {fn} contains: {names}")
                    for name in names:
                        try:
                            yield os.path.basename(name), z.read(name)
                        except Exception as e:
                            print(f"    could not read {name} from {fn}: {e}")
            except Exception as e:
                print(f"    could not open zip {fn}: {e}")
            continue

        yield fn, payload


def _classify(fn: str):
    """Return the TARGETS output name this attachment satisfies, or None."""
    low = fn.lower()
    for out, (keys, exts) in TARGETS.items():
        if low.endswith(exts) and any(k in low for k in keys):
            return out
    return None


def _select_mailbox(M) -> str:
    for candidate in (os.environ.get("IMAP_MAILBOX"), '"[Gmail]/All Mail"',
                      '"[Google Mail]/All Mail"', "INBOX"):
        if not candidate:
            continue
        try:
            if M.select(candidate, readonly=True)[0] == "OK":
                return candidate
        except Exception:
            continue
    raise RuntimeError("could not select any mailbox")

def run():
    if os.path.isdir(INBOX):
        shutil.rmtree(INBOX)
    os.makedirs(INBOX, exist_ok=True)

    M = imaplib.IMAP4_SSL(HOST)
    M.login(USER, PASS)

    mailbox = _select_mailbox(M)
    print(f"scanning mailbox {mailbox}")

    # Bound the scan by date, not by mailbox size — All Mail can be huge.
    since = (dt.date.today() - dt.timedelta(days=DAYS_BACK)).strftime("%d-%b-%Y")
    typ, data = M.search(None, "SINCE", since)
    if typ != "OK" or not data or not data[0]:
        typ, data = M.search(None, "ALL")
    ids = data[0].split()[-LOOKBACK:]
    print(f"{len(ids)} message(s) to scan since {since}")

    # bucket key (business date) -> {output filename: payload}
    buckets = {}
    hnf_snaps = 0
    star_saved = {"monthly": False, "weekly": False}

    for n, i in enumerate(reversed(ids)):  # newest first
        _, md = M.fetch(i, "(RFC822)")
        if not md or not md[0]:
            continue
        raw = md[0][1]
        if not raw:
            continue

        msg = email.message_from_bytes(raw)
        msg_date = _message_date(msg)
        subject = _normalize_filename(msg.get("Subject", ""))[:70]

        hnf_parts = []
        interesting = []
        attachments = list(_iter_attachments(msg))

        # An old-style bundle names most files "... 26.08.2026.pdf" but leaves
        # one as "Daily Operations (2).xlsx". When a message talks about exactly
        # one date, undated attachments in it belong to that date too.
        named_dates = {d for d in (_date_from_name(fn) for fn, _ in attachments) if d}
        default_key = named_dates.pop() if len(named_dates) == 1 else msg_date

        for fn, payload in attachments:
            low = fn.lower()

            out = _classify(fn)
            if out:
                key = _date_from_name(fn) or default_key
                slot = buckets.setdefault(key, {})
                if out not in slot:          # newest wins
                    slot[out] = payload
                    interesting.append(f"{out} <- {fn} [{key}]")

            if _is_hnf(low):
                hnf_parts.append((fn, payload))

            if _is_star(low):
                kind = "weekly" if "weekly" in low else "monthly"
                if not star_saved[kind]:
                    with open(os.path.join(INBOX, f"star_{kind}.xlsx"), "wb") as f:
                        f.write(payload)
                    star_saved[kind] = True
                    interesting.append(f"star_{kind}.xlsx <- {fn}")

        if hnf_parts:
            folder = os.path.join(INBOX, f"hnf_{n:04d}")
            os.makedirs(folder, exist_ok=True)
            for fn, payload in hnf_parts:
                with open(os.path.join(folder, os.path.basename(fn)), "wb") as f:
                    f.write(payload)
            hnf_snaps += 1
            interesting.append(f"{len(hnf_parts)} H&F pdf(s) -> hnf_{n:04d}")

        if interesting:
            print(f"  [{msg_date}] {subject}")
            for line in interesting:
                print(f"      {line}")

        if (n + 1) % 25 == 0:
            print(f"processed {n + 1}/{len(ids)} messages...")

    M.logout()
    # Write out every complete set; report the incomplete ones so a change in
    # what the property sends is visible in the log instead of silent.
    sets = 0
    incomplete = []
    for key in sorted(buckets):
        slot = buckets[key]
        if len(slot) == len(TARGETS):
            folder = os.path.join(INBOX, f"set_{key}")
            os.makedirs(folder, exist_ok=True)
            for out, payload in slot.items():
                with open(os.path.join(folder, out), "wb") as f:
                    f.write(payload)
            sets += 1
        else:
            missing = sorted(set(TARGETS) - set(slot))
            incomplete.append(f"{key}: have {sorted(slot)}, missing {missing}")

    if incomplete:
        print("incomplete day(s) — not written:")
        for line in incomplete:
            print(f"  {line}")

    print(
        f"downloaded {sets} DBR set(s), {hnf_snaps} H&F snapshot email(s), "
        f"STAR monthly={star_saved['monthly']} weekly={star_saved['weekly']} into {INBOX}/"
    )


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error: {e}")
