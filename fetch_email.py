#!/usr/bin/env python3

import imaplib
import email
import os
import shutil

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]

DAYS_BACK = int(os.environ.get("IMAP_DAYS_BACK", "30"))
INBOX_DIR = "inbox"

REPORTS_DIR = os.path.join(INBOX_DIR, "latest_dbr")

os.makedirs(INBOX_DIR, exist_ok=True)

def save_file(path, payload):
    with open(path, "wb") as f:
        f.write(payload)

def is_hnf(subject, filename):
    s = subject.lower()
    f = filename.lower()

    return (
        "history and forecast" in s
        or "r106" in s
        or "history_forecast" in f
        or "history and forecast" in f
    )

def is_trial_balance(subject, filename):
    s = subject.lower()
    f = filename.lower()

    return (
        "trial balance" in s
        or "trial_balance" in f
        or "d156" in s
