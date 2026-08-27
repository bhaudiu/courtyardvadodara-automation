#!/usr/bin/env python3
"""
fetch_email.py

Downloads:
1. DBR report sets
2. History & Forecast snapshots
3. STAR reports
"""

import imaplib
import email
import os
import shutil

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]

LOOKBACK_DAYS = int(os.environ.get("IMAP_DAYS_BACK", "30"))
INBOX = "inbox"

TARGETS = {
    "manager_flash.pdf": (
        ("manager", "flash", "e106", "hkroomstatus"),
        (".pdf",),
    ),
    "trial_balance.pdf": (
        ("trial", "balance", "trial_balance", "d156"),
        (".pdf",),
    ),
    "daily_operations.xlsx": (
        (
            "daily operation",
            "daily_operation",
            "operations",
            "operation",
            "simphony",
            "grr",
            "audit",
        ),
        (".xlsx", ".xls"),
    ),
}


def is_hnf(subject, filename):
    s = subject.lower()
    f = filename.lower()

    return (
        "history and forecast" in s
        or "r106" in s
        or "history_forecast" in f
        or "history and forecast" in f
        or "history_forecast" in s
    )


def is_star(subject, filename):
    s = subject.lower()
    f = filename.lower()

    return (
        "star" in s
        or "star" in f
        or "glance" in f
    )


def search_messages(mail):
    query = f"newer_than:{LOOKBACK_DAYS}d has:attachment"

    status, data = mail.search(
        None,
        "X-GM-RAW",
        f'"{query}"'
    )

    if status != "OK":
        print("Gmail search failed, falling back to ALL")
        status, data = mail.search(None, "ALL")

        if status != "OK":
            return []

    if not data or not data[0]:
        return []

    return data
def run():

    if os.path.isdir(
