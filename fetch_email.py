#!/usr/bin/env python3

import imaplib
import email
import os
import shutil

HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
USER = os.environ["IMAP_USER"]
PASS = os.environ["IMAP_PASS"]

INBOX_DIR = "inbox"
DAYS_BACK = int(os.environ.get("IMAP_DAYS_BACK", "60"))

TARGETS = {
    "manager_flash.pdf": (
        ("manager", "flash", "e106"),
        (".pdf",)
    ),
    "trial_balance.pdf": (
        ("trial", "balance", "e100"),
        (".pdf",)
    ),
    "daily_operations.xlsx": (
        ("daily operation", "daily_operation", "operations", "operation", "simphony"),
        (".xlsx", ".xls")
    ),
}


def is_hnf(filename):
    filename = filename.lower()
    return (
        filename.endswith(".pdf")
        and (
            "hnf" in filename
            or "history and forecast" in filename
            or "history & forecast" in filename
        )
    )


def is_star(filename):
    filename = filename.lower()
    return filename.endswith((".xlsx", ".xls")) and "star" in filename


def search_relevant_messages(mail):
    searches = [
        f'newer_than:{DAYS_BACK}d has:attachment'
    ]

    msg_ids = set()

    for query in searches:
        status, data = mail.search(
            None,
            "X-GM-RAW",
            f'"{query}"'
        )

        if status == "OK" and data and data[0]:
            msg_ids.update(data[0].split())

    return sorted(msg_ids, key=int   if os.path.isdir(INBOX_DIR):
        shutil.rmtree(INBOX_DIR)

    os.makedirs(INBOX_DIR, exist_ok=True)

    mail = imaplib.IMAP4_SSL(HOST)
    mail.login(USER, PASS)
    mail.select("INBOX")

    ids = search_relevant_messages(mail)

    print(f"Found {len(ids)} candidate emails")

    sets = 0
    hnf_snaps = 0

    star_saved = {
        "monthly": False,
        "weekly": False,
    }

    for idx, msg_id in enumerate(reversed(ids)):

        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue

        try:
            msg = email.message_from_bytes(msg_data[0][1])
        except Exception:
            continue

        found = {}
        hnf_parts = []

        for part in msg.walk():

            if part.get_content_disposition() != "attachment":
                continue

            filename = part.get_filename()

            if not filename:
                continue

            low = filename.lower()

            # DBR Reports
            for output_name, (keywords, exts) in TARGETS.items():

                if output_name in found:
                    continue

                if low.endswith(exts) and any(k in low for k in keywords):
                    found[output_name] = part.get_payload(decode=True)

            # H&F Reports
            if is_hnf(low):
                hnf_parts.append(
                    (filename, part.get_payload(decode=True))
                )

            # STAR Reports
            if is_star(low):

                kind = (
                    "weekly"
                    if "weekly" in low
                    else "monthly"
                )

                if not star_saved[kind]:

                    with  os.path.join(
                            INBOX_DIR,
                            f"star_{kind}.xlsx"
                        ),
                        "wb",
                    ) as f:
                        f.write(part.get_payload(decode=True))

                    star_saved[kind] = True

                    print(
                        f"Saved star_{kind}.xlsx ({filename})"
                    )

        # Save DBR Set
        if len(found) == len(TARGETS):

            folder = os.path.join(
                INBOX_DIR,
                f"set_{idx:04d}"
            )

            os.makedirs(folder, exist_ok=True)

            for name, payload in found.items():
                with open(
                    os.path.join(folder, name),
                    "wb"
                ) as f:
                    f.write(payload)

            sets += 1

        # Save H&F Snapshot
        if hnf_parts:

            folder = os.path.join(
                INBOX_DIR,
                f"hnf_{idx:04d}"
            )

            os.makedirs(folder, exist_ok=True)

            for filename, payload in hnf_parts:

                with open(
                    os.path.join(
                        folder,
                        os.path.basename(filename)
                    ),
                    "wb"
                ) as f:
                    f.write(payload)

            hnf_snaps += 1

    mail.logout()

    print(
        f"\nDownloaded {sets} DBR set(s), "
        f"{hnf_snaps} H&F snapshot email(s), "
        f"STAR monthly={star_saved['monthly']} "
        f"weekly={star_saved['weekly']}"
    )


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        print(f"fetch_email error: {e}")
