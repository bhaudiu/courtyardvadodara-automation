#!/usr/bin/env python3
"""provisional_days.py — reconstruct a day when the Manager Flash is missing.

Since 2026-08-28 the property mails its reports as separate Opera / Simphony
schedules. The Trial Balance and the R106 History & Forecast arrive reliably;
the job named "Manager Flash" is emitting the T120 Rooms report instead, so no
day can be assembled the normal way.

fetch_email.py keeps whatever it did receive for such a day in inbox/part_*/ .
This module turns that into a usable record:

    rooms, occupancy, room revenue, ADR   <- R106 History section (hf.json)
    F&B and other revenue, by outlet      <- Trial Balance GL buckets
    covers / pax / average check          <- unavailable (Daily Operations)

The record is flagged `provisional: True`. It is never written over a real
report set, and the next run replaces it as soon as a genuine Manager Flash
arrives for that date.

Caveats, deliberate:
  * R106 counts every occupied room; the Manager Flash nets out complimentary
    and house-use rooms. Occupied rooms, occupancy and ADR therefore read
    slightly high on a provisional day (2 rooms on 2026-08-26, where both
    sources exist). Room revenue matches the Manager Flash exactly.
  * MTD / YTD are reconstructed from daily history rather than taken from the
    Manager Flash MONTH / YEAR columns, because those columns do not exist here.
"""

import datetime as dt
import json
import os

TB = "trial_balance.pdf"

OUTLETS = [
    ("ce_812", "812 CE"),
    ("meal_plan_812ce", "812 Meal Plan"),
    ("ird", "IRD"),
    ("brubon", "Brubon Café"),
    ("the_deck", "The Deck (Starbucks)"),
    ("bgc", "Baroda Grill (BGC)"),
    ("banquet", "Banquet"),
    ("mbow", "MBOW"),
]

OODS = [
    ("spa", "Spa"),
    ("transportation", "Transportation"),
    ("guest_laundry", "Guest Laundry"),
    ("wine_shop", "Wine Shop"),
    ("other_misc", "Others Misc."),
]


def hf_history_row(hf_doc, iso):
    """The newest R106 snapshot that records `iso` as closed history."""
    snaps = hf_doc.get("snaps") or {}
    for snap_date in sorted(snaps, reverse=True):
        row = snaps[snap_date].get(iso)
        if row and row.get("sec") == "history" and row.get("r") and row.get("rev"):
            return row
    return None


def resolve_business_date(key, tb_room, hf_doc):
    """Work out which business date a Trial Balance covers.

    inbox/part_<key> is keyed by the day the mail arrived, and the Trial Balance
    covers an earlier business date. Rather than assume the offset, match the
    Trial Balance's room-revenue GL sum against R106 History room revenue — the
    two agree to the rupee on days where both sources exist.
    """
    try:
        base = dt.date.fromisoformat(key)
    except ValueError:
        return None, None

    best = None
    for off in (1, 2, 0, 3):
        iso = (base - dt.timedelta(days=off)).isoformat()
        row = hf_history_row(hf_doc, iso)
        if not row:
            continue
        gap = abs(tb_room - row["rev"])
        if best is None or gap < best[0]:
            best = (gap, iso, row)

    if best is None:
        print(f"  provisional: {key}: no R106 history near this date; skipping")
        return None, None

    gap, iso, row = best
    tol = max(5000.0, 0.02 * row["rev"])
    if gap > tol:
        print(f"  provisional: {key}: Trial Balance room revenue {tb_room:,.0f} "
              f"matches no R106 day (closest {iso}, off by {gap:,.0f}); skipping")
        return None, None
    return iso, row


def record(row, buckets):
    """Build a store-schema day from an R106 history row + Trial Balance buckets."""
    rooms = row["r"]
    room_rev = row["rev"]
    occ = row.get("occ")
    adr = row.get("adr")
    total_rooms = round(rooms / occ) if (occ and rooms) else None

    # Room revenue is authoritative from R106 (it equals the Manager Flash
    # figure). Push the small GL residual into Others so the total still adds up.
    resid = buckets.get("room_revenue", 0.0) - room_rev

    other = {k: buckets.get(k, 0.0) for k, _ in OODS}
    other["other_misc"] = buckets.get("other_misc", 0.0) + resid
    other_total = sum(other.values())

    fb = {k: buckets.get(k, 0.0) for k, _ in OUTLETS}
    fb_total = sum(fb.values())

    return {
        "operating_income": {
            "total_rooms": total_rooms,
            "occupied": rooms,
            "occupancy": occ,
            "adr": adr,
            "revpar": (room_rev / total_rooms) if total_rooms else None,
            "total_covers": None,
            "apc": None,
            "room_revenue": room_rev,
            "fb_revenue": fb_total,
            "ood_revenue": other_total,
            "total_hotel_revenue": room_rev + fb_total + other_total,
            "comp_rooms": None,
            "house_use": None,
        },
        "fb_details": {
            "total_rnl": fb_total - fb["banquet"] - fb["ird"],
            "banquet_other": fb["banquet"],
            "ird": fb["ird"],
        },
        "outlet_details": [
            {"name": n, "covers": None, "avg_check": None, "revenue": fb.get(k, 0.0)}
            for k, n in OUTLETS
        ],
        "ood": [{"name": n, "revenue": other.get(k, 0.0)}
                for k, n in OODS if other.get(k, 0.0)],
        "mf_periods": None,
        "provisional": True,
    }


def fold(days, hf_path, inbox="inbox"):
    """Add a provisional record for every inbox/part_* that carries a Trial
    Balance. Returns the ISO dates added. Never overwrites a real day."""
    if not os.path.isdir(inbox):
        return []
    try:
        import parse_reports as pr
    except Exception as e:
        print(f"  provisional: parse_reports unavailable ({e}); skipping")
        return []

    try:
        hf_doc = json.load(open(hf_path)) if os.path.exists(hf_path) else {}
    except Exception as e:
        print(f"  provisional: could not read {hf_path} ({e}); skipping")
        return []

    added = []
    for d in sorted(os.listdir(inbox)):
        if not d.startswith("part_"):
            continue
        tb_path = os.path.join(inbox, d, TB)
        if not os.path.exists(tb_path):
            continue
        try:
            tb = pr.parse_trial_balance(tb_path)
        except Exception as e:
            print(f"  provisional parse error in {d}: {e}")
            continue

        buckets = tb.get("buckets") or {}
        iso, row = resolve_business_date(
            d[len("part_"):], buckets.get("room_revenue", 0.0), hf_doc)
        if not iso:
            continue
        if iso in days and not days[iso].get("provisional"):
            continue                      # a real report set already covers it

        rec = record(row, buckets)
        days[iso] = rec
        added.append(iso)
        oi = rec["operating_income"]
        print(f"  provisional {iso}: {oi['occupied']:.0f} rooms, "
              f"room {oi['room_revenue']:,.0f}, F&B {oi['fb_revenue']:,.0f}, "
              f"total {oi['total_hotel_revenue']:,.0f} (no covers)")
    return added

