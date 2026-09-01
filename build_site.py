#!/usr/bin/env python3
"""
build_site.py — Courtyard Vadodara DBR site builder.

Runs in GitHub Actions after fetch_email.py. Self-contained (no Excel recalc):
  1. Load persistent store.json (all daily records + the budget/YTD anchor).
  2. If the 3 report files are present, parse them into today's record and add it.
  3. Recompute the full Today/MTD/YTD owner report for every date + insights.
  4. Write data.json (consumed by dbr.html) and save store.json back.
"""
import os
import sys
import json
import datetime as dt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_reports import build_grr_day  # noqa

STORE = "store.json"
DATA = "data.json"
HF = "hf.json"
STAR = "star.json"
# filenames that fetch_email.py writes
MF, TB, DO = "manager_flash.pdf", "trial_balance.pdf", "daily_operations.xlsx"


# ---------------- normalize a freshly parsed day into the store schema ----------------
def normalize(day):
    oi = day["operating_income"]
    return {
        "operating_income": {
            "total_rooms": oi["total_rooms"], "occupied": oi["occupied_rooms"],
            "occupancy": oi["occupancy_pct"], "adr": oi["adr"], "revpar": oi["revpar"],
            "total_covers": oi["total_covers"], "apc": oi["apc"],
            "room_revenue": oi["room_revenue"], "fb_revenue": oi["fb_revenue"],
            "ood_revenue": oi["ood_revenue"], "total_hotel_revenue": oi["total_hotel_revenue"],
            "comp_rooms": oi.get("comp_rooms"), "house_use": oi.get("house_use"),
        },
        "fb_details": day["fb_details"],
        "outlet_details": [{"name": o["name"], "covers": o["covers"],
                            "avg_check": o["avg_check"], "revenue": o["revenue"]}
                           for o in day["outlet_details"]],
        "ood": [{"name": o["name"], "revenue": o["revenue"]} for o in day["ood"]],
        "mf_periods": day.get("mf_periods"),
    }


# ---------------- report reconstruction (ported, budget/history based) ----------------
def _outlet_key(l):
    l = l.lower()
    if "meal plan" in l: return "812 Meal Plan"
    if "812 ce" in l: return "812 CE"
    if "brubon" in l: return "Brubon Café"
    if "mbow" in l: return "MBOW"
    if "bgc" in l or "baroda" in l: return "Baroda Grill (BGC)"
    if "the deck" in l or "starbucks" in l: return "The Deck (Starbucks)"
    if "bqts" in l or ("banquet" in l and "rev" in l): return "Banquet"
    if l.startswith("ird"): return "IRD"
    return None


def _ood_key(l):
    l = l.lower()
    if l.startswith("spa"): return "Spa"
    if "transport" in l: return "Transportation"
    if "laundry" in l: return "Guest Laundry"
    if "wine" in l: return "Wine Shop"
    if "banquet other" in l: return "Banquet Other (Rental/AV)"
    if "others misc" in l: return "Others Misc."
    if "event management" in l: return "Event Management Fee"
    return None


def _spec(label):
    l = label.strip()
    if l == "Occupancy (%)": return ("rate", "occupied", "available")
    if l == "ADR": return ("rate", "room_rev", "occupied")
    if l == "RevPar": return ("rate", "room_rev", "available")
    if l == "Total APC": return ("rate", "fb_rev", "covers")
    if "Avg. Check" in l or "Avg. Cvr" in l or "Avg Cvr" in l:
        k = _outlet_key(l); return ("rate", "orev:" + k, "ocov:" + k) if k else None
    simple = {"Total Rooms in Hotel": "available", "Total Rooms in Hotel minus OOO Rooms": "available",
              "Occupied Rooms": "occupied", "Total Covers": "covers", "Room Revenue": "room_rev",
              "F&B Revenue": "fb_rev", "OOD Revenue": "ood_rev", "Total Hotel Revenue": "total_rev",
              "Total RNL": "rnl", "Banquet & Other Than F&B": "banquet_other",
              "Comp.": "comp_rooms", "HSG": "house_use"}   # room-night counts (Value, not money)
    if l in simple: return ("sum", simple[l])
    if l == "IRD": return ("sum", "ird_fb")
    if "No. of Cvrs" in l or "No.of Cvrs" in l or l.endswith("Cvrs"):
        k = _outlet_key(l); return ("sum", "ocov:" + k) if k else None
    if l.endswith("Rev.") or "Rev." in l:
        k = _outlet_key(l); return ("sum", "orev:" + k) if k else None
    k = _ood_key(l)
    if k: return ("sum", "ood:" + k)
    return None


def _base(d):
    oi, fb = d["operating_income"], d["fb_details"]
    b = {"available": oi.get("total_rooms") or 0, "occupied": oi.get("occupied") or 0,
         "room_rev": oi.get("room_revenue") or 0, "fb_rev": oi.get("fb_revenue") or 0,
         "ood_rev": oi.get("ood_revenue") or 0, "total_rev": oi.get("total_hotel_revenue") or 0,
         "covers": oi.get("total_covers") or 0, "rnl": fb.get("total_rnl") or 0,
         "banquet_other": fb.get("banquet_other") or 0, "ird_fb": fb.get("ird") or 0,
         "comp_rooms": oi.get("comp_rooms") or 0, "house_use": oi.get("house_use") or 0}
    for o in d["outlet_details"]:
        b["orev:" + o["name"]] = o.get("revenue") or 0
        b["ocov:" + o["name"]] = o.get("covers") or 0
    for o in d["ood"]:
        b["ood:" + o["name"]] = o.get("revenue") or 0
    return b


def build_report_for(iso, days, anchor_iso, anchor_sections):
    """Full Today/MTD/YTD x Actual/LY/Budget/Var report for a 2026 date."""
    y, m, D = int(iso[:4]), int(iso[5:7]), int(iso[8:10])
    ad = int(anchor_iso[8:10])
    base_ty = {int(k[8:10]): _base(days[k]) for k in days if k[:7] == f"{y:04d}-{m:02d}"}
    base_ly = {int(k[8:10]): _base(days[k]) for k in days if k[:7] == f"{y-1:04d}-{m:02d}"}
    comp2anchor = {}
    for s in anchor_sections:
        for r in s["rows"]:
            sp = _spec(r["label"])
            if sp and sp[0] == "sum":
                comp2anchor.setdefault(sp[1], r)

    def agg(comp):
        a = comp2anchor.get(comp)
        tyd = base_ty.get(D, {}).get(comp, 0)
        ty_mtd = sum(base_ty.get(d, {}).get(comp, 0) for d in range(1, D + 1))
        ly_mtd = sum(base_ly.get(d, {}).get(comp, 0) for d in range(1, D + 1))
        fwd = sum(base_ty.get(d, {}).get(comp, 0) for d in base_ty if ad < d <= D)
        bwd = sum(base_ty.get(d, {}).get(comp, 0) for d in base_ty if D < d <= ad)
        fwd_ly = sum(base_ly.get(d, {}).get(comp, 0) for d in base_ly if ad < d <= D)
        bwd_ly = sum(base_ly.get(d, {}).get(comp, 0) for d in base_ly if D < d <= ad)
        ty_ytd = (a["ytd_ty"] - bwd + fwd) if a and a.get("ytd_ty") is not None else None
        ly_ytd = (a["ytd_ly"] - bwd_ly + fwd_ly) if a and a.get("ytd_ly") is not None else None
        bt = a["today_bud"] if a else None
        return {"ty_today": tyd, "ty_mtd": ty_mtd, "ly_mtd": ly_mtd, "ty_ytd": ty_ytd, "ly_ytd": ly_ytd,
                "bud_today": bt, "bud_mtd": bt * D if bt is not None else None,
                "bud_ytd": (a["ytd_bud"] + bt * (D - ad)) if (a and a.get("ytd_bud") is not None and bt is not None) else None}

    def var(a, b): return (a - b) / b if (a is not None and b) else None
    def div(n, d): return (n / d) if (n is not None and d) else None

    out = []
    for s in anchor_sections:
        rows = []
        for r in s["rows"]:
            sp = _spec(r["label"])
            if sp is None:
                rows.append({"label": r["label"], "kind": r["kind"], "today_ty": None, "today_bud": None,
                             "today_var": None, "mtd_ty": None, "mtd_ly": None, "mtd_bud": None, "mtd_var": None,
                             "ytd_ty": None, "ytd_ly": None, "ytd_bud": None, "ytd_var": None})
                continue
            if sp[0] == "sum":
                g = agg(sp[1])
                row = {"today_ty": g["ty_today"], "today_bud": g["bud_today"], "mtd_ty": g["ty_mtd"],
                       "mtd_ly": g["ly_mtd"], "mtd_bud": g["bud_mtd"], "ytd_ty": g["ty_ytd"],
                       "ytd_ly": g["ly_ytd"], "ytd_bud": g["bud_ytd"]}
            else:
                n, d = agg(sp[1]), agg(sp[2])
                row = {"today_ty": div(n["ty_today"], d["ty_today"]), "today_bud": r["today_bud"],
                       "mtd_ty": div(n["ty_mtd"], d["ty_mtd"]), "mtd_ly": div(n["ly_mtd"], d["ly_mtd"]),
                       "mtd_bud": r["mtd_bud"], "ytd_ty": div(n["ty_ytd"], d["ty_ytd"]),
                       "ytd_ly": div(n["ly_ytd"], d["ly_ytd"]), "ytd_bud": r["ytd_bud"]}
            row["today_var"] = var(row["today_ty"], row["today_bud"])
            row["mtd_var"] = var(row["mtd_ty"], row["mtd_bud"])
            row["ytd_var"] = var(row["ytd_ty"], row["ytd_bud"])
            row["label"], row["kind"] = r["label"], r["kind"]
            rows.append(row)
        out.append({"title": s["title"], "rows": rows})

    # Override the summary revenue MTD/YTD with the AUTHORITATIVE Manager Flash
    # MONTH/YEAR columns for this day — the reconstructed daily sums can't see
    # missing days (e.g. 13 Aug) or un-re-parsed older days, so MTD/YTD drift.
    mfp = days.get(iso, {}).get("mf_periods")
    if mfp:
        LMAP = {"Room Revenue": "room", "F&B Revenue": "fb",
                "OOD Revenue": "ood", "Total Hotel Revenue": "total",
                "Occupied Rooms": "occupied"}
        for s in out:
            for r in s["rows"]:
                key = LMAP.get(r["label"].strip())
                p = mfp.get(key) if key else None
                if not p:
                    continue
                if p.get("day") is not None:
                    r["today_ty"] = p["day"]
                if p.get("mtd") is not None:
                    r["mtd_ty"] = p["mtd"]
                if p.get("ytd") is not None:
                    r["ytd_ty"] = p["ytd"]
                if p.get("ly_mtd") is not None:
                    r["mtd_ly"] = p["ly_mtd"]
                if p.get("ly_ytd") is not None:
                    r["ytd_ly"] = p["ly_ytd"]
                r["today_var"] = var(r["today_ty"], r["today_bud"])
                r["mtd_var"] = var(r["mtd_ty"], r["mtd_bud"])
                r["ytd_var"] = var(r["ytd_ty"], r["ytd_bud"])

        # Keep the rooms-statistics block internally consistent with the
        # Manager-Flash-overridden components. Occupied Rooms / Room Revenue are
        # now full-period totals, so the available-room denominator must be too —
        # a missing daily report leaves the daily-sum capacity short by a day.
        # Physical capacity (= budgeted rooms, OOO is 0 here) is the correct
        # full-period available count, so pin actual & LY available to budget,
        # then recompute Occupancy% / ADR / RevPAR from the corrected figures.
        idx = {}
        for s in out:
            for r in s["rows"]:
                idx[r["label"].strip()] = r
        for cap_lbl in ("Total Rooms in Hotel", "Total Rooms in Hotel minus OOO Rooms"):
            r = idx.get(cap_lbl)
            if r:
                if r.get("mtd_bud") is not None:
                    r["mtd_ty"] = r["mtd_ly"] = r["mtd_bud"]
                if r.get("ytd_bud") is not None:
                    r["ytd_ty"] = r["ytd_ly"] = r["ytd_bud"]
                r["mtd_var"] = var(r["mtd_ty"], r["mtd_bud"])
                r["ytd_var"] = var(r["ytd_ty"], r["ytd_bud"])
        occ = idx.get("Occupied Rooms")
        rrev = idx.get("Room Revenue")
        avail = idx.get("Total Rooms in Hotel minus OOO Rooms") or idx.get("Total Rooms in Hotel")

        def _derive(lbl, num_row, den_row):
            r = idx.get(lbl)
            if not (r and num_row and den_row):
                return
            for per in ("mtd", "ytd"):
                n_ty, d_ty = num_row.get(per + "_ty"), den_row.get(per + "_ty")
                if n_ty is not None and d_ty:
                    r[per + "_ty"] = n_ty / d_ty
                n_ly, d_ly = num_row.get(per + "_ly"), den_row.get(per + "_ly")
                if n_ly is not None and d_ly:
                    r[per + "_ly"] = n_ly / d_ly
                r[per + "_var"] = var(r.get(per + "_ty"), r.get(per + "_bud"))

        _derive("Occupancy (%)", occ, avail)
        _derive("ADR", rrev, occ)
        _derive("RevPar", rrev, avail)
    return out


# ---------------- insights (budget-based) ----------------
def _inr(v):
    n = int(round(v)); neg = n < 0; s = str(abs(n))
    if len(s) > 3:
        last3, rest, parts = s[-3:], s[:-3], []
        while len(rest) > 2:
            parts.insert(0, rest[-2:]); rest = rest[:-2]
        if rest: parts.insert(0, rest)
        s = ",".join(parts) + "," + last3
    return ("-" if neg else "") + "₹" + s


def build_insights(sections):
    def find(lbl):
        for s in sections:
            for r in s["rows"]:
                if r["label"].startswith(lbl): return r
    OOD = {"Spa", "Transportation", "Guest Laundry", "Wine Shop"}
    items = []
    for s in sections:
        for r in s["rows"]:
            lab = r["label"].strip()
            if r["kind"] != "currency": continue
            if not (lab.endswith(("Rev.", "Revenue")) or lab in OOD) or lab.startswith("Total Hotel Revenue"):
                continue
            if r.get("today_ty") is None or r.get("today_bud") is None: continue
            items.append({"label": lab, "today": r["today_ty"], "amt": r["today_ty"] - r["today_bud"],
                          "pct": r.get("today_var"), "mtd_ty": r.get("mtd_ty"), "mtd_ly": r.get("mtd_ly")})
    MAT = 8000
    wins = sorted([i for i in items if i["amt"] >= MAT], key=lambda x: -x["amt"])[:3]
    cons = sorted([i for i in items if i["amt"] <= -MAT], key=lambda x: x["amt"])[:3]
    ps = lambda p: ("+" if p >= 0 else "") + f"{p*100:.0f}%"
    clean = lambda l: l.replace("-Rev.", "").replace("Rev.", "").strip()
    tot = find("Total Hotel Revenue")
    hi = []
    if tot and tot.get("today_var") is not None:
        hi.append({"title": f"Total revenue {ps(tot['today_var'])} vs budget",
                   "detail": f"{_inr(tot['today_ty'])} today against a {_inr(tot['today_bud'])} budget."})
    hi += [{"title": f"{clean(w['label'])} beat budget by {_inr(w['amt'])}",
            "detail": f"{_inr(w['today'])} today ({ps(w['pct']) if w['pct'] is not None else ''} vs budget)."} for w in wins]
    watch = [{"title": f"{clean(c['label'])} {ps(c['pct']) if c['pct'] is not None else ''} vs budget",
              "detail": f"{_inr(c['today'])} today, {_inr(-c['amt'])} below the {_inr(c['budget']) if 'budget' in c else _inr(c['today']-c['amt'])} budget."} for c in cons]

    def action(l):
        l = l.lower()
        if "banquet" in l or "bqts" in l: return "Push the banquet & events pipeline — confirm tentative bookings."
        if "bgc" in l or "baroda" in l: return "Review Baroda Grill covers, footfall and pricing."
        if "deck" in l or "starbucks" in l: return "Lift The Deck footfall with a promotion."
        if "f&b" in l: return "Review F&B covers and average check to close the gap."
        return "Review pricing, covers and promotions to close the gap."
    acts = [{"title": f"{clean(c['label'])}: {_inr(-c['amt'])} behind budget", "detail": action(c["label"])} for c in cons]
    drivers = ", ".join(clean(w["label"]) for w in wins[:2])
    tv = tot.get("today_var") if tot else None
    tone = "Strong day" if (tv or 0) >= 0.05 else ("Soft day" if (tv or 0) <= -0.05 else "On-plan day")
    summary = (f"{tone} — total hotel revenue {_inr(tot['today_ty'])}, {ps(tv) if tv is not None else ''} vs budget"
               + (f", led by {drivers}." if drivers else ".")) if tot else ""
    return {"summary": summary, "highlights": hi[:4], "watch": watch, "action_items": acts}


# ---------------- H&F (Pick-up) accumulation from inbox/hnf_* ----------------
def update_hf():
    """Merge every inbox/hnf_* snapshot email into the persistent hf.json.

    hf.json accumulates across runs: each snapshot is keyed by the RUN date
    printed on the R106 PDF, so re-fetching an old email is idempotent and a
    new day simply adds a new snapshot. Never loses history already stored.
    """
    if os.path.exists(HF):
        doc = json.load(open(HF))
    else:
        doc = {"latest": None, "dates": [], "snaps": {}}
    snaps = doc.get("snaps", {})

    inbox = "inbox"
    added = []
    if os.path.isdir(inbox):
        try:
            import parse_hnf  # noqa
        except Exception as e:
            print(f"  hf: parse_hnf unavailable ({e}); skipping H&F update")
            return
        for d in sorted(os.listdir(inbox)):
            if not d.startswith("hnf_"):
                continue
            folder = os.path.join(inbox, d)
            try:
                run_date, data = parse_hnf.parse_hnf_folder(folder)
            except Exception as e:
                print(f"  hf parse error in {d}: {e}")
                continue
            if run_date and data:
                # Opera now mails one R106 PDF per forward month in its own
                # email, so several folders share a run date. Merge them into
                # one snapshot instead of letting the last folder replace it.
                snap = snaps.setdefault(run_date, {})
                snap.update(data)
                added.append(run_date)
                print(f"  hf: snapshot {run_date} +{len(data)} stay rows from {d} "
                      f"({len(snap)} total)")

    if not snaps:
        return
    dates = sorted(snaps.keys())
    doc = {"latest": dates[-1], "dates": dates, "snaps": snaps}
    json.dump(doc, open(HF, "w"), separators=(",", ":"))
    print(f"Wrote {HF}: {len(snaps)} snapshots, latest {doc['latest']}"
          + (f"; added {sorted(set(added))}" if added else "; no new snapshot"))


# ---------------- STR / STAR from inbox/star_*.xlsx ----------------
def update_star():
    """Refresh star.json from any STAR workbook fetch_email.py saved.

    Monthly and Weekly are updated independently; a missing file keeps whatever
    was stored before (the newest STAR report each carries a full period view).
    """
    monthly_x = os.path.join("inbox", "star_monthly.xlsx")
    weekly_x = os.path.join("inbox", "star_weekly.xlsx")
    if not (os.path.exists(monthly_x) or os.path.exists(weekly_x)):
        return
    if os.path.exists(STAR):
        doc = json.load(open(STAR))
    else:
        doc = {"monthly": None, "weekly": None}
    try:
        import parse_star  # noqa
    except Exception as e:
        print(f"  star: parse_star unavailable ({e}); skipping STAR update")
        return
    for key, path in (("monthly", monthly_x), ("weekly", weekly_x)):
        if os.path.exists(path):
            try:
                doc[key] = parse_star.parse_glance(path)
                print(f"  star: {key} -> {doc[key].get('period_label')} "
                      f"(created {doc[key].get('created')})")
            except Exception as e:
                print(f"  star parse error ({key}): {e}")
    json.dump(doc, open(STAR, "w"), separators=(",", ":"), default=str)
    print(f"Wrote {STAR}")


# ---------------- main ----------------
def main():
    store = json.load(open(STORE))
    days = store["days"]
    anchor_iso = store["anchor_date"]
    anchor_sections = store["anchor_sections"]

    # H&F first: a day with no Manager Flash is reconstructed from the R106
    # history rows in hf.json, so those have to be current before the DBR build.
    try:
        update_hf()
    except Exception as e:
        print(f"update_hf failed: {e}")

    # fold in last-year reference history (2024-2025 daily records from the monthly
    # GRR files) so the LY line + year-over-year columns have data. Never overwrites
    # this-year (2026) fetched days.
    if os.path.exists("history.json"):
        hist = json.load(open("history.json"))
        n = 0
        for k, v in hist.get("days", {}).items():
            if k < "2026-01-01":
                days[k] = v
                n += 1
        print(f"  merged {n} last-year history day(s) from history.json")

    # fold in EVERY complete report set that fetch_email.py downloaded (backfills many days)
    added = []
    inbox = "inbox"
    if os.path.isdir(inbox):
        for d in sorted(os.listdir(inbox)):
            p = os.path.join(inbox, d)
            mf, tb, do = os.path.join(p, MF), os.path.join(p, TB), os.path.join(p, DO)
            if not all(os.path.exists(x) for x in (mf, tb, do)):
                continue
            try:
                day = build_grr_day(mf, tb, do)
                iso = day["business_date"]
                days[iso] = normalize(day)
                added.append(iso)
                print(f"  parsed {iso} (reconciliation diff {day['totals']['diff']:.2f})")
            except Exception as e:
                print(f"  parse error in {d}: {e}")
    # also accept a single set dropped at repo root (manual/testing)
    elif all(os.path.exists(f) for f in (MF, TB, DO)):
        try:
            day = build_grr_day(MF, TB, DO)
            days[day["business_date"]] = normalize(day)
            added.append(day["business_date"])
        except Exception as e:
            print(f"  parse error: {e}")
    # Days that arrived without a Manager Flash: rebuild the rooms side from the
    # R106 history and the revenue side from the Trial Balance. Flagged
    # provisional, and replaced the moment a real report set turns up.
    try:
        import provisional_days
        added += provisional_days.fold(days, HF)
    except Exception as e:
        print(f"  provisional days failed: {e}")

    if added:
        store["days"] = days
        json.dump(store, open(STORE, "w"), indent=0, default=str)
        print(f"Added/updated {len(set(added))} day(s): {sorted(set(added))}")

    latest = max(days)
    reports = {}
    for iso in days:
        if iso.startswith("2026"):
            if iso == anchor_iso:
                reports[iso] = anchor_sections
            else:
                reports[iso] = build_report_for(iso, days, anchor_iso, anchor_sections)

    out_days = {}
    for iso, v in days.items():
        rec = dict(v)
        if iso in reports:
            rec["report"] = reports[iso]
        out_days[iso] = rec

    latest_report = reports.get(latest, anchor_sections)
    data = {
        "property": "Courtyard by Marriott Vadodara",
        "generated_at": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "latest_date": latest,
        "days": out_days,
        "sections": latest_report,
        "insights": build_insights(latest_report),
    }
    json.dump(data, open(DATA, "w"), separators=(",", ":"), default=str)
    print(f"Wrote {DATA}: {len(out_days)} days, latest {latest}")

    # STR lives in its own persistent json file, fed from the same inbox/ that
    # fetch_email.py filled. Errors here must never break the DBR build, which
    # is already written above. (H&F ran at the top of main.)
    try:
        update_star()
    except Exception as e:
        print(f"update_star failed: {e}")

    # Generate the downloadable GRR owner-report workbook from the report we just
    # built. Never let a failure here break the DBR data build.
    try:
        import build_grr_xlsx  # noqa
        build_grr_xlsx.build(DATA, "grr.xlsx")
    except Exception as e:
        print(f"build_grr_xlsx failed: {e}")


if __name__ == "__main__":
    main()
