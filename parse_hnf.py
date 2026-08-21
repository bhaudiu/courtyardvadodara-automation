#!/usr/bin/env python3
"""parse_hnf.py — parse Marriott R106 'History and Forecast' (HNF) PDFs into
per-snapshot, per-stay-date metrics, and build hf.json for the dashboard.

Each HNF PDF (one month) lists, for the day it was RUN (the snapshot date):
  stay date -> Total Occ. Rooms, Occ.%, Room Revenue, Average Rate (ADR)
split into a History section (closed days = actuals) and a Forecast section
(future days = on-the-books). Comparing the same stay date across two snapshot
run-dates gives revenue-management PICK-UP.

Input layout (as forwarded): <DD-MM-YYYY>/<Morning|Evening>/HNF <Mon> 2026.pdf
Output: hf.json = {latest, dates:[snap...], snaps:{snap_iso:{stay_iso:{r,rev,adr,occ,sec}}}}
"""
import pdfplumber, re, os, glob, json, sys


def to_num(s):
    s = str(s).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def parse_hnf(path):
    with pdfplumber.open(path) as pdf:
        text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    out, section = {}, None
    for ln in text.splitlines():
        s = ln.strip()
        if s == "History":
            section = "history"; continue
        if s == "Forecast":
            section = "forecast"; continue
        m = re.match(r"^(\d{2})-(\d{2})-(\d{2})\s+[A-Za-z]{3}\s+(.*)$", s)
        if not m:
            continue
        d, mo, y = m.group(1), m.group(2), m.group(3)
        iso = f"20{y}-{mo}-{d}"
        toks = m.group(4).split()
        if not toks:
            continue
        rooms = to_num(toks[0])
        pi = next((i for i, t in enumerate(toks) if t.endswith("%")), None)
        if pi is None:
            continue
        occ = to_num(toks[pi][:-1])
        rev = to_num(toks[pi + 1]) if pi + 1 < len(toks) else None
        adr = to_num(toks[pi + 2]) if pi + 2 < len(toks) else None
        out[iso] = {"r": rooms, "rev": rev, "adr": adr,
                    "occ": (occ / 100.0 if occ is not None else None), "sec": section}
    return out


def hnf_run_date(path):
    """The date the report was RUN (snapshot date), from the page header
    e.g. 'Courtyard By Marriott Vadodara 19-08-26' -> '2026-08-19'."""
    try:
        with pdfplumber.open(path) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return None
    head = "\n".join(text.splitlines()[:4])
    m = re.search(r"(\d{2})-(\d{2})-(\d{2})", head)
    if m:
        d, mo, y = m.groups()
        return f"20{y}-{mo}-{d}"
    return None


def parse_hnf_folder(folder):
    """Merge every HNF PDF in a folder into one snapshot; return (run_date_iso, data)."""
    pdfs = [p for p in sorted(glob.glob(os.path.join(folder, "*.pdf")))
            if ("hnf" in os.path.basename(p).lower() or "history" in os.path.basename(p).lower())]
    if not pdfs:
        return None, {}
    run_date, merged = None, {}
    for p in pdfs:
        if run_date is None:
            run_date = hnf_run_date(p)
        try:
            merged.update(parse_hnf(p))
        except Exception as e:
            print(f"  hnf parse error {os.path.basename(p)}: {e}")
    return run_date, merged


def snapshot_for_daydir(datedir, months=("Aug", "Sep", "Oct")):
    """Prefer Morning; fall back to Evening. Merge all month PDFs."""
    subs = {x.lower(): x for x in os.listdir(datedir) if os.path.isdir(os.path.join(datedir, x))}
    order = [subs[k] for k in ("morning", "evening") if k in subs]
    if not order:  # files directly in the day folder
        order = [None]
    for sub in order:
        base = os.path.join(datedir, sub) if sub else datedir
        merged = {}
        for mon in months:
            f = os.path.join(base, f"HNF {mon} 2026.pdf")
            if os.path.exists(f):
                merged.update(parse_hnf(f))
        if merged:
            return merged
    return {}


def build(root, out_path):
    snaps = {}
    for datedir in sorted(glob.glob(os.path.join(root, "[0-3][0-9]-[0-1][0-9]-2026"))):
        name = os.path.basename(datedir)
        dd, mm, yy = name.split("-")
        snap_iso = f"{yy}-{mm}-{dd}"
        data = snapshot_for_daydir(datedir)
        if data:
            snaps[snap_iso] = data
    dates = sorted(snaps.keys())
    doc = {"latest": dates[-1] if dates else None, "dates": dates, "snaps": snaps}
    with open(out_path, "w") as f:
        json.dump(doc, f, separators=(",", ":"))
    return doc


if __name__ == "__main__":
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    out = sys.argv[2] if len(sys.argv) > 2 else "hf.json"
    d = build(root, out)
    print(f"hf.json: {len(d['snaps'])} snapshots, latest {d['latest']}, "
          f"{sum(len(v) for v in d['snaps'].values())} stay-date rows")
