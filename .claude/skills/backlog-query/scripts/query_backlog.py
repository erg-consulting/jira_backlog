#!/usr/bin/env python3
"""Query versioned Jira backlog CSV snapshots in backlog_history/.

  query_backlog.py list
  query_backlog.py add <export.csv> [--date YYYY-MM-DD[THH:MM[+ZZZZ]]] [--force]
  query_backlog.py query [--match REGEX] [--in summary|summary,description] [--label LABEL]
                         [--status S ...] [--bucket in_progress|to_be_reviewed|done|excluded|other]
                         [--snapshot DATE|FILE] [--tickets] [--json]
  query_backlog.py trend  [same filters]      # count + status breakdown for every snapshot
Default snapshot: the latest. --match is a case-insensitive regex (e.g. "multi[- ]?repo|multiple repo|cross[- ]?repo").
"""
import argparse, csv, hashlib, json, re, shutil, sys
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
HIST = ROOT / "backlog_history"
TELE = json.loads((ROOT / ".claude/skills/jira-backlog-teardown/config.json").read_text())
norm = lambda s: (s or "").strip().lower()
# Filename timestamp forms (all give an aware datetime, so snapshots sort in true time order):
#   normalized   backlog_2026-11-13T001909+0200.csv
#   Jira export  ...2026-11-13T00_19_090200.csv   (HH_MM_SS then UTC offset, sign omitted = +)
#   legacy/loose ...2026-11-13T0019.csv, ...T00:19.csv (offset unknown -> treated as UTC)
PATTERNS = [
    re.compile(r"(\d{4}-\d{2}-\d{2})T(\d{2})(\d{2})(\d{2})([+-]\d{4})(?!\d)"),
    re.compile(r"(\d{4}-\d{2}-\d{2})T(\d{2})_(\d{2})_(\d{2})([+-]?\d{4})(?!\d)"),
    re.compile(r"(\d{4}-\d{2}-\d{2})T(\d{2})[_:]?(\d{2})()()(?!\d)"),
]


def parse_ts(text):
    """Return an aware datetime from a filename/--date string, or None."""
    for rx in PATTERNS:
        m = rx.search(text)
        if m:
            d, hh, mm, ss, off = m.groups()
            if off and off[0] not in "+-":
                off = "+" + off
            sign = -1 if off.startswith("-") else 1
            tz = timezone(sign * timedelta(hours=int(off[1:3]), minutes=int(off[3:5]))) if off else timezone.utc
            return datetime.strptime(f"{d} {hh}:{mm}:{ss or '00'}", "%Y-%m-%d %H:%M:%S").replace(tzinfo=tz)
    return None


def parse_date_arg(v):
    m = re.fullmatch(r"(\d{4}-\d{2}-\d{2})(?:T(\d{2}):?(\d{2})(?::?(\d{2}))?)?([+-]\d{4})?", v.strip())
    if not m:
        sys.exit("--date must look like 2026-11-13, 2026-11-13T00:19 or 2026-11-13T00:19+0200")
    d, hh, mm, ss, off = m.groups()
    return parse_ts(f"{d}T{hh or '00'}{mm or '00'}{ss or '00'}{off or '+0000'}")


def snap_time(p):
    t = parse_ts(p.name)
    return t if t else datetime.fromtimestamp(p.stat().st_mtime, tz=timezone.utc)


def fmt_time(t):
    return t.strftime("%Y-%m-%d %H:%M %z")


def norm_name(t):
    return f"backlog_{t.strftime('%Y-%m-%dT%H%M%S%z')}.csv"


def has_ts(p):
    return parse_ts(p.name) is not None


def max_created(rows):
    best = None
    for t in rows:
        m = re.match(r"^\s*(\d+)\.(\d+)\.(\d{4}) (\d+):(\d{2}) ?([AaPp][Mm])\s*$", t["created"])
        if m:
            d, mo, y, h, mi, ap = m.groups()
            dt = datetime(int(y), int(mo), int(d), int(h) % 12 + (12 if ap.upper() == "PM" else 0), int(mi))
            best = dt if best is None or dt > best else best
    return best


def signature(rows):
    return hashlib.sha256(json.dumps(sorted((t["key"], t["status"], t["resolution"], t["priority"], t["assignee"], t["summary"]) for t in rows)).encode()).hexdigest()


def snapshots():
    return sorted(HIST.glob("*.csv"), key=snap_time)


def pick(sel):
    s = snapshots()
    if not s:
        sys.exit(f"No CSV snapshots in {HIST}. Add one with: query_backlog.py add <export.csv>")
    if not sel:
        return s[-1]
    hits = [p for p in s if sel in p.name or snap_time(p).strftime("%Y-%m-%d").startswith(sel) or sel in fmt_time(snap_time(p))]
    if len(hits) != 1:
        sys.exit(f"--snapshot '{sel}' matched {len(hits)} files: {[p.name for p in hits]}")
    return hits[0]


def load(p):
    with open(p, newline="", encoding="utf-8-sig") as fh:
        r = csv.reader(fh)
        h = next(r)
        idx = {}
        for i, c in enumerate(h):
            idx.setdefault(norm(c), []).append(i)
        need = ["status", "resolution", "summary"]
        key = "issue key" if "issue key" in idx else "key"
        miss = [c for c in need + [key] if c not in idx]
        if miss:
            sys.exit(f"{p.name}: missing columns {miss}")
        g = lambda row, c: row[idx[c][0]].strip() if c in idx and idx[c][0] < len(row) else ""
        out = []
        for row in r:
            if not any(row):
                continue
            out.append({
                "key": g(row, key), "status": g(row, "status"), "resolution": g(row, "resolution"),
                "priority": g(row, "priority"), "created": g(row, "created"), "assignee": g(row, "assignee") or "unassigned",
                "summary": g(row, "summary"), "description": g(row, "description"),
                "labels": [row[i].strip() for i in idx.get("labels", []) if i < len(row) and row[i].strip()],
            })
        return out


def bucket(t):
    st, res = norm(t["status"]), norm(t["resolution"])
    if st in {norm(x) for x in TELE["to_be_reviewed"]["statuses"]}: return "to_be_reviewed"
    if st in {norm(x) for x in TELE["in_progress"]["statuses"]} and res in {norm(x) for x in TELE["in_progress"]["resolutions"]}: return "in_progress"
    if st in {norm(x) for x in TELE["done"]["statuses"]}: return "done"
    if st in {norm(x) for x in TELE["excluded_statuses"]}: return "excluded"
    return "other"


def filt(rows, a):
    rx = re.compile(a.match, re.I) if a.match else None
    fields = [f.strip() for f in a.fields.split(",")]
    out = []
    for t in rows:
        if rx and not any(rx.search(t[f]) for f in fields): continue
        if a.label and norm(a.label) not in {norm(l) for l in t["labels"]}: continue
        if a.status and norm(t["status"]) not in {norm(s) for s in a.status}: continue
        if a.bucket and bucket(t) != a.bucket: continue
        out.append(t)
    return out


def summarize(rows):
    return {"count": len(rows), "by_status": dict(Counter(t["status"] for t in rows).most_common()),
            "by_bucket": dict(Counter(bucket(t) for t in rows).most_common())}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["list", "add", "query", "trend"])
    ap.add_argument("file", nargs="?")
    ap.add_argument("--match"); ap.add_argument("--in", dest="fields", default="summary")
    ap.add_argument("--label"); ap.add_argument("--status", nargs="+"); ap.add_argument("--bucket")
    ap.add_argument("--snapshot"); ap.add_argument("--date"); ap.add_argument("--force", action="store_true"); ap.add_argument("--tickets", action="store_true"); ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    if a.cmd == "add":
        if not a.file: sys.exit("add needs a CSV path")
        src = Path(a.file); new_rows = load(src)  # validates columns
        t = parse_date_arg(a.date) if a.date else parse_ts(src.name)
        if t is None:
            sys.exit("NEEDS_DATE: no export timestamp in the filename. Ask the user when this export was taken, "
                     "then re-run with --date YYYY-MM-DD[THH:MM[+ZZZZ]]. (File modification time is not trusted.)")
        HIST.mkdir(exist_ok=True)
        dst = HIST / norm_name(t)
        existing = snapshots()
        sig, digest = signature(new_rows), hashlib.sha256(src.read_bytes()).hexdigest()
        warnings = []
        for p in existing:
            if hashlib.sha256(p.read_bytes()).hexdigest() == digest or signature(load(p)) == sig:
                sys.exit(f"DUPLICATE: same ticket content as {p.name}; not added (use --force to add anyway).") if not a.force else warnings.append(f"same ticket content as {p.name}")
        if dst.exists(): sys.exit(f"{dst.name} already exists; not overwriting")
        if existing:
            last = existing[-1]; last_rows = load(last)
            if t < snap_time(last):
                warnings.append(f"export time {fmt_time(t)} is OLDER than the latest snapshot {last.name} ({fmt_time(snap_time(last))}); it will not become the default")
            nc, lc = max_created(new_rows), max_created(last_rows)
            if nc and lc and ((nc < lc and t > snap_time(last)) or (nc > lc and t < snap_time(last))):
                warnings.append(f"newest Created date ({nc:%Y-%m-%d %H:%M}) {'is older' if nc < lc else 'is newer'} than in {last.name} ({lc:%Y-%m-%d %H:%M}), which disagrees with the export times; check the date")
            warnings.append(f"ticket count {len(new_rows)} vs {len(last_rows)} in {last.name}") if len(new_rows) < len(last_rows) else None
        shutil.copy2(src, dst); print(f"Added {dst.name} ({len(new_rows)} rows, exported {fmt_time(t)})")
        for w in filter(None, warnings): print("WARNING:", w)
        return
    if a.cmd == "list":
        for p in snapshots(): print(f"{fmt_time(snap_time(p))}  {p.name}  {len(load(p))} rows{'' if has_ts(p) else '  (time from file mtime: no timestamp in name)'}")
        return
    if a.cmd == "trend":
        res = [{"snapshot": p.name, "date": fmt_time(snap_time(p)), **summarize(filt(load(p), a))} for p in snapshots()]
        print(json.dumps(res, indent=2) if a.json else "\n".join(f"{r['date']}  {r['snapshot']}: {r['count']} tickets {r['by_status']}" for r in res)); return
    p = pick(a.snapshot); rows = filt(load(p), a)
    res = {"snapshot": p.name, "date": fmt_time(snap_time(p)), "filters": {k: v for k, v in vars(a).items() if k in ("match", "fields", "label", "status", "bucket") and v}, **summarize(rows)}
    if a.tickets: res["tickets"] = [{k: t[k] for k in ("key", "status", "resolution", "priority", "assignee", "summary")} for t in rows]
    if a.json: print(json.dumps(res, indent=2)); return
    print(f"Snapshot {res['snapshot']} ({res['date']}) | filters {res['filters']}\nMatching tickets: {res['count']}\nBy status:")
    for s, n in res["by_status"].items(): print(f"  {s}: {n}")
    print("By bucket: " + ", ".join(f"{k}: {v}" for k, v in res["by_bucket"].items()))
    for t in res.get("tickets", []): print(f"  {t['key']} | {t['status']} | {t['priority']} | {t['assignee']} | {t['summary'][:90]}")


main()
