#!/usr/bin/env python3
"""Query versioned Jira backlog CSV snapshots in backlog_history/.

  query_backlog.py list
  query_backlog.py add <export.csv>
  query_backlog.py query [--match REGEX] [--in summary|summary,description] [--label LABEL]
                         [--status S ...] [--bucket in_progress|to_be_reviewed|done|excluded|other]
                         [--snapshot DATE|FILE] [--tickets] [--json]
  query_backlog.py trend  [same filters]      # count + status breakdown for every snapshot
Default snapshot: the latest. --match is a case-insensitive regex (e.g. "multi[- ]?repo|multiple repo|cross[- ]?repo").
"""
import argparse, csv, json, re, shutil, sys
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
HIST = ROOT / "backlog_history"
TELE = json.loads((ROOT / ".claude/skills/jira-backlog-teardown/config.json").read_text())
norm = lambda s: (s or "").strip().lower()
TS = re.compile(r"(\d{4}-\d{2}-\d{2})T(\d{2})[_:]?(\d{2})")


def snap_time(p):
    m = TS.search(p.name)
    if m:
        return datetime.strptime(f"{m[1]} {m[2]}:{m[3]}", "%Y-%m-%d %H:%M")
    return datetime.fromtimestamp(p.stat().st_mtime)


def snapshots():
    return sorted(HIST.glob("*.csv"), key=snap_time)


def pick(sel):
    s = snapshots()
    if not s:
        sys.exit(f"No CSV snapshots in {HIST}. Add one with: query_backlog.py add <export.csv>")
    if not sel:
        return s[-1]
    hits = [p for p in s if sel in p.name or snap_time(p).strftime("%Y-%m-%d").startswith(sel)]
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
                "priority": g(row, "priority"), "assignee": g(row, "assignee") or "unassigned",
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
    ap.add_argument("--snapshot"); ap.add_argument("--tickets", action="store_true"); ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    if a.cmd == "add":
        if not a.file: sys.exit("add needs a CSV path")
        src = Path(a.file); load(src)  # validates columns
        HIST.mkdir(exist_ok=True)
        dst = HIST / f"backlog_{snap_time(src).strftime('%Y-%m-%dT%H%M')}.csv"
        if dst.exists(): sys.exit(f"{dst.name} already exists; not overwriting")
        shutil.copy2(src, dst); print(f"Added {dst.name} ({len(load(dst))} rows)"); return
    if a.cmd == "list":
        for p in snapshots(): print(f"{snap_time(p):%Y-%m-%d %H:%M}  {p.name}  {len(load(p))} rows")
        return
    if a.cmd == "trend":
        res = [{"snapshot": p.name, "date": f"{snap_time(p):%Y-%m-%d %H:%M}", **summarize(filt(load(p), a))} for p in snapshots()]
        print(json.dumps(res, indent=2) if a.json else "\n".join(f"{r['date']}  {r['snapshot']}: {r['count']} tickets {r['by_status']}" for r in res)); return
    p = pick(a.snapshot); rows = filt(load(p), a)
    res = {"snapshot": p.name, "date": f"{snap_time(p):%Y-%m-%d %H:%M}", "filters": {k: v for k, v in vars(a).items() if k in ("match", "fields", "label", "status", "bucket") and v}, **summarize(rows)}
    if a.tickets: res["tickets"] = [{k: t[k] for k in ("key", "status", "resolution", "priority", "assignee", "summary")} for t in rows]
    if a.json: print(json.dumps(res, indent=2)); return
    print(f"Snapshot {res['snapshot']} ({res['date']}) | filters {res['filters']}\nMatching tickets: {res['count']}\nBy status:")
    for s, n in res["by_status"].items(): print(f"  {s}: {n}")
    print("By bucket: " + ", ".join(f"{k}: {v}" for k, v in res["by_bucket"].items()))
    for t in res.get("tickets", []): print(f"  {t['key']} | {t['status']} | {t['priority']} | {t['assignee']} | {t['summary'][:90]}")


main()
