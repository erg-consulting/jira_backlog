#!/usr/bin/env python3
"""Count Jira backlog tickets by bucket. Usage: teardown.py <file.csv> [--json]"""
import csv, json, sys
from collections import Counter
from pathlib import Path

CONFIG = Path(__file__).resolve().parent.parent / "config.json"


def norm(s):
    return (s or "").strip().lower()


def main(path, as_json=False):
    cfg = json.loads(CONFIG.read_text())
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        header = {norm(h): h for h in reader.fieldnames or []}
        missing = [c for c in (cfg["status_column"], cfg["resolution_column"]) if norm(c) not in header]
        if missing:
            sys.exit(f"ERROR: required column(s) not found: {missing}. Columns present: {list(header.values())}")
        sc, rc = header[norm(cfg["status_column"])], header[norm(cfg["resolution_column"])]
        rows = list(reader)

    tbr = {norm(s) for s in cfg["to_be_reviewed"]["statuses"]}
    ip_s = {norm(s) for s in cfg["in_progress"]["statuses"]}
    ip_r = {norm(r) for r in cfg["in_progress"]["resolutions"]}
    done = {norm(s) for s in cfg["done"]["statuses"]}
    excl = {norm(s) for s in cfg["excluded_statuses"]}

    counts = Counter()
    unclassified = Counter()
    ip_breakdown = Counter()
    for r in rows:
        st, res = norm(r[sc]), norm(r[rc])
        if st in tbr:
            counts["to_be_reviewed"] += 1
        elif st in ip_s and res in ip_r:
            counts["in_progress"] += 1
            ip_breakdown[r[sc].strip()] += 1
        elif st in done:
            counts["done"] += 1
        elif st in excl:
            counts["excluded"] += 1
        else:
            unclassified[(r[sc].strip(), r[rc].strip())] += 1

    n_unc = sum(unclassified.values())
    total = len(rows)
    out = {
        "total_rows": total,
        "to_be_reviewed": counts["to_be_reviewed"],
        "in_progress": counts["in_progress"],
        "done": counts["done"],
        "in_progress_by_status": dict(ip_breakdown),
        "excluded": counts["excluded"],
        "unclassified": n_unc,
        "unclassified_by_status_resolution": {f"{s} | {r or '(blank)'}": n for (s, r), n in unclassified.most_common()},
        "reconciles": counts["to_be_reviewed"] + counts["in_progress"] + counts["done"] + counts["excluded"] + n_unc == total,
    }
    if as_json:
        print(json.dumps(out, indent=2))
        return
    print(f"Rows: {total}")
    print(f"To be reviewed: {out['to_be_reviewed']}")
    print(f"In progress:    {out['in_progress']}  {dict(ip_breakdown)}")
    print(f"Done:           {out['done']}")
    print(f"Excluded ({', '.join(cfg['excluded_statuses'])}): {out['excluded']}")
    print(f"Unclassified:   {n_unc}")
    for k, n in out["unclassified_by_status_resolution"].items():
        print(f"    {k}: {n}")
    print(f"Reconciles to total rows: {out['reconciles']}")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        sys.exit(__doc__)
    main(args[0], "--json" in sys.argv)
