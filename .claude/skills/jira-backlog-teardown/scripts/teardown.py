#!/usr/bin/env python3
"""Count Jira backlog tickets by bucket and in-progress tickets by standardized priority.
Usage: teardown.py <file.csv> [--json] [--out enriched.csv]
--out writes a copy of the data with the "priority standardized" column added."""
import csv, json, sys
from collections import Counter
from pathlib import Path

CONFIG = Path(__file__).resolve().parent.parent / "config.json"


def norm(s):
    return (s or "").strip().lower()


def main(path, as_json=False, out_path=None):
    cfg = json.loads(CONFIG.read_text())
    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        header = {norm(h): h for h in reader.fieldnames or []}
        required = (cfg["status_column"], cfg["resolution_column"], cfg["priority_column"])
        missing = [c for c in required if norm(c) not in header]
        if missing:
            sys.exit(f"ERROR: required column(s) not found: {missing}. Columns present: {list(header.values())}")
        sc, rc = header[norm(cfg["status_column"])], header[norm(cfg["resolution_column"])]
        rows = list(reader)
        fieldnames = list(reader.fieldnames)
    pc = header[norm(cfg["priority_column"])]
    pmap = {norm(k): v for k, v in cfg["priority_map"].items()}
    new_col = cfg["priority_standardized_column"]
    for r in rows:
        r[new_col] = pmap.get(norm(r[pc]), "UNMAPPED")
    if out_path:
        with open(out_path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fieldnames + [new_col])
            w.writeheader()
            w.writerows(rows)

    tbr = {norm(s) for s in cfg["to_be_reviewed"]["statuses"]}
    ip_s = {norm(s) for s in cfg["in_progress"]["statuses"]}
    ip_r = {norm(r) for r in cfg["in_progress"]["resolutions"]}
    done = {norm(s) for s in cfg["done"]["statuses"]}
    excl = {norm(s) for s in cfg["excluded_statuses"]}

    counts = Counter()
    unclassified = Counter()
    ip_breakdown = Counter()
    ip_priority = Counter()
    for r in rows:
        st, res = norm(r[sc]), norm(r[rc])
        if st in tbr:
            counts["to_be_reviewed"] += 1
        elif st in ip_s and res in ip_r:
            counts["in_progress"] += 1
            ip_breakdown[r[sc].strip()] += 1
            ip_priority[r[new_col]] += 1
        elif st in done:
            counts["done"] += 1
        elif st in excl:
            counts["excluded"] += 1
        else:
            unclassified[(r[sc].strip(), r[rc].strip())] += 1

    n_unc = sum(unclassified.values())
    total = len(rows)
    levels = list(dict.fromkeys(cfg["priority_map"].values()))
    prio = {lvl: ip_priority.get(lvl, 0) for lvl in levels}
    unmapped = {k: v for k, v in ip_priority.items() if k not in levels}
    out = {
        "total_rows": total,
        "to_be_reviewed": counts["to_be_reviewed"],
        "in_progress": counts["in_progress"],
        "done": counts["done"],
        "in_progress_by_status": dict(ip_breakdown),
        "in_progress_by_priority": prio,
        "in_progress_unmapped_priority": unmapped,
        "excluded": counts["excluded"],
        "not_in_buckets": counts["excluded"] + n_unc,
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
    print("In-progress tickets by priority standardized:")
    for lvl, n in prio.items():
        print(f"    {lvl}: {n}")
    if unmapped:
        print(f"    UNMAPPED priority values (not in config): {unmapped}")
    print(f"Done:           {out['done']}")
    print(f"Not in the three buckets: {out['not_in_buckets']}")
    print(f"Excluded ({', '.join(cfg['excluded_statuses'])}): {out['excluded']}")
    print(f"Unclassified:   {n_unc}")
    for k, n in out["unclassified_by_status_resolution"].items():
        print(f"    {k}: {n}")
    print(f"Reconciles to total rows: {out['reconciles']}")


if __name__ == "__main__":
    argv = sys.argv[1:]
    out_path = None
    if "--out" in argv:
        i = argv.index("--out")
        if i + 1 >= len(argv):
            sys.exit(__doc__)
        out_path = argv[i + 1]
        del argv[i:i + 2]
    args = [a for a in argv if not a.startswith("--")]
    if len(args) != 1:
        sys.exit(__doc__)
    main(args[0], "--json" in argv, out_path)
