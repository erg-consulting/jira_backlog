---
name: backlog-query
description: Answer questions about the Jira backlog across versioned CSV snapshots in backlog_history/ - e.g. "how many tickets are on multi-repo and what are their statuses", "how has that changed since last export", "add this export". Use when the user asks counts/statuses/lists of tickets by keyword, label or status, or wants to add a new backlog snapshot.
---

# Backlog query

Snapshots live in `backlog_history/` (one CSV per export). Never answer by reading the CSVs yourself; use the script:

```bash
python3 .claude/skills/backlog-query/scripts/query_backlog.py list
python3 .claude/skills/backlog-query/scripts/query_backlog.py add "<export.csv>"
python3 .claude/skills/backlog-query/scripts/query_backlog.py query --match "multi[- ]?repo|multiple repo|cross[- ]?repo" --tickets
python3 .claude/skills/backlog-query/scripts/query_backlog.py trend --match "multi[- ]?repo|multiple repo|cross[- ]?repo"
```

## When the user gives you a new backlog CSV
Always add it to the history first (never ask the user to rename it):
1. Run `add "<csv>"`. Jira's own filename carries the export time, so it is normally enough; the file is stored as `backlog_<date>T<HHMMSS><+ZZZZ>.csv`.
2. If it prints `NEEDS_DATE`, ask the user when the export was taken, then re-run with `--date YYYY-MM-DD[THH:MM[+ZZZZ]]`. Do not guess from file modification time.
3. If it prints `DUPLICATE`, tell the user it matches an existing snapshot and do not add it (use `--force` only if they insist).
4. Relay every `WARNING` (older than the latest snapshot, Created dates disagreeing with export times, fewer tickets than before) and ask the user to confirm the date if one fires.
The latest snapshot is the one with the most recent export time (offset-aware); `list` shows the order.

Filters: `--match REGEX` (case-insensitive), `--in summary|summary,description` (default summary only), `--label`, `--status A B`, `--bucket in_progress|to_be_reviewed|done|excluded|other`, `--snapshot DATE|FILE` (default latest), `--json`.
Buckets use the rules in `.claude/skills/jira-backlog-teardown/config.json`.

## Reporting rules
- State which snapshot (date) the answer comes from.
- State the matching rule (regex, fields searched). "Multi-repo" is a keyword match, not a Jira label; say so, and mention how many extra tickets match if description is included (re-run with `--in summary,description`) when it changes the answer.
- Give the count and the status breakdown, and the ticket list when asked. Statuses outside the teardown buckets (To Do, New, In Progress, Ready for Development) appear as bucket `other`; call them out.
- For "over time" questions use `trend`; note snapshots are only as frequent as the exports added.
- If `backlog_history/` is empty, ask the user for a CSV and `add` it.
- Do not commit CSVs unless the user explicitly asks (`*.csv` is git-ignored; they hold real backlog data).
