---
name: jira-backlog-teardown
description: Count tickets in a Jira backlog CSV export by bucket - "To be reviewed", "In progress" and "Done" - and break in-progress tickets down by standardized priority (P0, P1, P2, Lower priority). Use when the user runs /jira-backlog-teardown or provides a Jira backlog CSV (with Status and Resolution columns) and asks for ticket counts.
---

# Jira backlog teardown

Input: path to a Jira backlog CSV export (one row per ticket, with `Status` and `Resolution` columns). If the user gives no path, ask for it.

## Run

```bash
python3 .claude/skills/jira-backlog-teardown/scripts/teardown.py "<path-to-csv>"
```

Do not compute the counts by reading the CSV yourself; use the script. Bucket rules live in `config.json` (matching is case-insensitive, exact text):

| Bucket | Rule |
|---|---|
| To be reviewed | Status = "Not started" |
| In progress | Status in In Analysis, On Hold, In Testing, In Development, In Review, **and** Resolution is blank or "Unresolved" (Jira leaves Resolution blank for open tickets) |
| Done | Status = "Done" |

### Priority standardized

A `priority standardized` column is derived from `Priority`: Highest -> P0, High -> P1, Medium -> P2, Low -> Lower priority. It is counted over the **in-progress** tickets only (the same filter as the "In progress" bucket: unresolved and status in In Analysis, On Hold, In Testing, In Development, In Review). Any Priority value not in the map shows as `UNMAPPED`; report it and ask before extending the map.

To also get the data with the new column added, pass `--out <path>.csv` (keep the output out of git; CSVs are ignored).

"Obsolete" is excluded. Every other status is reported as unclassified.

## Report to the user

1. State the three counts, with the in-progress breakdown by status, then the in-progress counts by standardized priority (P0, P1, P2, Lower priority; they must sum to the In progress count).
2. State total rows, excluded count, and unclassified count with the status values behind it. Unclassified rows are not in any bucket; if a new or unexpected status appears (for example "In Progress", "To Do", "New"), say so and ask whether it should be mapped. Do not remap it silently.
3. Confirm the counts reconcile to total rows. If `reconciles` is false or the script errors on missing columns, report that and stop.

To change a mapping, edit `config.json` only after the user confirms.

## Data handling

Backlog exports can contain real client data. Do not commit CSVs to the repo unless the user explicitly asks.
