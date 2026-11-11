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

Report exactly this, with the numbers from the script and nothing else (no excluded, unclassified or reconciliation lines):

> **Backlog teardown (<total_rows> tickets)**
> - To be reviewed: **<to_be_reviewed>**
> - In progress: **<in_progress>** (<status>: <n>, ... in descending order)
>   - P0: <n> · P1: <n> · P2: <n> · Lower priority: <n>
> - Done: **<done>**

Only deviate from this format in these cases:
- The script errors (for example a missing column) or `reconciles` is false: report that and stop.
- `in_progress_unmapped_priority` is non-empty: add one line naming the unmapped Priority values and ask before extending the map.

The script's full output still lists excluded and unclassified rows if the user asks for it.

To change a mapping, edit `config.json` only after the user confirms.

## Data handling

Backlog exports can contain real client data. Do not commit CSVs to the repo unless the user explicitly asks.
