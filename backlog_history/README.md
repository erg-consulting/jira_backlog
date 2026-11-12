# backlog_history

Drop each Jira backlog CSV export here (one file per export date). Query them with
`.claude/skills/backlog-query/scripts/query_backlog.py` (or ask Claude, which uses the `backlog-query` skill).

- Add an export: `python3 .claude/skills/backlog-query/scripts/query_backlog.py add "<export.csv>"`
  (copies it in as `backlog_<YYYY-MM-DD>T<HHMM>.csv`, taking the time from the Jira filename, else file mtime).
- Files are never edited or deleted by the query script.
- `*.csv` is git-ignored (the exports hold real backlog data). Snapshots stay local unless you decide to
  commit them; in an ephemeral cloud session, local files are lost when the session ends.
