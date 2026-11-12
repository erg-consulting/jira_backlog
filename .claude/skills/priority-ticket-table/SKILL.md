---
name: priority-ticket-table
description: Rebuild the "Dev Priorities" summary in a Jira backlog workbook (.xlsx) - adds Priority standardized and Segment columns plus helper columns to the export sheet, and creates or refreshes a "Dev Priorities" tab with live formulas giving per-segment ticket counts and the top-priority tickets with owners. Use when the user runs /priority-ticket-table or pastes a fresh Jira export into a workbook and asks to rebuild the development priorities summary.
---

# Priority ticket table

Input: an .xlsx workbook containing a Jira backlog export. If the user gives no path, ask for it. The user has said the sheet is not Bain-branded: do not apply Bain styling, do not load the Bain Excel skill, and do not ask about it.

## Run

```bash
python3 .claude/skills/priority-ticket-table/scripts/build_priority_table.py "<workbook.xlsx>" [--out "<out.xlsx>" | --in-place]
```

Do not build the workbook by hand or compute the numbers yourself; use the script. Defaults (project prefix, excluded status, tickets per segment, segment lookup, priority map) live in `config.json`. By default the script writes `<workbook>_priorities.xlsx` and leaves the input untouched; use `--in-place` only if the user asks.

The script:
1. Checks that the source columns match: Assignee (A), Created (C), Labels (D-I), Status (J), Priority (K), Issue key (M), Summary (P), Resolution (Q). If any differ it changes nothing and prints the mismatches; tell the user which columns differ and stop.
2. Adds to the source sheet (columns A-Q are never changed): `Priority standardized` (R) and `Segment` (S), a priority lookup table and the label prefix cell (U:V), and four helper columns after one empty column: `Created (date)` (X), `Priority rank` (Y), `Ticket line` (Z), `Open & groomed` (AA). Everything is a live formula, filled to the last data row. A re-run on the same workbook replaces these columns; it refuses to write over any non-empty cell in R:AA that it did not create.
3. Creates or refreshes the `Dev Priorities` tab in place: parameter table, segment lookup, and a summary table with one row per business segment only (Segment, Description, Total tickets, top-N ticket lines, Recent notes/updates); there are no "Multi-repo / multi-team" or "Hackathon feedback" rows. On a refresh it keeps what the user typed: parameter values, segment lookup edits and notes. Extra rows the user typed below the table are kept only if they contain text; empty extra rows (such as the old manual rows) are removed and the script says so.
4. Writes `<output>_slide_table.html`: the summary table as one inline-styled HTML table (grey bold header, Calibri) for copy/paste into PowerPoint. The values come from the independent Python calculation, which the cross-check compares with the workbook formulas, and notes come from the tab. Send it to the user with SendUserFile (display: render) so they can select the table, copy it and paste it into a slide; also attach the .xlsx.
5. Recalculates a copy in LibreOffice and compares every Segment, Priority rank, Ticket line, Open & groomed and Created value, and each segment's total and top list, with an independent Python calculation. The report ends with PASS or the mismatches. If it does not say PASS, do not present the numbers as verified.

## Report to the user

Use the script's output. Give:
1. The ticket count for each segment ("Open & groomed" tickets carrying that BU label).
2. Tickets with more than one BU label: key, labels, and whether the ticket is Open & groomed (only those appear in the summary, once per label).
3. Values that did not map cleanly: Priority values not in the map (shown as `UNMAPPED`) and blank priorities among open tickets; BU labels not in the segment lookup, with how many tickets carry them; labels cleaned of stray characters or in the wrong case; statuses with stray whitespace or a near-miss of the excluded status (and how many rows match the excluded status exactly); Created values that could not be parsed; rows whose key does not start with the project prefix; open tickets with no BU label.
4. Anything that cannot be confirmed: say if LibreOffice was unavailable and the cross-check did not run.

## Behaviours to know

- Segment matching is exact on the whole label, so `BU_AN` does not match `BU_RAN`. Label prefix and Priority/Status comparisons are case-insensitive, as in Excel.
- "Resolution is blank" means an empty cell. A literal "Unresolved" text value would not count as open.
- Priority rank is the position in the lookup table (P0 = 1 ... Lower priority = 4); blank or unmapped priority = 99. Ties on rank are ordered newest Created first, then sheet order.
- Formulas run to the last data row at build time. After pasting a new export, run the skill again so the ranges cover the new rows.
- At most 10 tickets per segment are supported; the script stops if the parameter is higher.
- The formulas (TEXTJOIN) need Excel 2019 or Microsoft 365. They were verified by LibreOffice recalculation, not in Excel itself. The saved file has no cached values, so Excel calculates on opening.
- Do not commit workbooks to the repo; they contain real backlog data (`*.xlsx` and `*.csv` are git-ignored).

To change defaults (segments, project prefix, priority map), edit `config.json` only after the user confirms; for an existing tab, edit the blue input cells in the workbook instead.
