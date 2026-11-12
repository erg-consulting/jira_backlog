#!/usr/bin/env python3
"""Build the Dev Priorities summary in a Jira backlog workbook (live formulas only).

Usage: build_priority_table.py <workbook.xlsx | export.csv> [--out OUT.xlsx | --in-place] [--template DECK.pptx] [--sheet NAME] [--no-verify] [--json]
Default output: <input>_priorities.xlsx (the input is never modified unless --in-place; --in-place is not allowed for a CSV).
A CSV is first opened as a workbook in the expected column layout (see csv_to_workbook), then processed as below.
"""
import argparse, csv, json, re, shutil, subprocess, sys, tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

CFG = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text())
MAXN = 25  # max tickets listed per segment (all P0/P1 tickets plus any top-up) the formula supports
BLUE, BLACK = Font(color="0000FF"), Font(color="000000")
HDR_FONT, HDR_FILL = Font(bold=True, color="000000"), PatternFill("solid", fgColor="D9D9D9")
WRAP = Alignment(wrap_text=True, vertical="top")
EXCEL_EPOCH = datetime(1899, 12, 30)

# new source-sheet columns
C_PS, C_SEG = "R", "S"            # after Resolution (Q)
C_LK1, C_LK2 = "U", "V"           # lookup table (T is empty)
C_CRE, C_RANK, C_LINE, C_OG, C_KEY = "X", "Y", "Z", "AA", "AB"  # helpers (W is empty)
NEW_COLS = [C_PS, C_SEG, "T", C_LK1, C_LK2, "W", C_CRE, C_RANK, C_LINE, C_OG, C_KEY]
OURS = {"Priority standardized", "Segment", "Created (date)", "Priority rank", "Ticket line", "Open & groomed", "Sort key"}
PRIO_KEY = "always_list_through"  # parameter: every ticket up to this priority is always listed


def q(name):
    return "'" + name.replace("'", "''") + "'"


def check_headers(ws):
    bad = []
    for col, want in CFG["source_columns"].items():
        got = ws[f"{col}1"].value
        if (str(got).strip().lower() if got is not None else "") != want.lower():
            bad.append(f"{col}1: expected '{want}', found '{got}'")
    return bad


def find_source(wb, name):
    dp = CFG["dev_priorities_sheet"]
    cands = [wb[name]] if name else [w for w in wb.worksheets if w.title != dp]
    best = None
    for ws in cands:
        bad = check_headers(ws)
        if not bad:
            return ws, []
        best = best or (ws, bad)
    return None, best[1] if best else ["no sheets found"]


def last_data_row(ws):
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, 18)):
            return r
    return 1


def read_existing(wb):
    """Pull user-typed state out of an existing Dev Priorities tab so a refresh keeps it."""
    st = {"params": {}, "segments": None, "notes": {}, "manual": {}}
    name = CFG["dev_priorities_sheet"]
    if name not in wb.sheetnames:
        return st
    ws = wb[name]
    keys = {"Project prefix": "project_prefix", "Excluded status": "excluded_status", "Tickets listed per segment": "tickets_per_segment",
            "Always list all tickets through priority": PRIO_KEY}
    for r in range(2, 6):
        k = keys.get(str(ws.cell(r, 1).value).strip())
        if k and ws.cell(r, 2).value not in (None, ""):
            st["params"][k] = ws.cell(r, 2).value
    segs, r = [], 7
    while ws.cell(r, 1).value not in (None, "") and ws.cell(r, 2).value not in (None, ""):
        segs.append([ws.cell(r, 1).value, ws.cell(r, 2).value]); r += 1
    if segs:
        st["segments"] = segs
    hdr = None
    for r in range(r, ws.max_row + 1):
        if str(ws.cell(r, 1).value).strip() == "Segment" and str(ws.cell(r, 2).value).strip() == "Description":
            hdr = r; break
    if hdr:
        names = {s[0] for s in (segs or [])}
        heads = [str(ws.cell(hdr, c).value or "").strip() for c in range(1, 10)]
        nc = heads.index("Recent notes/updates") + 1 if "Recent notes/updates" in heads else 5
        has_brk = "Priority breakdown" in heads
        for r in range(hdr + 1, ws.max_row + 1):
            n = ws.cell(r, 1).value
            if n in (None, ""):
                continue
            vals = [ws.cell(r, c).value for c in range(2, nc + 1)]  # Description, Total, [Breakdown], Tickets, Notes
            if not has_brk:
                vals.insert(2, None)
            if str(n).startswith("="):  # summary row for a lookup segment
                st["notes"][r] = vals[-1]
                st.setdefault("summary_rows", []).append(r)
            elif n in names:
                st["notes"][n] = vals[-1]
            else:
                st["manual"][n] = vals
        # summary rows are formulas (=A7): map their notes by the lookup order
        sr = st.pop("summary_rows", [])
        for i, rr in enumerate(sr):
            if segs and i < len(segs):
                st["notes"][segs[i][0]] = st["notes"].pop(rr)
    return st


def style(c, font=BLACK, hdr=False):
    c.font = HDR_FONT if hdr else font
    if hdr:
        c.fill = HDR_FILL
    c.alignment = WRAP


def write_source(ws, last, wb, dpq):
    """Step 1 and 2 on the source sheet. dpq = quoted Dev Priorities sheet name."""
    lo, hi = 2, last
    pm = CFG["priority_map"]
    plo, phi = 2, 1 + len(pm)
    pref_row = phi + 2
    heads = {C_PS: "Priority standardized", C_SEG: "Segment", C_LK1: "Priority", C_LK2: "Priority standardized",
             C_CRE: "Created (date)", C_RANK: "Priority rank", C_LINE: "Ticket line", C_OG: "Open & groomed", C_KEY: "Sort key"}
    for col, h in heads.items():
        style(ws[f"{col}1"], hdr=True); ws[f"{col}1"] = h
    for i, (a, b) in enumerate(pm):
        ws[f"{C_LK1}{plo+i}"], ws[f"{C_LK2}{plo+i}"] = a, b
        style(ws[f"{C_LK1}{plo+i}"], BLUE); style(ws[f"{C_LK2}{plo+i}"], BLUE)
    ws[f"{C_LK1}{pref_row}"], ws[f"{C_LK2}{pref_row}"] = "Segment label prefix", CFG["segment_label_prefix"]
    style(ws[f"{C_LK1}{pref_row}"], hdr=True); style(ws[f"{C_LK2}{pref_row}"], BLUE)
    lk = f"${C_LK1}${plo}:${C_LK2}${phi}"
    lkv = f"${C_LK2}${plo}:${C_LK2}${phi}"
    pref = f"${C_LK2}${pref_row}"
    for r in range(lo, hi + 1):
        f = {}
        f[C_PS] = f'=IF(K{r}="","",IFERROR(VLOOKUP(K{r},{lk},2,FALSE),"UNMAPPED"))'
        terms = []
        for c in "DEFGHI":
            t = f'TRIM(SUBSTITUTE(SUBSTITUTE({c}{r},";",""),",",""))'
            terms.append(f'IF(LEFT({t},LEN({pref}))={pref},{t},"")')
        f[C_SEG] = '=_xlfn.TEXTJOIN("; ",TRUE,' + ",".join(terms) + ")"
        d1 = f'FIND(".",C{r})'
        d2 = f'FIND(".",C{r},{d1}+1)'
        sp, co = f'FIND(" ",C{r})', f'FIND(":",C{r})'
        f[C_CRE] = (f'=IF(C{r}="","",IF(ISNUMBER(C{r}),C{r},IFERROR(DATE(VALUE(MID(C{r},{d2}+1,4)),'
                    f'VALUE(MID(C{r},{d1}+1,{d2}-{d1}-1)),VALUE(LEFT(C{r},{d1}-1)))'
                    f'+TIME(MOD(VALUE(MID(C{r},{sp}+1,{co}-{sp}-1)),12)+IF(UPPER(RIGHT(TRIM(C{r}),2))="PM",12,0),'
                    f'VALUE(MID(C{r},{co}+1,2)),0),"")))')
        f[C_RANK] = f'=IF({C_PS}{r}="",99,IFERROR(MATCH({C_PS}{r},{lkv},0),99))'
        f[C_LINE] = (f'=IF({C_PS}{r}="","No priority",{C_PS}{r})&": #"&IFERROR(MID(M{r},FIND("-",M{r})+1,99),M{r})'
                     f'&" – "&P{r}&" – "&IF(A{r}="","unassigned",A{r})')
        f[C_OG] = (f'=IF(AND(LEFT(M{r},LEN(DP_Project)+1)=DP_Project&"-",Q{r}="",J{r}<>DP_ExcludedStatus),1,0)')
        f[C_KEY] = f'={C_RANK}{r}*100000000+(100000000-IFERROR(ROUND({C_CRE}{r}*1440,0),0))+ROW()/10000'
        for col, formula in f.items():
            c = ws[f"{col}{r}"]; c.value = formula; style(c)
        ws[f"{C_CRE}{r}"].number_format = "dd.mm.yyyy hh:mm"
    widths = {C_PS: 20, C_SEG: 24, "T": 3, C_LK1: 22, C_LK2: 22, "W": 3, C_CRE: 18, C_RANK: 13, C_LINE: 70, C_OG: 15, C_KEY: 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    sn = q(ws.title)
    names = {"Src_Segment": C_SEG, "Src_Created": C_CRE, "Src_Rank": C_RANK, "Src_Line": C_LINE, "Src_OG": C_OG, "Src_Key": C_KEY}
    for n, col in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=f"{sn}!${col}${lo}:${col}${hi}")
    wb.defined_names["Src_PrioList"] = DefinedName("Src_PrioList", attr_text=f"{sn}!${C_LK2}${plo}:${C_LK2}${phi}")


def write_dev(wb, st, exp):
    name = CFG["dev_priorities_sheet"]
    if name in wb.sheetnames:
        ws = wb[name]
        ws.delete_rows(1, ws.max_row + 5)
    else:
        ws = wb.create_sheet(name)
    P = {**CFG["parameters"], **st["params"]}
    segs = st["segments"] or CFG["segments"]
    n = len(segs)
    N = int(P["tickets_per_segment"])
    pm = CFG["priority_map"]
    for c, h in ((1, "Parameter"), (2, "Value")):
        ws.cell(1, c).value = h; style(ws.cell(1, c), hdr=True)
    for r, (lab, key) in enumerate([("Project prefix", "project_prefix"), ("Excluded status", "excluded_status"),
                                    ("Tickets listed per segment", "tickets_per_segment"),
                                    ("Always list all tickets through priority", PRIO_KEY)], 2):
        ws.cell(r, 1).value = lab; style(ws.cell(r, 1))
        ws.cell(r, 2).value = P[key]; style(ws.cell(r, 2), BLUE)
    dp = q(name)
    wb.defined_names["DP_Project"] = DefinedName("DP_Project", attr_text=f"{dp}!$B$2")
    wb.defined_names["DP_ExcludedStatus"] = DefinedName("DP_ExcludedStatus", attr_text=f"{dp}!$B$3")
    wb.defined_names["DP_TopN"] = DefinedName("DP_TopN", attr_text=f"{dp}!$B$4")
    wb.defined_names["DP_ListThrough"] = DefinedName("DP_ListThrough", attr_text=f"{dp}!$B$5")
    for c, h in ((1, "Segment"), (2, "BU label")):
        ws.cell(6, c).value = h; style(ws.cell(6, c), hdr=True)
    for i, (s, lab) in enumerate(segs):
        for c, v in ((1, s), (2, lab)):
            ws.cell(7 + i, c).value = v; style(ws.cell(7 + i, c), BLUE)
    hdr = 8 + n
    heads = ["Segment", "Description", "Total tickets", "Priority breakdown",
             "Highest priority associated tickets with owners (groomed only)", "Recent notes/updates", "Tickets listed"]
    for c, h in enumerate(heads, 1):
        ws.cell(hdr, c).value = h; style(ws.cell(hdr, c), hdr=True)
    M = '(Src_OG=1)*ISNUMBER(SEARCH("; "&{lab}&";","; "&Src_Segment&";"))'
    CUT = "IFERROR(MATCH(DP_ListThrough,Src_PrioList,0),0)"
    for i in range(n):
        r, lr = hdr + 1 + i, 7 + i
        lab = f"$B${lr}"
        m = M.format(lab=lab)
        ws.cell(r, 1).value = f"=A{lr}"
        ws.cell(r, 2).value = (f'="Project = "&DP_Project&CHAR(10)&"Resolution = Unresolved"&CHAR(10)&"Labels = "&{lab}'
                               f'&CHAR(10)&"Status != "&DP_ExcludedStatus')
        ws.cell(r, 3).value = f'=SUMPRODUCT({m})&" total tickets"'
        parts = [f'INDEX(Src_PrioList,{k})&": "&SUMPRODUCT({m}*(Src_Rank={k}))' for k in range(1, len(pm) + 1)]
        nop = f"SUMPRODUCT({m}*(Src_Rank=99))"
        ws.cell(r, 4).value = "=" + '&CHAR(10)&'.join(parts) + f'&IF({nop}>0,CHAR(10)&"No/unmapped priority: "&{nop},"")'
        # tickets listed = every ticket through the 'always list' priority, or the top-N if that is more
        ws.cell(r, 7).value = f"=MAX(DP_TopN,SUMPRODUCT({m}*(Src_Rank<={CUT})))"
        terms = [f'IF({j}>$G{r},"",IFERROR(INDEX(Src_Line,MATCH(SMALL(IF({m},Src_Key),{j}),IF({m},Src_Key),0)),""))'
                 for j in range(1, MAXN + 1)]
        ref = f"E{r}"
        ws[ref] = ArrayFormula(ref, '=_xlfn.TEXTJOIN(CHAR(10),TRUE,' + ",".join(terms) + ")")
        ws.cell(r, 6).value = st["notes"].get(segs[i][0])
        for c in range(1, 8):
            style(ws.cell(r, c), BLUE if c == 6 else BLACK)
        ns = exp[segs[i][0]]["n_show"]
        ws.row_dimensions[r].height = min(409, 15 * max(5, 1.5 * ns) + 6)
    # No default manual rows. Rows the user added below the table are kept only if they hold typed content.
    manual = [k for k, vals in st["manual"].items() if any(v not in (None, "") for v in vals)]
    for k, mname in enumerate(manual):
        r = hdr + 1 + n + k
        ws.cell(r, 1).value = mname; style(ws.cell(r, 1), BLUE)
        vals = st["manual"][mname]
        for c in range(2, 7):
            ws.cell(r, c).value = vals[c - 2]; style(ws.cell(r, c), BLUE)
        ws.row_dimensions[r].height = 36
    st["kept_manual"] = manual
    st["dropped_manual"] = [k for k in st["manual"] if k not in manual]
    for col, w in {"A": 30, "B": 38, "C": 20, "D": 22, "E": 95, "F": 55, "G": 14}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = None
    return ws, segs, hdr, N


# ---------- independent Python computation (used for the report and to cross-check formulas) ----------
def clean(v):
    return " ".join(str(v or "").replace(";", "").replace(",", "").split())


def parse_created(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v
    m = re.match(r"^\s*(\d+)\.(\d+)\.(\d{4}) (\d+):(\d{2}) ?([AaPp][Mm])\s*$", str(v))
    if not m:
        return None
    d, mo, y, h, mi, ap = m.groups()
    h = int(h) % 12 + (12 if ap.upper() == "PM" else 0)
    return datetime(int(y), int(mo), int(d), h, int(mi))


def expected(ws, last, params, segs):
    pm = CFG["priority_map"]
    pref = CFG["segment_label_prefix"]
    rows = []
    for r in range(2, last + 1):
        v = lambda c: ws[f"{c}{r}"].value
        labels = [clean(v(c)) for c in "DEFGHI"]
        toks = [t for t in labels if t and t.lower().startswith(pref.lower())]
        prio = v("K")
        ps = "" if prio in (None, "") else next((b for a, b in pm if a.lower() == str(prio).lower()), "UNMAPPED")
        rank = 99 if ps in ("", "UNMAPPED") else [b for a, b in pm].index(ps) + 1
        key = str(v("M") or "")
        num = key.split("-", 1)[1] if "-" in key else key
        dt = parse_created(v("C"))
        mins = round((dt - EXCEL_EPOCH).total_seconds() / 60) if dt else 0
        og = int(key.lower().startswith(str(params["project_prefix"]).lower() + "-")
                 and v("Q") in (None, "") and str(v("J") or "").lower() != str(params["excluded_status"]).lower())
        line = f"{ps or 'No priority'}: #{num} – {v('P') or ''} – {v('A') or 'unassigned'}"
        rows.append(dict(r=r, key=key, labels=labels, toks=toks, seg="; ".join(toks), ps=ps, rank=rank, dt=dt,
                         mins=mins, og=og, line=line, status=v("J"), prio=prio, assignee=v("A")))
    out = {}
    pnames = [b for a, b in pm]
    lt = str(params.get(PRIO_KEY) or "").strip().lower()
    cut = next((k for k, nm in enumerate(pnames, 1) if nm.lower() == lt), 0)
    for s, lab in segs:
        match = [x for x in rows if x["og"] == 1 and lab.lower() in [t.lower() for t in x["toks"]]]
        match.sort(key=lambda x: (x["rank"], -x["mins"], x["r"]))
        n_show = max(int(params["tickets_per_segment"]), sum(1 for x in match if cut and x["rank"] <= cut))
        brk = [f"{nm}: {sum(1 for x in match if x['rank'] == k)}" for k, nm in enumerate(pnames, 1)]
        nop = sum(1 for x in match if x["rank"] == 99)
        if nop:
            brk.append(f"No/unmapped priority: {nop}")
        out[s] = dict(label=lab, total=len(match), n_show=n_show, breakdown="\n".join(brk),
                      by_priority={nm: sum(1 for x in match if x["rank"] == k) for k, nm in enumerate(pnames, 1)},
                      top=[x["line"] for x in match[:n_show]])
    return rows, out


def flags(rows, segs, params):
    f = {}
    labels_known = {lab.lower() for _, lab in segs}
    og = [x for x in rows if x["og"] == 1]
    f["multi_bu_label (all rows)"] = [(x["key"], x["seg"], x["og"]) for x in rows if len({t.lower() for t in x["toks"]}) > 1]
    f["duplicate_label_same_ticket"] = [x["key"] for x in rows if len(x["toks"]) != len({t.lower() for t in x["toks"]})]
    f["priority_not_mapped"] = sorted({str(x["prio"]) for x in rows if x["ps"] == "UNMAPPED"})
    f["priority_blank_open_groomed"] = sum(1 for x in og if x["ps"] == "")
    f["bu_labels_not_in_segment_lookup"] = {}
    for x in rows:
        for t in x["toks"]:
            if t.lower() not in labels_known:
                d = f["bu_labels_not_in_segment_lookup"].setdefault(t, [0, 0])
                d[0] += 1; d[1] += x["og"]
    f["bu_label_wrong_case"] = sorted({t for x in rows for t in x["toks"] if not t.startswith(CFG["segment_label_prefix"])})
    f["open_groomed_without_bu_label"] = sum(1 for x in og if not x["toks"])
    f["rows_not_matching_project_prefix"] = sum(1 for x in rows if not x["key"].lower().startswith(str(params["project_prefix"]).lower() + "-"))
    f["created_unparseable"] = [x["key"] for x in rows if x["dt"] is None]
    ex = str(params["excluded_status"])
    sts = {}
    for x in rows:
        s = x["status"]
        sts[s] = sts.get(s, 0) + 1
    f["excluded_status_rows"] = sts.get(ex, 0)
    f["excluded_status_near_misses"] = [s for s in sts if s and s != ex and s.strip().lower() == ex.strip().lower()]
    f["status_with_stray_whitespace"] = [s for s in sts if s and s != s.strip()]
    f["status_values_among_open_groomed"] = {s: sum(1 for x in og if x["status"] == s) for s in sorted({x["status"] for x in og if x["status"]})}
    return f


def recalc_check(out_path, ws_name, dp_name, rows, exp, segs, hdr):
    """Recalculate a COPY with LibreOffice and compare every formula result with the Python expectation."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return ["LibreOffice not found: formulas were NOT recalculated or cross-checked"]
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "in.xlsx"; shutil.copy(out_path, src)
        outd = Path(td) / "out"; outd.mkdir()
        p = subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(outd), str(src)],
                           capture_output=True, text=True, timeout=300)
        res = outd / "in.xlsx"
        if not res.exists():
            return [f"LibreOffice recalculation failed: {p.stderr[-300:]}"]
        wb = load_workbook(res, data_only=True)
    ws, dp = wb[ws_name], wb[dp_name]
    bad = []
    for x in rows:
        r = x["r"]
        got = {"ps": ws[f"{C_PS}{r}"].value or "", "seg": ws[f"{C_SEG}{r}"].value or "", "rank": ws[f"{C_RANK}{r}"].value,
               "line": ws[f"{C_LINE}{r}"].value, "og": ws[f"{C_OG}{r}"].value, "cre": ws[f"{C_CRE}{r}"].value}
        for k in ("ps", "seg", "rank", "line", "og"):
            if got[k] != x[k]:
                bad.append(f"row {r} {k}: sheet={got[k]!r} expected={x[k]!r}")
        if x["dt"] is None:
            if got["cre"] not in (None, ""):
                bad.append(f"row {r} created: sheet={got['cre']!r} expected blank")
        else:
            gc = got["cre"]
            if isinstance(gc, datetime):
                ok = abs((gc - x["dt"]).total_seconds()) < 1
            else:
                ok = isinstance(gc, (int, float)) and abs(gc - (x["dt"] - EXCEL_EPOCH).total_seconds() / 86400) < 1e-6
            if not ok:
                bad.append(f"row {r} created: sheet={gc!r} expected={x['dt']}")
    for i, (s, lab) in enumerate(segs):
        r = hdr + 1 + i
        tot, brk, top = dp.cell(r, 3).value, dp.cell(r, 4).value or "", dp.cell(r, 5).value or ""
        e = exp[s]
        if tot != f"{e['total']} total tickets":
            bad.append(f"{s} total: sheet={tot!r} expected={e['total']}")
        if brk != e["breakdown"]:
            bad.append(f"{s} priority breakdown: sheet={brk!r} expected={e['breakdown']!r}")
        if dp.cell(r, 7).value != e["n_show"]:
            bad.append(f"{s} tickets listed: sheet={dp.cell(r, 7).value!r} expected={e['n_show']}")
        if top != "\n".join(e["top"]):
            bad.append(f"{s} top list differs:\n  sheet={top!r}\n  expected={chr(10).join(e['top'])!r}")
    return bad


def slide_html(path, segs, exp, notes, params):
    """Standalone HTML with one inline-styled table, for copy/paste into PowerPoint."""
    from html import escape
    th = "background:#D9D9D9;font-weight:bold;border:1px solid #808080;padding:4px 6px;text-align:left;vertical-align:top;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000"
    td = "border:1px solid #808080;padding:4px 6px;text-align:left;vertical-align:top;font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#000000"
    def cell(text, style_):
        return f'<td style="{style_}">' + "<br>".join(escape(t) for t in str(text or "").split("\n")) + "</td>"
    heads = ["Segment", "Description", "Total tickets", "Priority breakdown", "Highest priority associated tickets with owners (groomed only)", "Recent notes/updates"]
    rows = ["<tr>" + "".join(f'<th style="{th}">{escape(h)}</th>' for h in heads) + "</tr>"]
    for s, lab in segs:
        e = exp[s]
        desc = (f"Project = {params['project_prefix']}\nResolution = Unresolved\nLabels = {lab}\nStatus != {params['excluded_status']}")
        rows.append("<tr>" + cell(s, td + ";font-weight:bold") + cell(desc, td) + cell(f"{e['total']} total tickets", td)
                    + cell(e["breakdown"], td) + cell("\n".join(e["top"]), td) + cell(notes.get(s), td) + "</tr>")
    doc = ('<!doctype html><html><head><meta charset="utf-8"><title>Dev Priorities - slide table</title></head>'
           '<body style="margin:16px;background:#ffffff"><p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555555">'
           'Select the table, copy, and paste into PowerPoint (Keep Source Formatting keeps the grey header).</p>'
           '<table style="border-collapse:collapse;width:100%">' + "".join(rows) + "</table></body></html>")
    Path(path).write_text(doc, encoding="utf-8")


# ---------- CSV input: open the export as a workbook in the layout the rest of the script expects ----------
CANON = ["Assignee", "Description", "Created", "Labels", "Labels", "Labels", "Labels", "Labels", "Labels",
         "Status", "Priority", "Custom field (Case ID)", "Issue key", "Issue id", "Parent id", "Summary", "Resolution"]
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def csv_to_workbook(csv_path, xlsx_path):
    """Write a CSV Jira export to an .xlsx in the canonical A-Q layout (sheet "Backlog") and return a report dict.

    Columns are found by header name, so extra or missing optional columns do not matter. A Jira export has one
    "Labels" column per label slot (the count varies by export; the layout here has six), or one comma-joined column:
    all labels are collected, BU labels first, and the first six are kept (any dropped label is reported). A Resolution
    of "Unresolved" is written as an empty cell, the way Jira leaves it for open tickets. Nothing else is changed.
    """
    from openpyxl import Workbook
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        data = list(csv.reader(fh))
    if not data:
        sys.exit("ERROR: the CSV is empty; nothing was changed.")
    hdr = [h.strip() for h in data[0]]
    low = [h.lower() for h in hdr]
    def col(*names):
        for nm in names:
            if nm.lower() in low:
                return low.index(nm.lower())
        return None
    idx = {"Assignee": col("Assignee"), "Description": col("Description"), "Created": col("Created"),
           "Status": col("Status"), "Priority": col("Priority"), "Custom field (Case ID)": col("Custom field (Case ID)", "Case ID"),
           "Issue key": col("Issue key", "Key"), "Issue id": col("Issue id"), "Parent id": col("Parent id"),
           "Summary": col("Summary"), "Resolution": col("Resolution")}
    lab_cols = [i for i, h in enumerate(low) if h == "labels"]
    required = ["Assignee", "Created", "Status", "Priority", "Issue key", "Summary", "Resolution"]
    missing = [r for r in required if idx[r] is None] + (["Labels"] if not lab_cols else [])
    if missing:
        sys.exit("ERROR: the CSV is missing required columns " + ", ".join(missing) + "; nothing was changed.\n  Found: " + ", ".join(hdr))
    pref = CFG["segment_label_prefix"].lower()
    wb = Workbook(); ws = wb.active; ws.title = "Backlog"
    ws.append(CANON)
    rep = dict(csv_rows=len(data) - 1, label_columns_found=len(lab_cols), labels_dropped=[], unresolved_blanked=0,
               descriptions_truncated=0, short_rows_padded=0)
    for n, row in enumerate(data[1:], 2):
        if not any(c.strip() for c in row):
            continue
        if len(row) < len(hdr):
            row = row + [""] * (len(hdr) - len(row)); rep["short_rows_padded"] += 1
        g = lambda name: CTRL.sub("", row[idx[name]]) if idx[name] is not None else ""
        labels = [t.strip() for i in lab_cols for t in CTRL.sub("", row[i]).split(",") if t.strip()]
        labels = [t for t in labels if t.lower().startswith(pref)] + [t for t in labels if not t.lower().startswith(pref)]
        if len(labels) > 6:
            rep["labels_dropped"].append((g("Issue key"), labels[6:]))
            labels = labels[:6]
        labels += [""] * (6 - len(labels))
        res = g("Resolution")
        if res.strip().lower() == "unresolved":
            res = ""; rep["unresolved_blanked"] += 1
        desc = g("Description")
        if len(desc) > 32000:
            desc = desc[:32000]; rep["descriptions_truncated"] += 1
        vals = [g("Assignee"), desc, g("Created")] + labels + [g("Status"), g("Priority"), g("Custom field (Case ID)"),
                g("Issue key"), g("Issue id"), g("Parent id"), g("Summary"), res]
        ws.append([v if v != "" else None for v in vals])
    wb.save(xlsx_path)
    rep["rows_written"] = ws.max_row - 1
    return rep


# ---------- PowerPoint output: the same summary table on one or more slides ----------
def build_slides(path, template, segs, exp, notes, params, n_rows, no_bu_open):
    """Write the Dev Priorities summary (Segment, Total, Priority breakdown, ticket lines, notes) to a .pptx.

    Segments are packed onto as few slides as fit (usually one or two). With a template, its master, layout and theme
    are used (all template slides are dropped); without one, python-pptx's default 16:9 master is used. The values come
    from the independent Python calculation that the workbook formulas are cross-checked against.
    """
    import math
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from lxml import etree
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    sc = CFG["slide"]; col = sc["colors"]
    prs = Presentation(template) if template else Presentation()
    if not template:
        prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    for sid in list(prs.slides._sldIdLst):  # drop the template's own slides completely (relationship + part)
        prs.part.drop_rel(sid.rId); prs.slides._sldIdLst.remove(sid)
    layouts = [l for m in prs.slide_masters for l in m.slide_layouts]
    layout = next((l for l in layouts if l.name == sc["layout"]), None)
    if layout is None and not template:
        layout = prs.slide_layouts[5]
    if layout is None:
        sys.exit(f"ERROR: layout '{sc['layout']}' not found in the template; layouts: {sorted({l.name for l in layouts})}")
    W_ = prs.slide_width
    L, TOP, W = 338328, 1000813, W_ - 2 * 338328
    NOTE_TOP = prs.slide_height - 1000000
    avail = NOTE_TOP - TOP - 120000 - 300000        # height for body rows under the header row
    has_notes = any(str(notes.get(s) or "").strip() for s, _ in segs)
    w_seg, w_tot, w_brk, w_notes = 1150000, 800000, 1500000, (2300000 if has_notes else 0)
    w_tix = W - w_seg - w_tot - w_brk - w_notes
    cpl = int(w_tix / 914400 * 17)                   # ~characters per line at 8 pt
    LH, PAD = 122000, 70000
    def lines_of(e):
        return max(sum(max(1, math.ceil(len(t) / cpl)) for t in e["top"]) if e["top"] else 1, len(e["breakdown"].split("\n")))
    pages, cur, used = [], [], 0
    for s, lab in segs:
        h = lines_of(exp[s]) * LH + PAD
        if cur and used + h > avail:
            pages.append(cur); cur, used = [], 0
        cur.append((s, lab, h)); used += h
    pages.append(cur)

    def fixp(p):
        pPr = p._p.get_or_add_pPr(); pPr.set("marL", "0"); pPr.set("indent", "0")
        for ch in list(pPr): pPr.remove(ch)
        etree.SubElement(etree.SubElement(pPr, A + "lnSpc"), A + "spcPct", val="100000")
        for tg in ("spcBef", "spcAft"):
            etree.SubElement(etree.SubElement(pPr, A + tg), A + "spcPts", val="0")
        etree.SubElement(pPr, A + "buNone")
    def cell(c, lines, size, bold, color, align, fill):
        c.margin_left = c.margin_right = Emu(54864); c.margin_top = c.margin_bottom = Emu(13716)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.text_frame.word_wrap = True
        for i, (tx, sz, b, colr, it) in enumerate(lines if isinstance(lines[0], tuple) else [(t, size, bold, color, False) for t in lines]):
            p = c.text_frame.paragraphs[0] if i == 0 else c.text_frame.add_paragraph()
            p.alignment = align
            r = p.add_run(); r.text = tx; r.font.size = Pt(sz); r.font.bold = b; r.font.italic = it
            r.font.color.rgb = RGBColor.from_string(colr)
            fixp(p)
        c.fill.solid(); c.fill.fore_color.rgb = RGBColor.from_string(fill)
        tcPr = c._tc.get_or_add_tcPr()
        for k, tag in enumerate(("lnL", "lnR", "lnT", "lnB")):
            ln = etree.Element(A + tag, w="12700", cap="flat", cmpd="sng", algn="ctr")
            etree.SubElement(etree.SubElement(ln, A + "solidFill"), A + "srgbClr", val="FFFFFF")
            etree.SubElement(ln, A + "prstDash", val="solid"); tcPr.insert(k, ln)

    pm = CFG["priority_map"]
    note = (f"Note: Open = {params['project_prefix']} tickets with Resolution unresolved and Status other than {params['excluded_status']}, "
            f"carrying the segment's BU label. Each list holds every ticket at {params[PRIO_KEY]} or higher, ranked by priority and then newest first, "
            f"topped up to {params['tickets_per_segment']} tickets; the breakdown counts all open tickets in the segment. "
            f"Tickets with several BU labels count in each segment. {no_bu_open} open tickets carry no BU label and are not shown. "
            "Ticket titles are quoted as written in Jira.")
    src_line = f"Source: Jira backlog export ({n_rows} issues)"
    for pi, page in enumerate(pages, 1):
        sl = prs.slides.add_slide(layout)
        t = sl.shapes.title; tf = t.text_frame; tf.clear(); p = tf.paragraphs[0]
        r = p.add_run(); r.text = sc["title_lead"] + " | "; r.font.bold = True; r.font.size = Pt(28)
        tail = sc["title_tail"] + (f" ({pi} of {len(pages)})" if len(pages) > 1 else "")
        r = p.add_run(); r.text = tail; r.font.size = Pt(28)
        t.left, t.top, t.width, t.height = Emu(330200), Emu(476191), Emu(W_ - 660400), Emu(400110)
        heads = ["Segment", "Total tickets", "Priority breakdown", "Highest priority associated tickets with owners (groomed only)"]
        widths = [w_seg, w_tot, w_brk, w_tix]
        if has_notes:
            heads.append("Recent notes/updates"); widths.append(w_notes)
        H = 300000 + sum(h for _, _, h in page)
        gf = sl.shapes.add_table(len(page) + 1, len(heads), Emu(L), Emu(TOP), Emu(W), Emu(H)); gf.name = "Priority tickets table"
        tb = gf.table
        for a_ in ("firstRow", "bandRow"): tb._tbl.tblPr.attrib.pop(a_, None)
        for i, w in enumerate(widths): tb.columns[i].width = Emu(w)
        tb.rows[0].height = Emu(300000)
        for j, h in enumerate(heads):
            cell(tb.cell(0, j), [h], 10, True, "FFFFFF", PP_ALIGN.CENTER if j in (1, 2) else PP_ALIGN.LEFT, col["header_fill"])
        for i, (s, lab, h) in enumerate(page, 1):
            e = exp[s]; tb.rows[i].height = Emu(h)
            cell(tb.cell(i, 0), [(s, 9, True, col["text"], False), (lab, 7, False, "595959", True)], 9, True, col["text"], PP_ALIGN.LEFT, col["label_fill"])
            cell(tb.cell(i, 1), [str(e["total"])], 9, False, col["text"], PP_ALIGN.CENTER, col["body_fill"])
            cell(tb.cell(i, 2), e["breakdown"].split("\n"), 8, False, col["text"], PP_ALIGN.LEFT, col["body_fill"])
            cell(tb.cell(i, 3), e["top"] or ["(none)"], 8, False, col["text"], PP_ALIGN.LEFT, col["body_fill"])
            if has_notes:
                cell(tb.cell(i, 4), [str(notes.get(s) or "")] , 8, False, col["text"], PP_ALIGN.LEFT, col["body_fill"])
        nb = sl.shapes.add_textbox(Emu(L), Emu(NOTE_TOP), Emu(W), Emu(492443)); nb.name = "Note and source"
        ntf = nb.text_frame; ntf.word_wrap = True
        ntf.margin_left = ntf.margin_right = ntf.margin_top = ntf.margin_bottom = 0
        for i, x in enumerate((note, src_line)):
            p = ntf.paragraphs[0] if i == 0 else ntf.add_paragraph()
            r = p.add_run(); r.text = x; r.font.size = Pt(8); r.font.bold = False; r.font.color.rgb = RGBColor.from_string(col["text"])
            fixp(p)
    prs.save(path)
    return len(pages)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook"); ap.add_argument("--out"); ap.add_argument("--in-place", action="store_true")
    ap.add_argument("--sheet"); ap.add_argument("--no-verify", action="store_true"); ap.add_argument("--json", action="store_true")
    ap.add_argument("--template", help="PowerPoint deck whose master/layout/theme style the slides (its own slides are not copied)")
    a = ap.parse_args()
    src = Path(a.workbook)
    conversion = None
    if src.suffix.lower() == ".csv":
        if a.in_place:
            sys.exit("ERROR: --in-place cannot be used with a CSV; the workbook is written to a new file.")
        out = Path(a.out) if a.out else src.with_name(src.stem + "_priorities.xlsx")
        tmp = Path(tempfile.mkdtemp()) / (src.stem + ".xlsx")
        conversion = csv_to_workbook(src, tmp)
        wb = load_workbook(tmp)
    else:
        out = src if a.in_place else Path(a.out) if a.out else src.with_name(src.stem + "_priorities.xlsx")
        wb = load_workbook(src)
    ws, bad = find_source(wb, a.sheet)
    if ws is None:
        sys.exit("ERROR: source columns do not match the expected layout; nothing was changed:\n  " + "\n  ".join(bad))
    last = last_data_row(ws)
    # refuse to overwrite anything in the new-column area that is not ours
    if str(ws[f"{C_PS}1"].value or "") == "Priority standardized":
        for col in NEW_COLS:
            for r in range(1, ws.max_row + 1):
                c = ws[f"{col}{r}"]; c.value = None; c.style = "Normal"
    else:
        for col in NEW_COLS:
            for r in range(1, ws.max_row + 1):
                if ws[f"{col}{r}"].value not in (None, ""):
                    sys.exit(f"ERROR: {col}{r} is not empty and is not from a previous run; nothing was changed.")
    before = [[ws.cell(r, c).value for c in range(1, 18)] for r in range(1, ws.max_row + 1)]
    st = read_existing(wb)
    dpq = q(CFG["dev_priorities_sheet"])
    params = {**CFG["parameters"], **st["params"]}
    if int(params["tickets_per_segment"]) > MAXN:
        sys.exit(f"ERROR: tickets per segment is {params['tickets_per_segment']}; formulas support at most {MAXN}.")
    segs_in = st["segments"] or CFG["segments"]
    rows, exp = expected(ws, last, params, segs_in)
    big = {s_: e["n_show"] for s_, e in exp.items() if e["n_show"] > MAXN}
    if big:
        sys.exit(f"ERROR: {big} tickets would be listed but the formulas list at most {MAXN} per segment; nothing was changed.")
    dws, segs, hdr, N = write_dev(wb, st, exp)
    write_source(ws, last, wb, dpq)
    after = [[ws.cell(r, c).value for c in range(1, 18)] for r in range(1, len(before) + 1)]
    assert before == after, "source data changed unexpectedly"
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    fl = flags(rows, segs, params)
    check = ["skipped (--no-verify)"] if a.no_verify else recalc_check(out, ws.title, dws.title, rows, exp, segs, hdr)
    slide = out.with_name(out.stem + "_slide_table.html")
    slide_html(slide, segs, exp, {k: v for k, v in st["notes"].items() if isinstance(k, str)}, params)
    pptx = out.with_name(out.stem + "_slides.pptx")
    notes_ = {k: v for k, v in st["notes"].items() if isinstance(k, str)}
    n_slides = build_slides(pptx, a.template, segs, exp, notes_, params, last - 1, fl["open_groomed_without_bu_label"])
    result = dict(output=str(out), slide_table=str(slide), slides=str(pptx), slide_count=n_slides, manual_rows_kept=st.get("kept_manual", []),
                  manual_rows_dropped_empty=st.get("dropped_manual", []), source_sheet=ws.title, last_data_row=last, parameters=params,
                  segment_counts={s: e["total"] for s, e in exp.items()},
                  priority_breakdown={s: e["by_priority"] for s, e in exp.items()},
                  tickets_listed={s: e["n_show"] for s, e in exp.items()}, csv_conversion=conversion, flags=fl, formula_check=check)
    if a.json:
        print(json.dumps(result, indent=2, default=str)); return
    print(f"Output: {out}\nSlides (.pptx, {n_slides}): {pptx}\nSlide table (HTML, copy/paste into PowerPoint): {slide}\nSource sheet: {ws.title} (rows 2-{last}); source columns A-Q confirmed and unchanged")
    if conversion:
        print(f"CSV opened as a workbook: {conversion['rows_written']} rows, {conversion['label_columns_found']} Labels column(s) found; "
              f"{conversion['unresolved_blanked']} 'Unresolved' resolutions written as blank; "
              f"labels dropped beyond 6 per ticket: {conversion['labels_dropped'] or 'none'}")
    print("Tickets per segment (Open & groomed), with priority breakdown and tickets listed:")
    for s, e in exp.items():
        bp = " | ".join(f"{k}: {v}" for k, v in e["by_priority"].items())
        print(f"  {s} ({e['label']}): {e['total']}  [{bp}]  listed: {e['n_show']}")
    if st.get("kept_manual"):
        print(f"Extra rows you typed below the table were kept: {st['kept_manual']}")
    if st.get("dropped_manual"):
        print(f"Empty extra rows removed from the tab: {st['dropped_manual']}")
    print("Flags:")
    for k, v in fl.items():
        print(f"  {k}: {v}")
    ok = check == []
    print("Formula cross-check (LibreOffice recalculation vs independent Python):",
          "PASS" if ok else check if len(check) < 15 else check[:15] + [f"... {len(check)} total"])


if __name__ == "__main__":
    main()
