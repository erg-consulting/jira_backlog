#!/usr/bin/env python3
"""Build the Dev Priorities summary in a Jira backlog workbook (live formulas only).

Usage: build_priority_table.py <workbook.xlsx> [--out OUT.xlsx | --in-place] [--sheet NAME] [--no-verify] [--json]
Default output: <workbook>_priorities.xlsx (the input is never modified unless --in-place).
"""
import argparse, json, re, shutil, subprocess, sys, tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

CFG = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text())
MAXN = 10  # max tickets per segment the formula supports
BLUE, BLACK = Font(color="0000FF"), Font(color="000000")
HDR_FONT, HDR_FILL = Font(bold=True, color="000000"), PatternFill("solid", fgColor="D9D9D9")
WRAP = Alignment(wrap_text=True, vertical="top")
EXCEL_EPOCH = datetime(1899, 12, 30)

# new source-sheet columns
C_PS, C_SEG = "R", "S"            # after Resolution (Q)
C_LK1, C_LK2 = "U", "V"           # lookup table (T is empty)
C_CRE, C_RANK, C_LINE, C_OG = "X", "Y", "Z", "AA"  # helpers (W is empty)
NEW_COLS = [C_PS, C_SEG, "T", C_LK1, C_LK2, "W", C_CRE, C_RANK, C_LINE, C_OG]
OURS = {"Priority standardized", "Segment", "Created (date)", "Priority rank", "Ticket line", "Open & groomed"}


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
    keys = {"Project prefix": "project_prefix", "Excluded status": "excluded_status", "Tickets listed per segment": "tickets_per_segment"}
    for r in range(2, 5):
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
        for r in range(hdr + 1, ws.max_row + 1):
            n = ws.cell(r, 1).value
            if n in (None, ""):
                continue
            row = [ws.cell(r, c).value for c in range(1, 6)]
            if str(n).startswith("="):  # summary row for a lookup segment
                st["notes"][r] = row[4]
                st.setdefault("summary_rows", []).append(r)
            elif n in names:
                st["notes"][n] = row[4]
            else:
                st["manual"][n] = row[1:5]
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
             C_CRE: "Created (date)", C_RANK: "Priority rank", C_LINE: "Ticket line", C_OG: "Open & groomed"}
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
        for col, formula in f.items():
            c = ws[f"{col}{r}"]; c.value = formula; style(c)
        ws[f"{C_CRE}{r}"].number_format = "dd.mm.yyyy hh:mm"
    widths = {C_PS: 20, C_SEG: 24, "T": 3, C_LK1: 22, C_LK2: 22, "W": 3, C_CRE: 18, C_RANK: 13, C_LINE: 70, C_OG: 15}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    sn = q(ws.title)
    names = {"Src_Segment": C_SEG, "Src_Created": C_CRE, "Src_Rank": C_RANK, "Src_Line": C_LINE, "Src_OG": C_OG}
    for n, col in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=f"{sn}!${col}${lo}:${col}${hi}")


def write_dev(wb, st, n_last_src):
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
    for c, h in ((1, "Parameter"), (2, "Value")):
        ws.cell(1, c).value = h; style(ws.cell(1, c), hdr=True)
    for r, (lab, key) in enumerate([("Project prefix", "project_prefix"), ("Excluded status", "excluded_status"),
                                    ("Tickets listed per segment", "tickets_per_segment")], 2):
        ws.cell(r, 1).value = lab; style(ws.cell(r, 1))
        ws.cell(r, 2).value = P[key]; style(ws.cell(r, 2), BLUE)
    dp = q(name)
    wb.defined_names["DP_Project"] = DefinedName("DP_Project", attr_text=f"{dp}!$B$2")
    wb.defined_names["DP_ExcludedStatus"] = DefinedName("DP_ExcludedStatus", attr_text=f"{dp}!$B$3")
    wb.defined_names["DP_TopN"] = DefinedName("DP_TopN", attr_text=f"{dp}!$B$4")
    for c, h in ((1, "Segment"), (2, "BU label")):
        ws.cell(6, c).value = h; style(ws.cell(6, c), hdr=True)
    for i, (s, lab) in enumerate(segs):
        for c, v in ((1, s), (2, lab)):
            ws.cell(7 + i, c).value = v; style(ws.cell(7 + i, c), BLUE)
    hdr = 8 + n
    heads = ["Segment", "Description", "Total tickets",
             "Highest priority associated tickets with owners (groomed only)", "Recent notes/updates"]
    for c, h in enumerate(heads, 1):
        ws.cell(hdr, c).value = h; style(ws.cell(hdr, c), hdr=True)
    M = '(Src_OG=1)*ISNUMBER(SEARCH("; "&{lab}&";","; "&Src_Segment&";"))'
    K = "Src_Rank*100000000+(100000000-IFERROR(ROUND(Src_Created*1440,0),0))+ROW(Src_Rank)/10000"
    height = 15 * max(4, 2 * N) + 6
    for i in range(n):
        r, lr = hdr + 1 + i, 7 + i
        lab = f"$B${lr}"
        ws.cell(r, 1).value = f"=A{lr}"
        ws.cell(r, 2).value = (f'="Project = "&DP_Project&CHAR(10)&"Resolution = Unresolved"&CHAR(10)&"Labels = "&{lab}'
                               f'&CHAR(10)&"Status != "&DP_ExcludedStatus')
        ws.cell(r, 3).value = f'=SUMPRODUCT({M.format(lab=lab)})&" total tickets"'
        m = M.format(lab=lab)
        terms = [f'IF({j}>DP_TopN,"",IFERROR(INDEX(Src_Line,MATCH(SMALL(IF({m},{K}),{j}),IF({m},{K}),0)),""))'
                 for j in range(1, MAXN + 1)]
        ref = f"D{r}"
        ws[ref] = ArrayFormula(ref, '=_xlfn.TEXTJOIN(CHAR(10),TRUE,' + ",".join(terms) + ")")
        ws.cell(r, 5).value = st["notes"].get(segs[i][0])
        for c in range(1, 6):
            style(ws.cell(r, c), BLUE if c == 5 else BLACK)
        ws.row_dimensions[r].height = height
    # No default manual rows. Rows the user added below the table are kept only if they hold typed content.
    manual = [k for k, vals in st["manual"].items() if any(v not in (None, "") for v in vals)]
    for k, mname in enumerate(manual):
        r = hdr + 1 + n + k
        ws.cell(r, 1).value = mname; style(ws.cell(r, 1), BLUE)
        vals = st["manual"][mname]
        for c in range(2, 6):
            ws.cell(r, c).value = vals[c - 2]; style(ws.cell(r, c), BLUE)
        ws.row_dimensions[r].height = 36
    st["kept_manual"] = manual
    st["dropped_manual"] = [k for k in st["manual"] if k not in manual]
    for col, w in {"A": 30, "B": 38, "C": 20, "D": 95, "E": 55}.items():
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
    for s, lab in segs:
        match = [x for x in rows if x["og"] == 1 and lab.lower() in [t.lower() for t in x["toks"]]]
        match.sort(key=lambda x: (x["rank"], -x["mins"], x["r"]))
        out[s] = dict(label=lab, total=len(match), top=[x["line"] for x in match[:int(params["tickets_per_segment"])]])
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
        tot, top = dp.cell(r, 3).value, dp.cell(r, 4).value or ""
        e = exp[s]
        if tot != f"{e['total']} total tickets":
            bad.append(f"{s} total: sheet={tot!r} expected={e['total']}")
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
    heads = ["Segment", "Description", "Total tickets", "Highest priority associated tickets with owners (groomed only)", "Recent notes/updates"]
    rows = ["<tr>" + "".join(f'<th style="{th}">{escape(h)}</th>' for h in heads) + "</tr>"]
    for s, lab in segs:
        e = exp[s]
        desc = (f"Project = {params['project_prefix']}\nResolution = Unresolved\nLabels = {lab}\nStatus != {params['excluded_status']}")
        rows.append("<tr>" + cell(s, td + ";font-weight:bold") + cell(desc, td) + cell(f"{e['total']} total tickets", td)
                    + cell("\n".join(e["top"]), td) + cell(notes.get(s), td) + "</tr>")
    doc = ('<!doctype html><html><head><meta charset="utf-8"><title>Dev Priorities - slide table</title></head>'
           '<body style="margin:16px;background:#ffffff"><p style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#555555">'
           'Select the table, copy, and paste into PowerPoint (Keep Source Formatting keeps the grey header).</p>'
           '<table style="border-collapse:collapse;width:100%">' + "".join(rows) + "</table></body></html>")
    Path(path).write_text(doc, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook"); ap.add_argument("--out"); ap.add_argument("--in-place", action="store_true")
    ap.add_argument("--sheet"); ap.add_argument("--no-verify", action="store_true"); ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    src = Path(a.workbook)
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
    dws, segs, hdr, N = write_dev(wb, st, last)
    write_source(ws, last, wb, dpq)
    after = [[ws.cell(r, c).value for c in range(1, 18)] for r in range(1, len(before) + 1)]
    assert before == after, "source data changed unexpectedly"
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    rows, exp = expected(ws, last, params, segs)
    fl = flags(rows, segs, params)
    check = ["skipped (--no-verify)"] if a.no_verify else recalc_check(out, ws.title, dws.title, rows, exp, segs, hdr)
    slide = out.with_name(out.stem + "_slide_table.html")
    slide_html(slide, segs, exp, {k: v for k, v in st["notes"].items() if isinstance(k, str)}, params)
    result = dict(output=str(out), slide_table=str(slide), manual_rows_kept=st.get("kept_manual", []),
                  manual_rows_dropped_empty=st.get("dropped_manual", []), source_sheet=ws.title, last_data_row=last, parameters=params,
                  segment_counts={s: e["total"] for s, e in exp.items()}, flags=fl, formula_check=check)
    if a.json:
        print(json.dumps(result, indent=2, default=str)); return
    print(f"Output: {out}\nSlide table (copy/paste into PowerPoint): {slide}\nSource sheet: {ws.title} (rows 2-{last}); source columns A-Q confirmed and unchanged")
    print("Tickets per segment (Open & groomed):")
    for s, e in exp.items():
        print(f"  {s} ({e['label']}): {e['total']}")
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
