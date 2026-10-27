"""Assemble ISKO_ERP_100K_Row_Resilience_Plan.md from plan_template.md, the
audit summary (measured facts) and the plan data modules.

Usage: python3 build_md.py [toc_pages.json]
"""
import json, os, re, sys
from collections import Counter, OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from data_sheets import S as SHEETS
from data_columns import BP, MAP, AUDIT
import data_misc as DM

AUD = json.load(open(os.path.join(HERE, "audit_summary.json")))
BOOKS = ["Stores", "Production", "Dispatch", "MIS"]
BOOKNAME = {"Stores": "1 Stores", "Production": "2 Production", "Dispatch": "3 Dispatch & Challan", "MIS": "4 MIS"}
FILESIZE = {"Stores": 2.79, "Production": 7.23, "Dispatch": 2.44, "MIS": 3.91}  # MB, measured from attachments


def esc(x):
    s = "" if x is None else str(x)
    return s.replace("|", "\\|").replace("\n", " ").strip()


def table(headers, rows, widths=None, cls=None):
    """Pipe table; dash counts set relative column widths in pandoc."""
    widths = widths or [10] * len(headers)
    out = ["| " + " | ".join(esc(h) for h in headers) + " |",
           "|" + "|".join("-" * max(3, w) for w in widths) + "|"]
    for r in rows:
        out.append("| " + " | ".join(esc(c) for c in r) + " |")
    t = "\n".join(out)
    if cls:
        t = f"::: {{.{cls}}}\n{t}\n:::"
    return t


def n(x):
    return f"{x:,}" if isinstance(x, (int, float)) else str(x)


sheets = AUD["sheets"]
tables = AUD["tables"]
tab_by_sheet = {}
for t in tables:
    tab_by_sheet.setdefault((t["wb"], t["sheet"]), []).append(t)

# --------------------------------------------------------------- inventory
def tbl_inventory():
    rows = []
    tot = Counter()
    for b in BOOKS:
        sh = [s for s in sheets if s["wb"] == b]
        vis = sum(1 for s in sh if s["state"] == "visible")
        f = sum(s["formulas"] for s in sh)
        nm = AUD["names"][b]
        tb = sum(1 for t in tables if t["wb"] == b)
        prot = sum(1 for s in sh if s["prot"]); frz = sum(1 for s in sh if s["freeze"])
        top = max(sh, key=lambda s: s["formulas"])
        rows.append([BOOKNAME[b], f"{len(sh)} ({vis} visible / {len(sh)-vis} hidden)", n(f), f"{FILESIZE[b]:.2f}", tb,
                     f"{nm['total']} ({nm['spill']} spill)", f"{prot} / {frz}", f"{top['sheet']} ({n(top['formulas'])})"])
        tot.update(dict(sh=len(sh), f=f, tb=tb, nm=nm["total"], sp=nm["spill"], prot=prot, frz=frz))
    rows.append(["**Total**", f"**{tot['sh']}**", f"**{n(tot['f'])}**", f"**{sum(FILESIZE.values()):.2f}**", f"**{tot['tb']}**",
                 f"**{tot['nm']} ({tot['sp']} spill)**", f"**{tot['prot']} / {tot['frz']}**", ""])
    return table(["Workbook", "Sheets", "Formula cells", "File MB", "Excel Tables", "Defined names", "Protected / freeze-pane sheets", "Largest formula sheet"],
                 rows, [16, 16, 10, 7, 7, 12, 12, 24])


def tbl_hotspots():
    allf = sum(s["formulas"] for s in sheets)
    top = sorted(sheets, key=lambda s: -s["formulas"])[:12]
    rows = []
    for s in top:
        t = tab_by_sheet.get((s["wb"], s["sheet"]), [])
        data = t[0]["data_rows"] if t else None
        body = t[0]["body_rows"] if t else None
        per = f"{s['formulas']/data:.1f}" if data else "-"
        rows.append([BOOKNAME[s["wb"]], s["sheet"], n(s["formulas"]), f"{100*s['formulas']/allf:.1f}%",
                     n(data) if data is not None else "-", n(body) if body is not None else "-", per, s["dim"]])
    return table(["Workbook", "Sheet", "Formula cells", "% of all", "Data rows", "Table rows (incl. pre-filled)", "Formulas per data row", "Used range"],
                 rows, [14, 22, 9, 7, 8, 12, 9, 11])


def tbl_patterns():
    rows = []
    keys = ["total", "post2016_fn", "dynamic_array", "running_range", "self_table_agg", "row_based_id", "volatile", "reads_import_table"]
    tot = Counter()
    for b in BOOKS:
        p = AUD["patterns"][b]["total"]
        rows.append([BOOKNAME[b]] + [n(p.get(k, 0)) for k in keys])
        tot.update({k: p.get(k, 0) for k in keys})
    rows.append(["**Total**"] + [f"**{n(tot[k])}**" for k in keys])
    return table(["Workbook", "Formula cells", "Post-2016 functions", "Dynamic-array functions", "Expanding ranges", "Self-table aggregates",
                  "Position-dependent (ROW())", "Volatile", "Reads Power Query tables"], rows, [14, 9, 9, 9, 9, 9, 10, 8, 10])


def tbl_growth():
    rows = []
    for t, cur, per, pm, py, yrs, f, fpr in DM.GROWTH:
        rows.append([t, n(cur), per, n(pm), n(py), f"{yrs:.1f}" if yrs < 50 else ">50", n(f), f"{fpr:.1f}"])
    return table(["Table", "Data rows now", "Period", "Rows / month", "Rows / year (est.)", "Years to 100,000 (est.)", "Formula cells now", "Formulas per row"],
                 rows, [20, 9, 10, 8, 9, 10, 9, 8])


def tbl_entity():
    return table(["Entity", "System of record", "Entered / imported", "Primary key", "Key foreign keys", "Maintained by", "Consumed by", "Type"],
                 DM.ENTITY, [16, 12, 12, 14, 18, 12, 18, 11])


DEC_ORDER = {"Keep": 0, "Simplify": 1, "Combine": 2, "Convert to report": 3, "Remove from Phase 1": 4, "Removed from Phase 1": 4}


def sheet_rows(book):
    rows = []
    for s in [x for x in sheets if x["wb"] == book]:
        d = SHEETS.get((book, s["sheet"]))
        if d is None:
            raise SystemExit(f"Missing decision for {book} / {s['sheet']}")
        t = tab_by_sheet.get((book, s["sheet"]), [])
        data = t[0]["data_rows"] if t else None
        size = f"{s['dim']}; {'%s rows; ' % n(data) if data is not None else ''}{n(s['formulas'])} f"
        hid = " (hidden)" if s["state"] != "visible" else ""
        dec = d["decision"].replace("Remove from Phase 1", "Removed from Phase 1")
        rows.append([s["sheet"] + hid, d["purpose"], size, f"**{dec}**", d["reason"], d["proposed"], d["target"],
                     d["io"], d["deps"], d["risk"]])
    return rows


def tbl_sheets():
    out = []
    for b in BOOKS:
        rows = sheet_rows(b)
        if b == "MIS":
            # Compact the 32 hidden import sheets into their own table
            main = [r for r in rows if not r[0].startswith("D_")]
            imp = [r for r in rows if r[0].startswith("D_")]
            out.append(f"#### {BOOKNAME[b]} - report sheets ({len(main)})\n")
            out.append(table(["Current sheet", "Apparent purpose", "Used range; rows; formulas", "Decision", "Reason and control impact", "Proposed final sheet / table",
                              "Target rows", "Inputs / outputs", "Dependencies", "Main risks / controls"], main, [12, 16, 10, 8, 22, 12, 7, 11, 12, 14]))
            out.append(f"\n#### {BOOKNAME[b]} - hidden Power Query import sheets ({len(imp)})\n")
            out.append(table(["Current sheet", "Apparent purpose", "Used range; rows; formulas", "Decision", "Reason", "Proposed", "Target rows", "Source"],
                             [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[8]] for r in imp], [12, 22, 12, 8, 22, 10, 8, 10]))
        else:
            out.append(f"#### {BOOKNAME[b]} ({len(rows)} sheets)\n")
            out.append(table(["Current sheet", "Apparent purpose", "Used range; rows; formulas", "Decision", "Reason and control impact", "Proposed final sheet / table",
                              "Target rows", "Inputs / outputs", "Dependencies", "Main risks / controls"], rows, [12, 16, 10, 8, 22, 12, 7, 11, 12, 14]))
        out.append("")
    # decision summary
    c = Counter()
    for (b, s), d in SHEETS.items():
        c[d["decision"].replace("Remove from Phase 1", "Removed from Phase 1")] += 1
    summ = ", ".join(f"{k}: {v}" for k, v in sorted(c.items(), key=lambda x: DEC_ORDER.get(x[0], 9)))
    return f"**Decision summary across all 120 sheets** - {summ}.\n\n" + "\n".join(out)


def tbl_audit():
    return table(["#", "Column", "Type / unit", "Class", "M/O", "Source or rule", "Validation", "Why needed"],
                 [[f"A{i+1}"] + list(r) for i, r in enumerate(AUDIT)], [3, 14, 12, 5, 4, 30, 18, 16])


MOCK_BEFORE = {"tblGRN": ("mock_grn.png", "Stores - GRN entry pad (sample data)"),
               "tblRMMove": ("mock_rmissue.png", "Stores - RM issue and return entry pad (sample data)"),
               "tblProd": ("mock_prod.png", "Production - production entry pad (sample data)"),
               "tblDCHead": ("mock_dc.png", "Dispatch - delivery challan document pad (sample data)")}


def tbl_blueprints():
    out = []
    groups = OrderedDict([("Stores", []), ("Production", []), ("Dispatch", []), ("All", []), ("MIS", [])])
    for k, v in BP.items():
        groups[v["book"]].append(k)
    sec = 0
    for g, keys in groups.items():
        for k in keys:
            v = BP[k]
            sec += 1
            if k in MOCK_BEFORE:
                img, cap = MOCK_BEFORE[k]
                out.append(f"![Illustrative Excel Layout - Mock-up: {cap}](assets/{img}){{.shot}}\n")
            aud = " + audit columns A1-A6" if v["audit"] else ""
            out.append(f"#### 9.{sec} {v['title']} - `{k}`\n")
            out.append(f"*Book:* {BOOKNAME.get(v['book'], v['book'])} · *Sheet:* {v['sheet']} · *Target volume:* {v['rows']}{aud}. {v['desc']}\n")
            rows = [[i + 1] + list(c) for i, c in enumerate(v["cols"])]
            out.append(table(["#", "Final column (in order)", "Type / unit", "Class", "M/O", "Source or calculation rule", "Validation (H / W)", "Why needed in Phase 1"],
                             rows, [3, 16, 11, 5, 4, 30, 20, 14]))
            out.append("")
    return "\n".join(out)


def tbl_scale():
    rows = [
        ["Formula cells in largest book", "217,775 at 12,538 rows", "<= 40,000 at 100,000 rows (pads + balances)"],
        ["Work per new row", "~18-25 formulas, several scanning the whole table", "0 in the register; pad row lookups only"],
        ["Aggregates per edit", "Per-row SUMIFS / expanding ranges over the table (grows with rows²)", "Balances for open items, rebuilt on posting"],
        ["Volatile triggers", "TODAY() in 5,571 cells; CELL(); 5-minute background refresh", "None"],
        ["IDs", "ROW() / running counts", "Values written at posting"],
        ["Imports", "Whole tables, all columns, blank rows included", "Needed columns, non-blank keys, lineage, logged"],
        ["History volume per file", "Unlimited growth", "One FY per file (~23k production rows/yr)"],
    ]
    return table(["Aspect", "Current design (measured)", "Target design"], rows, [18, 34, 34])


EXTRACTS = [
    ("X_Batch", "Stores tblRMMove (ISSUE lines + returns/scrap per batch)", "Production", "Batch ID, Move Date, Machine ID, Lot ID, Material ID, Grade, Thk, Size, Issued kg, Issued sheets, Returned kg, Stores scrap kg, Rate", "Start of shift; on demand"),
    ("X_Material", "Stores M_Material", "Production, MIS", "Material ID, Grade, Thk, Std L/W, Std sheet kg, Active", "Weekly / on change"),
    ("X_Product / X_Part / X_Machine / X_People", "Production masters", "Stores, Dispatch, MIS", "IDs, names, customer, route, Wt/pc, sq ft/pc, active", "Weekly / on change"),
    ("X_Customer", "Dispatch M_Customer", "Production, MIS", "Customer ID, name, group, Dispatch UOM, active", "Weekly / on change"),
    ("X_ProdOutput", "Production register + rejection events", "Dispatch", "Batch ID, Product ID, Part ID, Stage, QC-OK qty available to send, kg", "Start of shift; on demand"),
    ("X_GRN, X_LotQC, X_RMMove, X_Adjust, X_ConsMove, X_POStatus, X_LotBal, X_ConsBal", "Stores registers and balances", "MIS", "Register columns needed for KPIs and trace + lineage", "Before MIS review"),
    ("X_Prod, X_RejRework, X_BatchBal, X_WIP, X_BOM", "Production registers and balances", "MIS", "As above", "Before MIS review"),
    ("X_Requirement, X_JWCHead, X_JWCLine, X_CoaterMove, X_JWCBal, X_FGBal, X_DCHead, X_DCLine, X_CustPrice", "Dispatch registers and balances", "MIS", "As above", "Before MIS review"),
    ("Control Totals", "Each book (rows, kg, Rs per register and month)", "MIS", "Book, register, month, rows, kg, Rs", "Every save (formula-free: written at posting)"),
]


def tbl_extracts():
    return table(["Extract", "From (owner)", "To", "Columns that travel", "Refresh point"], EXTRACTS, [16, 22, 12, 36, 14])


def tbl_kpi():
    return table(["KPI", "Definition", "Source table", "Filters", "Action owner"], DM.KPIS, [18, 34, 16, 14, 12])


def tbl_checks():
    return table(["Code", "Exception", "Severity", "Owner"], DM.CHECKS, [6, 50, 8, 14])


def tbl_thresh():
    return table(["Operation", "Target at 10k-25k rows", "Target at 100k rows / 400k system", "How measured"], DM.THRESH, [26, 12, 14, 34])


def tbl_tests():
    return table(["ID", "Operation", "Volumes", "Steps", "Pass condition"], DM.TESTS, [5, 18, 18, 34, 22])


def tbl_risks():
    return table(["ID", "Risk", "Likelihood", "Impact", "Preventive control", "Detection", "Owner", "Fallback"], DM.RISKS, [5, 20, 7, 6, 24, 16, 10, 18])


def tbl_roadmap():
    return table(["Gate / phase", "When (proposed)", "Content", "Sign-off"], DM.ROADMAP, [20, 14, 44, 16])


def tbl_decisions():
    return table(["ID", "Decision required", "Why it matters", "Area"], DM.DECISIONS, [5, 50, 30, 8])


def tbl_reqs():
    return table(["Area", "Requirement", "Decision", "How"], DM.REQS, [12, 50, 12, 26])


def tbl_appx_a():
    rows = []
    for s in sheets:
        t = tab_by_sheet.get((s["wb"], s["sheet"]), [])
        rows.append([BOOKNAME[s["wb"]], s["sheet"], s["state"], s["dim"], n(s["cells"]), n(s["formulas"]),
                     ", ".join(x["name"] for x in t) or "-", n(t[0]["data_rows"]) if t else "-", s["dv"], s["cf"],
                     "Yes" if s["prot"] else "No", "Yes" if s["freeze"] else "No",
                     SHEETS[(s["wb"], s["sheet"])]["decision"].replace("Remove from Phase 1", "Removed from Phase 1")])
    return table(["Workbook", "Sheet", "State", "Used range", "Cells", "Formula cells", "Table", "Data rows", "Validations", "CF rules", "Protected", "Freeze", "Decision"],
                 rows, [12, 20, 7, 10, 8, 8, 12, 7, 6, 5, 6, 5, 12])


# table-level defaults for tables without a column map
TABLE_DEFAULT = {
    "tblJob": ("Remove", "X_Product (import)", "Stores copy of products replaced by import."),
    "tblPerson": ("Combine", "X_People (import)", "Single people master."),
    "tblWorker": ("Combine", "X_People (import)", "Single people master."),
    "tblStage": ("Combine", "M_Stage_Route", "Stage & route master."),
    "tblProduct@Dispatch": ("Combine", "X_Product + tblCustPrice", "Single product master + Dispatch commercial table."),
    "tblReason": ("Keep", "same", "Coded reasons; count reasons removed."),
}
COL_OVERRIDE = {
    ("tblPart", "Key"): ("Rename", "Part ID", "Permanent key replaces 'Product|Part' text."),
    ("tblPart", "Source"): ("Remove", "-", "Migration note."),
    ("tblPart", "Lines made"): ("Move to report", "MIS", ""),
    ("tblBOM", "Data Fix"): ("Combine", "Remarks", "Migration note."),
    ("tblBOMModel", "Next Review"): ("Rename", "Review Due (typed date)", "No TODAY()."),
    ("tblBOMModel", "Review Status"): ("Move to report", "MIS exception MST-02", ""),
    ("tblProduct", "Made kg"): ("Move to report", "MIS", ""),
    ("tblProduct", "Set Wt kg (BOM)"): ("Keep", "Set Wt kg (from BOM)", ""),
    ("tblMachine", "Capacity kg/day"): ("Keep", "Capacity kg/day (optional until Phase 2)", ""),
    ("tblMachine", "Shifts/day"): ("Keep", "Shifts/day (optional)", ""),
    ("tblMachine", "Hours/shift"): ("Keep", "Hours/shift (optional)", ""),
    ("tblCoaterOpening", "Confirmed By"): MAP["tblCoaterOpening"]["Confirmed By"],
}


def col_action(book, tname, col):
    if (tname, col) in COL_OVERRIDE:
        return COL_OVERRIDE[(tname, col)]
    if tname in MAP and col in MAP[tname]:
        return MAP[tname][col]
    key = f"{tname}@{book}" if f"{tname}@{book}" in TABLE_DEFAULT else tname
    if key in TABLE_DEFAULT:
        a, tgt, note = TABLE_DEFAULT[key]
        return (a, col if tgt == "same" else tgt, note)
    return ("Keep", col, "Master / reference field.")


def tbl_appx_b():
    rows = []
    unmapped = []
    for t in tables:
        if t["wb"] == "MIS" or t["name"].startswith("q"):
            continue
        tx = t["name"] in MAP
        for i, c in enumerate(t["cols"]):
            a, tgt, note = col_action(t["wb"], t["name"], c)
            if not tx and t["name"] not in TABLE_DEFAULT and (t["name"], c) not in COL_OVERRIDE and t["name"] not in ("tblSupplier", "tblMaterial", "tblGrade", "tblConsumable", "tblConsCat", "tblWhere", "tblPart", "tblBOM", "tblBOMModel", "tblProduct", "tblOperator", "tblMachine", "tblCustomer", "tblVendor"):
                unmapped.append((t["name"], c))
            calc = " (calc)" if c in t["calc"] else ""
            rows.append([BOOKNAME[t["wb"]], f"{t['sheet']} / {t['name']}", i + 1, c + calc, a, tgt, note])
    if unmapped:
        print("WARNING unmapped columns:", unmapped[:20], file=sys.stderr)
    return table(["Workbook", "Sheet / table", "#", "Current column", "Action", "Final column", "Note"], rows, [12, 22, 3, 18, 10, 22, 26])


def tbl_testcheck():
    rows = [[t[0], t[1], "[ ] 10k  [ ] 25k  [ ] 50k  [ ] 100k  [ ] 400k+", "", "", "Not yet tested"] for t in DM.TESTS]
    return table(["ID", "Test", "Volumes run", "Median time", "Errors", "Result"], rows, [5, 26, 30, 10, 10, 12])


def toc(pages):
    lines = []
    tmpl = open(os.path.join(HERE, "plan_template.md")).read()
    lines.append('<table class="toc">')
    lines.append('<tr><td class="toc-n">1.</td><td>Cover page</td><td class="toc-p">1</td></tr>')
    for m in re.finditer(r"^## (\d+)\. (.+)$", tmpl, re.M):
        num, title = m.group(1), m.group(2)
        p = pages.get(num, "")
        lines.append(f'<tr><td class="toc-n">{num}.</td><td><a href="#{slug(num, title)}">{title}</a></td><td class="toc-p">{p}</td></tr>')
    lines.append("</table>")
    return "\n".join(lines)


def slug(num, title):
    s = (num + ". " + title).lower()
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    s = re.sub(r"\s+", "-", s.strip())
    s = re.sub(r"^[0-9-]+", "", s)  # pandoc drops leading digits/punctuation
    return s


def build(pages):
    tm = open(os.path.join(HERE, "plan_template.md")).read()
    rep = {
        "TOC": toc(pages), "TBL_INVENTORY": tbl_inventory(), "TBL_HOTSPOTS": tbl_hotspots(), "TBL_PATTERNS": tbl_patterns(),
        "TBL_GROWTH": tbl_growth(), "TBL_ENTITY": tbl_entity(), "TBL_SHEETS": tbl_sheets(), "TBL_AUDIT": tbl_audit(),
        "TBL_BLUEPRINTS": tbl_blueprints(), "TBL_SCALE": tbl_scale(), "TBL_EXTRACTS": tbl_extracts(), "TBL_KPI": tbl_kpi(),
        "TBL_CHECKS": tbl_checks(), "TBL_THRESH": tbl_thresh(), "TBL_TESTS": tbl_tests(), "TBL_RISKS": tbl_risks(),
        "TBL_ROADMAP": tbl_roadmap(), "TBL_DECISIONS": tbl_decisions(), "TBL_REQS": tbl_reqs(), "TBL_APPX_A": tbl_appx_a(),
        "TBL_APPX_B": tbl_appx_b(), "TBL_TESTCHECK": tbl_testcheck(),
    }
    for k, v in rep.items():
        tm = tm.replace("{{" + k + "}}", v)
    left = re.findall(r"\{\{[A-Z_]+\}\}", tm)
    assert not left, left
    out = os.path.join(ROOT, "ISKO_ERP_100K_Row_Resilience_Plan.md")
    open(out, "w").write(tm)
    return out


if __name__ == "__main__":
    pages = json.load(open(sys.argv[1])) if len(sys.argv) > 1 else {}
    print(build(pages))
