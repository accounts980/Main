"""Write the three SVG diagrams used in the plan (light pastel palette)."""
import os
from xml.sax.saxutils import escape as _esc

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A = os.path.join(ROOT, "assets")
os.makedirs(A, exist_ok=True)

C = dict(bg="#F7F9FC", blue="#DCEAF7", lav="#EAE4F8", cream="#FFF4D6", mint="#E2F0E8", coral="#F8DCDD",
         amber="#FCECC9", teal="#D9F0EF", text="#243247", border="#D9E0E8", line="#7A8AA0", red="#B5525A")
FONT = "Inter, 'DejaVu Sans', Arial, sans-serif"


def svg(w, h, body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}" font-family="{FONT}" fill="{C["text"]}">'
            f'<defs><marker id="ar" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
            f'<path d="M0,0 L10,5 L0,10 z" fill="{C["line"]}"/></marker>'
            f'<marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
            f'<path d="M0,0 L10,5 L0,10 z" fill="{C["red"]}"/></marker></defs>'
            f'<rect width="{w}" height="{h}" fill="#FFFFFF"/>{body}</svg>')


def box(x, y, w, h, fill, stroke=None, r=8, dash=False):
    d = ' stroke-dasharray="5 4"' if dash else ""
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{fill}" stroke="{stroke or C["border"]}" stroke-width="1.2"{d}/>'


def text(x, y, s, size=11, weight=400, anchor="start", fill=None, italic=False):
    st = ' font-style="italic"' if italic else ""
    return f'<text x="{x}" y="{y}" font-size="{size}" font-weight="{weight}" text-anchor="{anchor}"{st} fill="{fill or C["text"]}">{_esc(s)}</text>'


def lines(x, y, items, size=10.5, gap=14, weight=400, anchor="start"):
    return "".join(text(x, y + i * gap, s, size, weight, anchor) for i, s in enumerate(items))


def arrow(x1, y1, x2, y2, red=False, dash=False):
    d = ' stroke-dasharray="5 4"' if dash else ""
    col = C["red"] if red else C["line"]
    m = "arr" if red else "ar"
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{col}" stroke-width="1.6" marker-end="url(#{m})"{d}/>'


def path(d, red=False, dash=False):
    da = ' stroke-dasharray="5 4"' if dash else ""
    col = C["red"] if red else C["line"]
    m = "arr" if red else "ar"
    return f'<path d="{d}" fill="none" stroke="{col}" stroke-width="1.6" marker-end="url(#{m})"{da}/>'


# ------------------------------------------------------------------ architecture
def architecture():
    W, H = 1100, 690
    b = []
    b.append(text(20, 28, "Four-workbook target architecture - ownership, values extracts, refresh points, permanent IDs", 15, 700))
    b.append(text(20, 48, "Arrows = Power Query values extracts (replace, never append) refreshed at defined points; every refresh writes a Refresh Log row.", 11, 400, fill="#4A5A70"))
    books = [
        (20, "1  STORES", "Owner: Stores head",
         ["Pads: PO · GRN · Lot QC · RM move · Consumable · Adjust"],
         ["tblPOHead / tblPOLine", "tblGRN · tblLotQC", "tblRMMove (issue = batch)", "tblConsRcpt · tblConsIssue", "tblAdjust (authorised)"],
         ["Lot Balance (open lots)", "Consumable Balance", "PO Status (open lines)"],
         ["Material · Grade · Supplier", "Consumable · Category · Where-used"]),
        (400, "2  PRODUCTION", "Owner: Production head",
         ["Pads: Production · Rejection/rework · Batch close"],
         ["tblProd", "tblRejRework (events)", "tblBatchClose"],
         ["Batch Balance (open batches)", "WIP Balance (calculated)"],
         ["Product · Part · BOM · Route", "Machine · People · Reasons"]),
        (780, "3  DISPATCH & CHALLAN", "Owner: Dispatch head",
         ["Pads: Requirement · JWC · Coater move · DC"],
         ["tblRequirement", "tblJWCHead / tblJWCLine", "tblCoaterMove", "tblDCHead / tblDCLine"],
         ["Open JWC Balance (at coater)", "FG Balance (calculated)"],
         ["Customer · Customer price & UOM", "Vendor · Coater rate (dated)"]),
    ]
    for x, title, owner, pads, regs, bals, masters in books:
        y = 70
        b.append(box(x, y, 300, 400, C["bg"]))
        b.append(f'<path d="M{x} {y+8} a8 8 0 0 1 8 -8 h284 a8 8 0 0 1 8 8 v34 h-300 z" fill="{C["blue"]}"/>')
        b.append(text(x + 14, y + 22, title, 13.5, 700))
        b.append(text(x + 14, y + 37, owner, 10, 400, fill="#4A5A70"))
        yy = y + 62
        b.append(box(x + 12, yy, 276, 34, C["cream"], r=6))
        b.append(text(x + 20, yy + 14, "ENTRY PADS (~200 rows, all checks)", 9.5, 700))
        b.append(text(x + 20, yy + 27, pads[0].replace("Pads: ", ""), 9.5))
        yy += 44
        hgt = 20 + 14 * len(regs)
        b.append(box(x + 12, yy, 276, hgt, "#FFFFFF", r=6))
        b.append(text(x + 20, yy + 14, "REGISTERS - values only, append-only", 9.5, 700))
        b.append(lines(x + 20, yy + 28, regs, 10))
        yy += hgt + 10
        hgt = 20 + 14 * len(bals)
        b.append(box(x + 12, yy, 276, hgt, C["teal"], r=6))
        b.append(text(x + 20, yy + 14, "BALANCES - rebuilt on posting", 9.5, 700))
        b.append(lines(x + 20, yy + 28, bals, 10))
        yy += hgt + 10
        hgt = 20 + 14 * len(masters)
        b.append(box(x + 12, yy, 276, hgt, C["lav"], r=6))
        b.append(text(x + 20, yy + 14, "MASTERS OWNED", 9.5, 700))
        b.append(lines(x + 20, yy + 28, masters, 10))
    # inter-book arrows
    b.append(arrow(322, 150, 398, 150)); b.append(lines(360, 128, ["X_Batch", "X_Material"], 9.5, 11, 600, "middle"))
    b.append(text(360, 166, "start of shift", 8.5, 400, "middle", fill="#4A5A70"))
    b.append(arrow(398, 210, 322, 210)); b.append(lines(360, 228, ["X_Product", "X_Machine", "X_People"], 9.5, 11, 600, "middle"))
    b.append(arrow(702, 150, 778, 150)); b.append(lines(740, 128, ["X_ProdOutput", "X_Product"], 9.5, 11, 600, "middle"))
    b.append(text(740, 166, "start of shift", 8.5, 400, "middle", fill="#4A5A70"))
    b.append(arrow(778, 210, 702, 210)); b.append(lines(740, 228, ["X_Customer"], 9.5, 11, 600, "middle"))
    # down to MIS
    for x in (170, 550, 930):
        b.append(arrow(x, 472, x, 528))
    b.append(text(560, 505, "X_ registers + balances + control totals  ·  refreshed before MIS review", 10.5, 600, "middle"))
    # MIS
    y = 530
    b.append(box(20, y, 1060, 140, C["bg"]))
    b.append(f'<path d="M20 {y+8} a8 8 0 0 1 8 -8 h1044 a8 8 0 0 1 8 8 v26 h-1060 z" fill="{C["lav"]}"/>')
    b.append(text(34, y + 21, "4  MIS  -  read-only reporting layer (nobody re-enters transactions)", 13.5, 700))
    tiles = [("Refresh Log", ["file · modified time · rows", "distinct IDs · control totals", "OK / STALE / MISSING"], C["amber"]),
             ("Exceptions", ["one row per open problem", "severity · owner · days open"], C["coral"]),
             ("Control Tower", ["director KPIs from small", "Power Query summaries"], C["teal"]),
             ("Trace", ["lot → batch → part → JWC → DC", "PO → GRN chain (by ID)"], C["mint"]),
             ("Monthly packs", ["MRM · yield · scrap", "PivotTables + slicers"], "#FFFFFF")]
    x = 34
    for t, ls, f in tiles:
        b.append(box(x, y + 46, 196, 80, f, r=6))
        b.append(text(x + 10, y + 64, t, 11, 700))
        b.append(lines(x + 10, y + 80, ls, 9.5, 13))
        x += 208
    open(os.path.join(A, "architecture.svg"), "w").write(svg(W, H, "".join(b)))


# ------------------------------------------------------------------ pattern
def pattern():
    W, H = 1000, 300
    b = []
    b.append(text(20, 26, "One way of working for every transaction table", 14, 700))
    b.append(box(20, 50, 250, 150, C["cream"]))
    b.append(text(34, 74, "ENTRY PAD", 12, 700))
    b.append(lines(34, 94, ["~200 rows, same columns as register", "lookups: one MATCH per key", "checks against balances", "pad totals per lot / batch / product", "Ready flag = all checks pass"], 10.5, 18))
    b.append(arrow(272, 110, 368, 110)); b.append(lines(320, 92, ["POST", "values + ID"], 10, 12, 700, "middle"))
    b.append(box(370, 50, 250, 150, "#FFFFFF"))
    b.append(f'<rect x="370" y="50" width="250" height="30" rx="8" fill="{C["mint"]}"/>')
    b.append(text(384, 70, "REGISTER (Excel Table)", 12, 700))
    b.append(lines(384, 100, ["values only - no formulas", "append-only; permanent IDs", "snapshots: rate, weight, UOM", "corrections = reversal rows", "up to 100,000 rows per FY file"], 10.5, 18))
    b.append(arrow(622, 110, 718, 110)); b.append(lines(670, 92, ["REFRESH", "group-by"], 10, 12, 700, "middle"))
    b.append(box(720, 50, 260, 150, C["teal"]))
    b.append(text(734, 74, "BALANCE (open items only)", 12, 700))
    b.append(lines(734, 94, ["Power Query group-by or", "bounded SUMIFS", "lots · batches · JWC lines · FG · WIP", "stores register row count", "-> STALE guard blocks entry"], 10.5, 18))
    b.append(path("M850 202 V240 H145 V204"))
    b.append(text(497, 256, "hard stops in the pad read the balance (available kg, batch kg left, JWC pending, FG available)", 10.5, 600, "middle"))
    b.append(arrow(495, 202, 495, 228, dash=True))
    b.append(text(505, 222, "X_ extract to other books / MIS", 9.5, 400, fill="#4A5A70"))
    b.append(text(20, 288, "Macro option: the POST button does copy-values, ID, Posted At/By and refresh in one step. Manual option: Copy → Paste Values, then Refresh All.", 10, 400, fill="#4A5A70"))
    open(os.path.join(A, "pattern.svg"), "w").write(svg(W, H, "".join(b)))


# ------------------------------------------------------------------ process map
def process_map():
    W, H = 1100, 660
    b = []
    b.append(text(20, 26, "End-to-end process map - PO to MIS (solid = normal flow, dashed coral = exception / reversal path)", 14, 700))
    lanes = [("STORES", 40, 160, C["blue"]), ("PRODUCTION", 200, 160, C["mint"]), ("DISPATCH & CHALLAN", 360, 170, C["amber"]), ("MIS", 530, 100, C["lav"])]
    for name, y, h, f in lanes:
        b.append(f'<rect x="10" y="{y}" width="1080" height="{h}" fill="{C["bg"]}" stroke="{C["border"]}"/>')
        b.append(f'<rect x="10" y="{y}" width="92" height="{h}" fill="{f}" stroke="{C["border"]}"/>')
        parts = name.split(" ")
        for i, p in enumerate(parts):
            b.append(text(56, y + h / 2 + (i - (len(parts) - 1) / 2) * 13 + 4, p, 10.5, 700, "middle"))
    cx = {1: 120, 2: 315, 3: 510, 4: 705, 5: 900}
    BW, BH = 160, 46

    def step(c, y, t1, t2="", fill="#FFFFFF", dash=False, stroke=None, w=BW, h=BH):
        x = cx[c]
        s = box(x, y, w, h, fill, stroke, 7, dash)
        s += text(x + w / 2, y + (19 if t2 else 27), t1, 10.5, 700, "middle")
        if t2:
            s += text(x + w / 2, y + 34, t2, 9.5, 400, "middle")
        return s
    # Stores lane
    y = 58
    b += [step(1, y, "Purchase Order", "header + lines, no requisition"), step(2, y, "GRN / receipt", "lot created, bill, PO line"),
          step(3, y, "QC & excess decision", "event log per lot", C["cream"]), step(4, y, "RM lot stock", "calculated, open lots", C["teal"]),
          step(5, y, "RM issue = batch", "lot, kg, sheets, machine")]
    for a, c in ((1, 2), (2, 3), (3, 4), (4, 5)):
        b.append(arrow(cx[a] + BW + 2, y + 23, cx[c] - 2, y + 23))
    ys = 132
    b.append(step(3, ys, "Hold / reject / excess", "return to supplier", C["coral"], True, C["red"], h=40))
    b.append(arrow(cx[3] + 80, y + BH + 2, cx[3] + 80, ys - 2, red=True, dash=True))
    b.append(step(4, ys, "Authorised adjustment", "reason + approver (no count)", C["coral"], True, C["red"], h=40))
    b.append(arrow(cx[4] + 80, ys - 2, cx[4] + 80, y + BH + 2, red=True, dash=True))
    b.append(step(5, ys, "Returns vs batch", "usable · offcut · damaged · scrap", C["coral"], True, C["red"], w=150, h=40))
    # Production lane (right to left)
    y = 218
    b += [step(5, y, "Cutting", "consumes batch kg"), step(4, y, "Bending", "stage completion"), step(3, y, "Fabrication", "stage completion"),
          step(2, y, "WIP", "transaction-calculated", C["teal"])]
    b.append(arrow(cx[5] + 160, 58 + BH + 2, cx[5] + 160, y - 2))
    b.append(text(cx[5] + 166, 194, "Batch ID", 9.5, 600, "end"))
    b.append(arrow(cx[5] + 60, y - 2, cx[5] + 60, 174, red=True, dash=True))
    for a, c in ((5, 4), (4, 3), (3, 2)):
        b.append(arrow(cx[a] - 2, y + 23, cx[c] + BW + 2, y + 23))
    ys = 292
    b.append(step(5, ys, "Batch balance & close", "variance reason + approver", C["cream"], h=40))
    b.append(arrow(cx[5] + 80, y + BH + 2, cx[5] + 80, ys - 2))
    b.append(step(3, ys, "Rejection case", "rework OK → WIP · scrap", C["coral"], True, C["red"], w=355, h=40))
    b.append(arrow(cx[4] + 60, y + BH + 2, cx[4] + 60, ys - 2, red=True, dash=True))
    b.append(arrow(cx[3] + 60, y + BH + 2, cx[3] + 60, ys - 2, red=True, dash=True))
    b.append(step(1, y, "Reversal rows", "any register: quote ID + reason", C["coral"], True, C["red"]))
    # Dispatch lane
    y = 378
    b += [step(2, y, "JWC to coater", "header + lines, rate frozen"), step(3, y, "Coater return", "BACK-OK / DEFECT-BACK"),
          step(4, y, "Finished goods", "transaction-calculated", C["teal"]), step(5, y, "Delivery challan", "UOM + rate frozen")]
    b.append(arrow(cx[2] + 80, 218 + BH + 2, cx[2] + 80, y - 2))
    b.append(text(cx[2] + 86, 330, "QC-OK parts", 9.5, 600))
    for a, c in ((2, 3), (3, 4), (4, 5)):
        b.append(arrow(cx[a] + BW + 2, y + 23, cx[c] - 2, y + 23))
    b.append(text(cx[4] - 17, y - 6, "back OK", 9, 400, "middle", fill="#4A5A70"))
    b.append(step(1, y, "Customer requirement", "versioned, Req ID", C["cream"]))
    b.append(path(f"M{cx[1]+80} {y+BH+2} V{y+140} H{cx[5]+110} V{y+BH+3}"))
    b.append(text(cx[1] + 90, y + 134, "Req ID on DC line (Decision D-12)", 9.5, 400, fill="#4A5A70"))
    b.append(step(2, y + 66, "Defect decision", "re-send · scrap · debit coater", C["coral"], True, C["red"], h=40))
    b.append(arrow(cx[3] + 20, y + BH + 2, cx[2] + BW - 10, y + 64, red=True, dash=True))
    b.append(path(f"M{cx[3]+140} {y+BH+2} V{y+120} H{cx[5]+60} V{y+BH+3}", red=True, dash=True))
    b.append(text(cx[3] + 150, y + 114, "direct dispatch from coater (DC quotes JWC line)", 9.5, 400, fill=C["red"]))
    # MIS lane
    y = 552
    b.append(box(cx[1], y, 355, 56, "#FFFFFF", r=7))
    b.append(text(cx[1] + 177, y + 22, "X_ extracts + Refresh Log", 10.5, 700, "middle"))
    b.append(text(cx[1] + 177, y + 38, "values from all registers and balances", 9.5, 400, "middle"))
    b += [step(3, y + 5, "Control Tower KPIs", "", C["teal"]), step(4, y + 5, "Exceptions", "owner · age", C["coral"]),
          step(5, y + 5, "Trace", "lot ⇄ DC by ID", C["mint"])]
    b.append(arrow(cx[1] + 357, y + 28, cx[3] - 2, y + 28))
    b.append(text(20, 650, "Not in the map (removed from Phase 1): purchase requisition, print / mail steps, physical-count and count-reconciliation flows.", 10, 600, fill="#4A5A70"))
    open(os.path.join(A, "process_map.svg"), "w").write(svg(W, H, "".join(b)))


if __name__ == "__main__":
    architecture(); pattern(); process_map()
    print("diagrams written")
