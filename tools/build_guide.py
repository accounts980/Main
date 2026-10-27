"""Build ISKO_ERP_Sheetwise_Guide.pdf (<= 20 pages): every sheet of the four target
workbooks with a screenshot rendered from the prototype workbooks.

    python3 tools/build_guide.py [scale]
"""
import html, json, os, re, subprocess, sys
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT_HTML = os.path.join(HERE, "_build", "guide.html")
OUT_PDF = os.path.join(ROOT, "ISKO_ERP_Sheetwise_Guide.pdf")
SCALE = float(sys.argv[1]) if len(sys.argv) > 1 else 0.70
MAXW = 271.0  # mm usable width (A4 landscape)
MAXH = 150.0  # mm max image height
meta = json.load(open(os.path.join(ROOT, "prototype", "sheets.json")))

BOOKS = [("Stores", "1 ISKO Stores", "Purchase order, receipt (GRN), lot QC, RM issue &amp; return, consumables, authorised adjustments, stock balances."),
         ("Production", "2 ISKO Production", "Cutting, bending, fabrication, rejection &amp; rework, batch balance and close, transaction-calculated WIP."),
         ("Dispatch", "3 ISKO Dispatch &amp; Challan", "Customer requirement, job-work challan, coater returns, FG balance, delivery challan."),
         ("MIS", "4 ISKO MIS", "Read-only reporting: refresh log, control tower, exceptions, stock/production/dispatch reports, trace, monthly pack.")]
KIND = {"pad": ("Entry pad", "#FFF4D6"), "reg": ("Register (values)", "#FFFFFF"), "bal": ("Balance (calculated)", "#D9F0EF"), "mst": ("Master", "#EAE4F8"),
        "imp": ("Import (values)", "#F1F4F8"), "ctl": ("Control", "#FCECC9"), "rpt": ("Report", "#E2F0E8"), "doc": ("README / settings", "#F7F9FC")}
ORDER = ["pad", "reg", "bal", "rpt", "mst", "imp", "ctl", "doc"]
SKIP_SHOT = {"Lists", "README", "Settings"}  # listed in the table, screenshot omitted (dropdown values only)

CSS = """
@page { size: A4 landscape; margin: 11mm 12mm 12mm 12mm; }
html{-webkit-print-color-adjust:exact;print-color-adjust:exact}
body{font-family:Inter,'DejaVu Sans',Arial,sans-serif;color:#243247;font-size:9pt;margin:0}
h1{font-size:24pt;margin:0 0 2mm} h2{font-size:15pt;margin:0 0 3mm;padding:2.5mm 4mm;background:#DCEAF7;border-radius:5px}
.blist{break-inside:avoid;margin-top:4mm}
h3{font-size:10.5pt;margin:3mm 0 2mm;padding:1.5mm 3mm;background:#EAE4F8;border-radius:4px;break-after:avoid}
.note{font-size:8.3pt;color:#5A6B82;margin:1mm 0 3mm}
table{border-collapse:collapse;width:100%;font-size:6.9pt;line-height:1.15} th{background:#DCEAF7;text-align:left;padding:0.8mm 1.6mm;border:1px solid #D9E0E8}
td{padding:0.3mm 1.6mm;border:1px solid #D9E0E8;vertical-align:top} tr{break-inside:avoid}
.kind{display:inline-block;padding:0.3mm 2mm;border-radius:6px;border:1px solid #D9E0E8;font-size:7.2pt;white-space:nowrap}
.grid{display:flex;flex-wrap:wrap;gap:3mm 4mm;align-items:flex-start}
figure{margin:0;break-inside:avoid} figure img{display:block;border:1px solid #C9D3DF}
figcaption{font-size:7.6pt;color:#3B4C66;margin:0.8mm 0 0;font-weight:600}
.cover{background:linear-gradient(90deg,#DCEAF7,#EAE4F8);border-radius:8px;padding:8mm 10mm;margin-bottom:5mm}
.pill{display:inline-block;padding:1mm 3mm;border-radius:8px;border:1px solid #D9E0E8;margin:0 1.5mm 1.5mm 0;font-size:8pt}
.box{border:1px solid #D9E0E8;border-radius:6px;padding:3mm 4mm;background:#F7F9FC;margin-bottom:3mm}
.cols{display:flex;gap:5mm} .cols>div{flex:1}
"""


def esc(s):
    return html.escape(str(s))


def img_mm(path):
    w, h = Image.open(os.path.join(ROOT, path)).size
    mmw = w * 25.4 / 130 * SCALE
    mmh = h * 25.4 / 130 * SCALE
    f = min(1.0, MAXW / mmw, MAXH / mmh)
    return mmw * f, mmh * f


def book_section(key, title, desc):
    m = meta[key]
    sh = m["sheets"]
    cnt = {}
    for s in sh:
        cnt[s["kind"]] = cnt.get(s["kind"], 0) + 1
    out = ["<div class='blist'>", f"<h2>{title} FY2027-28 &nbsp;-&nbsp; {len(sh)} sheets</h2>",
           f"<p class='note'>{desc} &nbsp;File: <b>{esc(m['file'].replace(' (prototype)', ''))}</b>. Sheet order below is the tab order.</p>",
           "<table><tr><th style='width:4%'>#</th><th style='width:17%'>Sheet</th><th style='width:13%'>Type</th><th>What it holds / does</th><th style='width:15%'>Who uses it</th></tr>"]
    for i, s in enumerate(sh, 1):
        k, col = KIND[s["kind"]]
        out.append(f"<tr><td>{i}</td><td><b>{esc(s['sheet'])}</b></td><td><span class='kind' style='background:{col}'>{k}</span></td><td>{esc(s['purpose'])}</td><td>{esc(s['user'])}</td></tr>")
    out.append("</table></div>")
    out.append(f"<h3>{title} - screenshots of every sheet (prototype workbook, sample data)</h3>")
    out.append("<div class='grid'>")
    idx = {s["sheet"]: i for i, s in enumerate(sh, 1)}
    shots = [s for s in sh if s["sheet"] not in SKIP_SHOT and "image" in s]
    pads = [s for s in shots if s["kind"] == "pad"]
    rest = sorted([s for s in shots if s["kind"] != "pad"], key=lambda s: (-img_mm(s["image"])[1], ORDER.index(s["kind"])))
    for s in pads + rest:
            kind = s["kind"]
            w, h = img_mm(s["image"])
            out.append(f"<figure style='width:{w:.1f}mm'><img src='../../{s['image']}' style='width:{w:.1f}mm;height:{h:.1f}mm'>"
                       f"<figcaption>{idx[s['sheet']]}. {esc(s['sheet'])} &nbsp;·&nbsp; {KIND[kind][0]}</figcaption></figure>")
    out.append("</div>")
    return "\n".join(out)


def cover():
    rows = []
    for key, title, desc in BOOKS:
        sh = meta[key]["sheets"]
        c = {}
        for s in sh:
            c[s["kind"]] = c.get(s["kind"], 0) + 1
        rows.append(f"<tr><td><b>{title}</b></td><td>{len(sh)}</td>" + "".join(f"<td>{c.get(k, 0)}</td>" for k in ORDER) + f"<td>{desc}</td></tr>")
    tot = sum(len(meta[k]["sheets"]) for k, _, _ in BOOKS)
    return f"""
<div class='cover'><h1>ISKO Excel ERP - what comes in each workbook, sheet by sheet</h1>
<div style='font-size:11pt'>Target design for FY2027-28 · four workbooks · {tot} sheets · screenshots of every sheet from the prototype workbooks · 27 Oct 2026</div></div>
<div class='cols'><div>
<table><tr><th>Workbook</th><th>Sheets</th>{''.join(f'<th>{KIND[k][0]}</th>' for k in ORDER)}<th>Scope</th></tr>{''.join(rows)}</table>
</div></div>
<div class='cols' style='margin-top:4mm'>
<div class='box'><b>One way of working in every book</b><br>
1. Type only in the <span class='pill' style='background:#FFF4D6'>Entry pad</span> sheets (cream cells). Every row shows lookups, available stock / batch / FG and a <b>Ready / issues</b> flag.<br>
2. <b>Post</b> rows that read READY: they are copied into the <span class='pill'>Register</span> as values with a permanent ID (macro button, or Copy &gt; Paste Values).<br>
3. <b>Refresh All</b>: the <span class='pill' style='background:#D9F0EF'>Balance</span> sheets rebuild for open items and every hard stop reads them.<br>
4. Corrections are reversal rows - posted rows are never edited or deleted.<br>
5. Other books receive data only through <span class='pill' style='background:#F1F4F8'>Import</span> sheets (Power Query, values) and MIS shows freshness in its Refresh Log.</div>
<div class='box'><b>Colour legend in the screenshots</b><br>
<span class='pill' style='background:#FFF4D6'>cream: you type / pick</span><span class='pill' style='background:#E2F0E8'>mint: calculated or snapshot, posted as value</span>
<span class='pill' style='background:#F7F9FC'>grey: lookup shown for information</span><span class='pill' style='background:#F8DCDD'>coral: hard stop - cannot post</span>
<span class='pill' style='background:#FCECC9'>amber: warning - needs authorised reason</span><span class='pill' style='background:#D9F0EF'>teal: totals</span><br><br>
<b>Not in any book</b> (removed from Phase 1): purchase requisition, print sheets and print buttons, Gmail / mail links and contact lists, physical stock / WIP / FG count sheets. No sheet protection, no freeze panes.</div>
</div>
<div class='box' style='background:#FFF4D6;break-after:page'><b>About the screenshots.</b> They are genuine renders of the four <b>prototype .xlsx workbooks</b> built to this design
(working Excel-2016 formulas, recalculated by LibreOffice when rendered), shown with row numbers and column letters. They use <b>sample / masked data</b> (IDs like GP-2728-0001, CUS-01, PER-007) and
example dates in April 2027 - they are <b>not</b> live screenshots of the company's current files, and the live ERP has not been built yet. The coral and amber rows are deliberate examples showing each hard stop and warning.
Full-size versions: <i>assets/prototype/</i>; the prototype workbooks: <i>prototype/</i>.</div>
"""


def flow():
    return """<h2>How the four workbooks connect</h2>
<p class='note'>Each record is entered once, in its owning book. Arrows are values-only Power Query imports refreshed at defined points (start of shift, after posting, before MIS review).</p>
<img src='../../assets/architecture.svg' style='width:100%;max-height:165mm;object-fit:contain'>"""


def build():
    body = [cover()] + [book_section(k, t, d) for k, t, d in BOOKS]
    doc = f"<!doctype html><html><head><meta charset='utf-8'><title>ISKO ERP sheet-wise guide</title><style>{CSS}</style></head><body>{''.join(body)}</body></html>"
    open(OUT_HTML, "w").write(doc)
    foot = ("<div style=\"font-family:Inter,Arial,sans-serif;font-size:7px;color:#5A6B82;width:100%;padding:0 12mm;display:flex;justify-content:space-between;\">"
            "<span>ISKO Excel ERP - sheet-wise guide to the four FY2027-28 workbooks - prototype screenshots, sample data</span>"
            "<span>Page <span class=\"pageNumber\"></span> of <span class=\"totalPages\"></span></span></div>")
    env = dict(os.environ, FOOTER=foot)
    subprocess.run(["node", os.path.join(HERE, "render.js"), "pdf", OUT_HTML, OUT_PDF], check=True, env=env, capture_output=True)
    pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", OUT_PDF], capture_output=True, text=True).stdout).group(1))
    print("scale", SCALE, "pages", pages)
    return pages


if __name__ == "__main__":
    build()
