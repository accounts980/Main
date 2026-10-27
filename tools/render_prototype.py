"""Render every sheet of the prototype workbooks to PNG via LibreOffice (full
recalculation on load). A temporary copy gets page setup (landscape, fit to one
page, row/column headings, gridlines) purely so each sheet becomes one image; the
delivered prototype files carry no print settings.

Usage: python3 render_prototype.py <libreoffice-profile-dir>
"""
import json, os, re, shutil, subprocess, sys
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
P = os.path.join(ROOT, "prototype")
B = os.path.join(HERE, "_build", "proto")
SHOTS = os.path.join(ROOT, "assets", "prototype")
os.makedirs(B, exist_ok=True); os.makedirs(SHOTS, exist_ok=True)
profile = sys.argv[1]

meta = json.load(open(os.path.join(P, "sheets.json")))
report = {}
for key, m in meta.items():
    src = os.path.join(P, m["file"])
    tmp = os.path.join(B, key + ".xlsx")
    wb = load_workbook(src)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.print_options.headings = True
        ws.print_options.gridLines = True
        ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = ws.page_margins.bottom = 0.2
        ws.page_margins.header = ws.page_margins.footer = 0
    wb.save(tmp)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", B, tmp],
                   check=True, capture_output=True)
    pdf = tmp[:-5] + ".pdf"
    pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    names = [ws.title for ws in wb.worksheets]
    assert pages == len(names), (key, pages, len(names))
    errs = {}
    for i, n in enumerate(names, 1):
        txt = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"], capture_output=True, text=True).stdout
        bad = re.findall(r"(Err:\d+|#NAME\?|#VALUE!|#REF!|#DIV/0!|#N/A|#NUM!)", txt)
        if bad:
            errs[n] = sorted(set(bad))
        safe = re.sub(r"[^A-Za-z0-9]+", "_", n).strip("_")
        out = os.path.join(SHOTS, f"{key}_{i:02d}_{safe}.png")
        base = os.path.join(B, "pg")
        subprocess.run(["pdftoppm", "-r", "130", "-png", "-f", str(i), "-l", str(i), "-singlefile", pdf, base], check=True)
        subprocess.run(["convert", base + ".png", "-trim", "-bordercolor", "white", "-border", "6", out], check=True)
        m["sheets"][i - 1]["image"] = os.path.relpath(out, ROOT)
    report[key] = errs
    print(key, pages, "sheets rendered; formula errors:", errs or "none")
json.dump(meta, open(os.path.join(P, "sheets.json"), "w"), indent=1)
