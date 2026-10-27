"""Build everything: diagrams, mock-ups, Markdown, PDF (two passes for TOC page
numbers) and the decision-register workbook.

    python3 tools/build_all.py
Requires: pandoc, pdftotext, node + playwright (Chromium), openpyxl.
"""
import json, os, re, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
B = os.path.join(HERE, "_build")
os.makedirs(B, exist_ok=True)
MD = os.path.join(ROOT, "ISKO_ERP_100K_Row_Resilience_Plan.md")
PDF = os.path.join(ROOT, "ISKO_ERP_100K_Row_Resilience_Plan.pdf")
HTML = os.path.join(B, "plan.html")


def run(cmd, **kw):
    print("+", " ".join(cmd) if isinstance(cmd, list) else cmd)
    subprocess.run(cmd, check=True, **kw)


def html_pdf():
    run(["pandoc", MD, "-f", "markdown+pipe_tables+fenced_divs+raw_html", "-t", "html5", "-s", "--embed-resources",
         "--resource-path", ROOT, "--css", os.path.join(HERE, "style.css"), "--metadata", "pagetitle=ISKO ERP Plan", "-o", HTML], cwd=ROOT)
    run(["node", os.path.join(HERE, "render.js"), "pdf", HTML, PDF])


def toc_pages():
    tm = open(os.path.join(HERE, "plan_template.md")).read()
    heads = re.findall(r"^## (\d+)\. (.+)$", tm, re.M)
    npages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", PDF], capture_output=True, text=True).stdout).group(1))
    found = {}
    for p in range(1, npages + 1):
        txt = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", PDF, "-"], capture_output=True, text=True).stdout
        top = " ".join(l.strip() for l in txt.splitlines()[:6] if l.strip())
        top = re.sub(r"\s+", " ", top)
        if "Contents" in top or p == 1:
            continue
        for num, title in heads:
            if num in found:
                continue
            key = re.sub(r"\s+", " ", f"{num}. {title}")[:28]
            if key in top:
                found[num] = p
    missing = [h for h in heads if h[0] not in found]
    if missing:
        print("TOC: headings not located:", missing)
    return found


if __name__ == "__main__":
    py = sys.executable
    run([py, os.path.join(HERE, "render_diagrams.py")])
    run([py, os.path.join(HERE, "mockups.py")])
    run(["node", os.path.join(HERE, "render.js"), "shots", os.path.join(B, "mock"), os.path.join(ROOT, "assets")])
    run([py, os.path.join(HERE, "build_md.py")])
    html_pdf()
    pages = toc_pages()
    json.dump(pages, open(os.path.join(B, "toc_pages.json"), "w"))
    run([py, os.path.join(HERE, "build_md.py"), os.path.join(B, "toc_pages.json")])
    html_pdf()
    pages2 = toc_pages()
    if pages2 != pages:
        print("TOC pages moved; third pass")
        json.dump(pages2, open(os.path.join(B, "toc_pages.json"), "w"))
        run([py, os.path.join(HERE, "build_md.py"), os.path.join(B, "toc_pages.json")])
        html_pdf()
    if os.path.exists(os.path.join(HERE, "build_xlsx.py")):
        run([py, os.path.join(HERE, "build_xlsx.py")])
    print("done")
