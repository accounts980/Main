# ISKO Excel ERP - 100,000-Row Resilience Plan

Planning deliverables for the rebuild of the four ISKO 2026 Excel workbooks (Stores, Production, Dispatch & Challan, MIS).
No production workbook is rebuilt or included here.

| File | What it is |
|---|---|
| `ISKO_ERP_100K_Row_Resilience_Plan.pdf` | Management planning document (88 pages) |
| `ISKO_ERP_100K_Row_Resilience_Plan.md` | Editable source of the PDF (generated tables included) |
| `ISKO_ERP_Sheet_Column_Decision_Register.xlsx` | Planning register: sheet decisions, column decisions, blueprints, risks, KPIs, tests, decisions |
| `assets/` | Diagrams (SVG) and illustrative Excel-layout mock-ups (PNG, sample data) |
| `tools/` | Audit scripts, plan data and the build pipeline |

Rebuild: `NODE_PATH=$(npm root -g) python3 tools/build_all.py` (needs pandoc, poppler-utils, Node Playwright/Chromium, openpyxl).
Edit narrative in `tools/plan_template.md` and tables in `tools/data_*.py`; the `.md` at the root is regenerated.

Performance status: **not yet benchmark-tested in Excel** - see Section 17 of the plan.
