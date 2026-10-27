"""Build the four PROTOTYPE workbooks of the target design (sample / masked data,
working Excel-2016 formulas). They demonstrate the layout of every sheet; they are
not the live ERP and contain no company transaction data.

Output: prototype/1 ISKO Stores FY2027-28 (prototype).xlsx ... 4 ISKO MIS ...
Each sheet's purpose/user text is also exported to prototype/sheets.json for the
sheet-wise guide.
"""
import datetime as dt
import json, os
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "prototype")
os.makedirs(OUT, exist_ok=True)

FILL = {k: PatternFill("solid", fgColor=v) for k, v in dict(
    h="DCEAF7", h2="EAE4F8", in_="FFF4D6", ca="E2F0E8", lk="F7F9FC", cf="FFFFFF", ex="F8DCDD", pe="FCECC9", to="D9F0EF", w="FFFFFF").items()}
TXT = "243247"
THIN = Side(style="thin", color="D9E0E8")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
D = dt.date
REG = 400  # bounded register range used in prototype formulas (real build: Excel Tables)

KIND_LABEL = {"pad": "Entry pad", "reg": "Register (values)", "bal": "Balance (calculated)", "mst": "Master", "imp": "Import (values)",
              "ctl": "Control", "rpt": "Report", "doc": "README / settings"}
INFO = {}  # book -> list of dict(sheet, kind, purpose, user, cols)


def rng(sheet, col, first=5, last=REG):
    return f"'{sheet}'!${col}${first}:${col}${last}"


class Book:
    def __init__(self, key, fname):
        self.key, self.fname = key, fname
        self.wb = Workbook(); self.wb.remove(self.wb.active)
        INFO[key] = []
        self.tables = 0

    def sheet(self, name, kind, title, note, cols, rows, purpose, user, pre=None, hdr=4, table=True, status=None, flagcol=None, widths=None):
        """cols: list of (header, cls, width, formula_template_or_None). Templates use {r}.
        rows: list of lists aligned to cols (value or None where a template applies)."""
        ws = self.wb.create_sheet(name)
        ws.sheet_view.showGridLines = True
        ws.sheet_view.zoomScale = 90
        ws["A1"] = title; ws["A1"].font = Font(bold=True, size=13, color=TXT)
        ws["A2"] = note; ws["A2"].font = Font(italic=True, size=9, color="5A6B82")
        if status:
            for j, (v, cls) in enumerate(status):
                c = ws.cell(3, 1 + j, v)
                c.fill = FILL[cls]; c.font = Font(bold=True, size=9, color=TXT); c.border = BOX
        if pre:
            for i, prow in enumerate(pre):
                for j, cell in enumerate(prow):
                    if cell is None:
                        continue
                    v, cls = cell
                    c = ws.cell(3 + i, 1 + j, v)
                    c.fill = FILL[cls]; c.border = BOX; c.font = Font(color=TXT, size=9.5, bold=cls in ("h", "h2", "to"))
                    c.alignment = Alignment(vertical="center")
        for j, col in enumerate(cols, 1):
            h, cls, w, f = col
            c = ws.cell(hdr, j, h); c.fill = FILL["h"]; c.font = Font(bold=True, color=TXT, size=9.5); c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.row_dimensions[hdr].height = 30
        if pre:
            for prow in pre:
                for j, cell in enumerate(prow):
                    if cell and cell[1] == "h2" and isinstance(cell[0], str):
                        L = get_column_letter(1 + j)
                        ws.column_dimensions[L].width = max(ws.column_dimensions[L].width or 8, len(cell[0]) + 3)
        r = hdr
        for row in rows:
            r += 1
            for j, col in enumerate(cols, 1):
                h, cls, w, f = col
                v = row[j - 1] if j - 1 < len(row) else None
                if isinstance(v, tuple) and v and v[0] == "ARR":
                    ref = f"{get_column_letter(j)}{r}"
                    ws[ref] = ArrayFormula(ref, v[1].format(r=r, p=r - 1))
                    c = ws[ref]
                elif v is None and f and f.startswith("ARR:"):
                    ref = f"{get_column_letter(j)}{r}"
                    ws[ref] = ArrayFormula(ref, f[4:].format(r=r, p=r - 1))
                    c = ws[ref]
                elif v is None and f:
                    c = ws.cell(r, j, f.format(r=r, p=r - 1))
                elif isinstance(v, str) and v.startswith("="):
                    c = ws.cell(r, j, v.format(r=r, p=r - 1))
                else:
                    c = ws.cell(r, j, v)
                c.fill = FILL[cls if cls != "cf" else "w"]; c.border = BOX; c.font = Font(color=TXT, size=9.5)
                if isinstance(v, dt.datetime):
                    c.number_format = "dd-mmm-yy hh:mm"
                elif isinstance(v, dt.date) or "Date" in h or h in ("Due", "Posted At", "As On"):
                    c.number_format = "dd-mmm-yy" if "Posted At" not in h else "dd-mmm-yy hh:mm"
                elif "%" in h:
                    c.number_format = "0.0%"
                elif any(k in h for k in ("kg", "Value", "Rs", "Rate", "Wt", "Cost")):
                    c.number_format = "#,##0.000" if "kg" in h or "Wt" in h else "#,##0.00"
        last = max(r, hdr + 1)
        if flagcol:
            L = get_column_letter(flagcol)
            area = f"{L}{hdr+1}:{L}{last}"
            ws.conditional_formatting.add(area, FormulaRule(formula=[f'OR(LEFT({L}{hdr+1},5)="READY",{L}{hdr+1}="OK",LEFT({L}{hdr+1},6)="CLOSED")'], fill=FILL["ca"], stopIfTrue=True))
            ws.conditional_formatting.add(area, FormulaRule(formula=[f'OR(LEFT({L}{hdr+1},4)="WARN",LEFT({L}{hdr+1},2)="IN",LEFT({L}{hdr+1},3)="NOT",LEFT({L}{hdr+1},7)="REORDER",LEFT({L}{hdr+1},5)="STALE",LEFT({L}{hdr+1},4)="PART",LEFT({L}{hdr+1},7)="PENDING",LEFT({L}{hdr+1},5)="READY")'], fill=FILL["pe"], stopIfTrue=True))
            ws.conditional_formatting.add(area, FormulaRule(formula=[f'LEN({L}{hdr+1})>0'], fill=FILL["ex"]))
        if table and rows:
            self.tables += 1
            tname = "tbl" + "".join(ch for ch in name.title() if ch.isalnum())[:24]
            t = Table(displayName=tname, ref=f"A{hdr}:{get_column_letter(len(cols))}{last}")
            t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
            ws.add_table(t)
        INFO[self.key].append(dict(sheet=name, kind=kind, purpose=purpose, user=user, cols=[c[0] for c in cols]))
        return ws

    def save(self):
        self.wb.calculation.fullCalcOnLoad = True
        p = os.path.join(OUT, self.fname)
        self.wb.save(p)
        return p


def readme(b, book_title, lines):
    ws = b.wb.create_sheet("README")
    ws["A1"] = book_title; ws["A1"].font = Font(bold=True, size=14, color=TXT)
    ws["A2"] = "PROTOTYPE - sample / masked data - layout of the target design, not the live ERP"; ws["A2"].font = Font(italic=True, size=9, color="B5525A")
    for i, (a, bb) in enumerate(lines, 4):
        ws.cell(i, 1, a).font = Font(bold=True, color=TXT, size=10); ws.cell(i, 1).fill = FILL["h2"] if not bb else FILL["w"]
        ws.cell(i, 2, bb).font = Font(color=TXT, size=10)
        ws.cell(i, 1).border = BOX; ws.cell(i, 2).border = BOX
    ws.column_dimensions["A"].width = 26; ws.column_dimensions["B"].width = 110
    leg = [("Cell colours", ""), ("Cream", "you type or pick here"), ("Mint", "calculated / snapshot - posted as value"), ("Light grey", "lookup shown for information"),
           ("Coral", "hard stop - row cannot be posted"), ("Amber", "warning - post only with an authorised reason"), ("Teal", "totals"), ("Blue", "column headers")]
    for i, (a, bb) in enumerate(leg, 6 + len(lines)):
        ws.cell(i, 1, a).font = Font(bold=True, size=9.5, color=TXT); ws.cell(i, 2, bb).font = Font(size=9.5, color=TXT)
    for i, k in enumerate(["h2", "in_", "ca", "lk", "ex", "pe", "to", "h"], 6 + len(lines)):
        ws.cell(i, 1).fill = FILL[k]
    INFO[b.key].append(dict(sheet="README", kind="doc", purpose="How the book works, sheet order, colour legend, daily routine.", user="Everyone", cols=[]))


def settings(b, rows):
    b.sheet("Settings", "doc", "SETTINGS", "The only place parameters are set. Typed values - no TODAY(), no CELL(). Every change goes to the Change Log.",
            [("Setting", "h2", 22, None), ("Value", "in_", 26, None), ("Note", "w", 70, None)], rows,
            "Financial year, as-on date, lock date, prefixes, tolerances, data folder.", "Book owner only", table=False)


def lists(b, cols):
    n = max(len(v) for _, v in cols)
    rows = [[v[i] if i < len(v) else None for _, v in cols] for i in range(n)]
    b.sheet("Lists", "doc", "LISTS - dropdown sources", "Visible; built from masters - no spill formulas.", [(h, "w", 20, None) for h, _ in cols], rows,
            "Dropdown sources (movement types, forms, statuses, units).", "Book owner", table=False)


# =============================================================== 1 STORES
def stores():
    b = Book("Stores", "1 ISKO Stores FY2027-28 (prototype).xlsx")
    readme(b, "1 ISKO STORES FY2027-28 - purchase, receipt, QC, RM and consumable stock", [
        ("Daily routine", ""), ("Morning", "Refresh All (X_Product, X_Machine, X_People). Check Lot Balance status = OK."),
        ("Through the day", "Type in the Entry sheets only. Post rows that read READY (macro button, or Copy > Paste Values into the Register)."),
        ("After posting", "Refresh All so Lot Balance / Consumable Balance include the new rows. Enter the posted totals in Posting Log."),
        ("End of day", "Save, back up (dated copy), Control Totals checked."), ("Never", "Edit or delete a register row - post a reversal row instead.")])
    settings(b, [["FY", "2027-28", "Financial year of this file"], ["FYCode", "2728", "Used inside IDs"], ["AsOnDate", D(2027, 4, 15), "Typed by the owner each morning"],
                 ["LockDate", D(2027, 3, 31), "Entries dated on/before this are rejected"], ["Density", 0.00000785, "kg per cubic mm (steel)"],
                 ["TolTheo", 0.03, "Theoretical-weight warning"], ["TolRate", 0.02, "Bill vs PO rate warning"], ["POPrefix", "PO/27-28/", "PO number prefix"],
                 ["AdjValueLimit", 25000, "Adjustments above this need MD approval"], ["DataFolder", "D:\\ISKO ERP\\FY2027-28\\", "Where all four books live"]])
    S = "Settings"
    AS, FY, DEN, TT = f"'{S}'!$B$7", f"'{S}'!$B$6", f"'{S}'!$B$9", f"'{S}'!$B$10"

    # ---------------- masters
    b.sheet("M_Material", "mst", "MATERIAL MASTER", "One row per material. Live stock is NOT kept here - see Lot Balance.",
            [("Material ID", "in_", 13, None), ("Grade Code", "in_", 10, None), ("Grade", "in_", 10, None), ("Thk mm", "in_", 8, None), ("Std L mm", "in_", 9, None),
             ("Std W mm", "in_", 9, None), ("Std Sheet kg", "ca", 11, f"=ROUND(D{{r}}*E{{r}}*F{{r}}*{DEN},3)"), ("Min Level kg", "in_", 11, None), ("Active", "in_", 7, None)],
            [["GPSP-1.6", "GP", "GPSP", 1.6, 2500, 1250, None, 1500, "Yes"], ["GPSP-1.2", "GP", "GPSP", 1.2, 2500, 1250, None, 1000, "Yes"],
             ["CRCA-1.0", "CR", "CRCA", 1.0, 2500, 1250, None, 800, "Yes"], ["HRCA-2.0", "HR", "HRCA", 2.0, 2500, 1250, None, 600, "Yes"]],
            "Material identity, grade code (prefix of Lot ID), standard size and sheet weight, minimum level.", "Stores head")
    b.sheet("M_Grade", "mst", "GRADE CODES", "Grade code is the front of every Lot ID.", [("Grade Code", "in_", 11, None), ("Grade", "in_", 10, None), ("Description", "in_", 30, None)],
            [["GP", "GPSP", "Galvanised plain"], ["CR", "CRCA", "Cold rolled"], ["HR", "HRCA", "Hot rolled"], ["AL", "ALUMINIUM", "Aluminium"]], "Grade codes.", "Stores head")
    b.sheet("M_Supplier", "mst", "SUPPLIERS", "Supplier ID is the key; names can change without touching history. Email = plain reference text only.",
            [("Supplier ID", "in_", 11, None), ("Supplier Name", "in_", 22, None), ("Supplies", "in_", 9, None), ("GSTIN", "in_", 18, None), ("Pay Days", "in_", 8, None),
             ("Tally Ledger", "in_", 22, None), ("Email (reference)", "in_", 20, None), ("Active", "in_", 7, None)],
            [["SUP-0007", "Supplier A (masked)", "RM", "24XXXXX0000X1Z0", 30, "SUPPLIER A PVT LTD", "", "Yes"], ["SUP-0012", "Supplier B (masked)", "RM", "24XXXXX0000X1Z1", 45, "SUPPLIER B", "", "Yes"],
             ["SUP-0020", "Supplier C (masked)", "CONS", "24XXXXX0000X1Z2", 30, "SUPPLIER C", "", "Yes"]],
            "RM and consumable suppliers with Tally ledger mapping and terms.", "Purchase")
    b.sheet("M_Consumable", "mst", "CONSUMABLE ITEMS", "Identity, unit, category, minimum level. Stock is in Consumable Balance.",
            [("Item ID", "in_", 10, None), ("Item Name", "in_", 28, None), ("Category", "in_", 12, None), ("Unit", "in_", 7, None), ("Min Level", "in_", 9, None), ("Active", "in_", 7, None)],
            [["ABR-001", "Flap disc 4 inch", "ABRASIVE", "NOS", 100, "Yes"], ["ABR-004", "Cutting disc 7 inch", "ABRASIVE", "NOS", 40, "Yes"],
             ["WLD-003", "Welding wire 1.2 mm", "WELDING", "KGS", 50, "Yes"], ["GAS-001", "CO2 gas", "GAS", "KGS", 200, "Yes"]],
            "Consumable item list.", "Stores head")
    b.sheet("M_Reason", "mst", "REASON CODES", "Codes, not free text. No count-difference reasons exist.",
            [("Code", "in_", 10, None), ("Reason", "in_", 34, None), ("Used For", "in_", 14, None)],
            [["RATE-MKT", "Market price moved", "GRN rate"], ["OVR-QLT", "Lot override - quality / size need", "Lot override"], ["ADJ-WGT", "Weighing error", "Adjustment"],
             ["ADJ-ENT", "Entry error (quote posted ID)", "Adjustment"], ["ADJ-DMG", "Damage write-off", "Adjustment"], ["ADJ-SCN", "Supplier credit note", "Adjustment"],
             ["COR-ENT", "Correction of a wrong entry", "Reversal"]], "Reason codes for rate variance, lot override, adjustment and reversal.", "Stores head")
    b.sheet("M_WhereUsed", "mst", "WHERE USED (cost centres)", "Consumable destinations.", [("Where Used", "in_", 20, None), ("Type", "in_", 12, None)],
            [["MC-01 SLTL Laser", "Machine"], ["MC-05 Amada Bending", "Machine"], ["Fabrication Bay", "Department"], ["Maintenance", "Department"]], "Consumable cost centres.", "Stores head")
    b.sheet("X_Product", "imp", "X_PRODUCT - imported from Production (values)", "Refreshed by Power Query. Do not type here.",
            [("Product ID", "lk", 11, None), ("Product Name", "lk", 20, None), ("Customer ID", "lk", 11, None), ("Active", "lk", 7, None), ("Run ID", "lk", 16, None)],
            [["PRD-A12", "Model A12 ODU", "CUS-01", "Yes", "RUN-270415-0830"], ["PRD-B07", "Model B07 IDU", "CUS-01", "Yes", "RUN-270415-0830"], ["PRD-C03", "Model C03 frame", "CUS-02", "Yes", "RUN-270415-0830"]],
            "Products for issue purpose (single master owned by Production).", "Read only")
    b.sheet("X_Machine", "imp", "X_MACHINE - imported from Production (values)", "Refreshed by Power Query.",
            [("Machine ID", "lk", 10, None), ("Machine", "lk", 18, None), ("Process", "lk", 13, None), ("Active", "lk", 7, None)],
            [["MC-01", "SLTL Laser", "CUTTING", "Yes"], ["MC-02", "Amada Punching", "CUTTING", "Yes"], ["MC-03", "Divine Laser", "CUTTING", "Yes"], ["MC-05", "Amada Bending", "BENDING", "Yes"], ["MC-07", "Fabrication Bay", "FABRICATION", "Yes"]],
            "Machines for RM issues.", "Read only")
    b.sheet("X_People", "imp", "X_PEOPLE - imported from Production (values)", "Refreshed by Power Query.",
            [("Person ID", "lk", 10, None), ("Name (masked)", "lk", 16, None), ("Role", "lk", 14, None), ("Active", "lk", 7, None)],
            [["PER-003", "Person 03", "Operator", "Yes"], ["PER-007", "Person 07", "Operator", "Yes"], ["PER-011", "Person 11", "Operator", "Yes"], ["PER-015", "Person 15", "Stores", "Yes"], ["PER-018", "Person 18", "QC / approver", "Yes"]],
            "People who receive, issue, approve.", "Read only")

    # ---------------- PO
    POH, POL = "PO Header Register", "PO Line Register"
    b.sheet(POH, "reg", "PO HEADER REGISTER - values only", "Posted from PO Entry. Status changes are new status lines, never edits.",
            [("PO No", "ca", 15, None), ("PO Date", "ca", 10, None), ("Supplier ID", "ca", 10, None), ("Supplier Name", "ca", 18, None), ("Delivery By", "ca", 10, None),
             ("Pay Days", "ca", 8, None), ("Delivery Location", "ca", 13, None), ("Approved By", "ca", 10, None), ("PO Status", "ca", 9, None), ("Posted At", "ca", 13, None), ("Posted By", "ca", 9, None), ("Entry Source", "ca", 9, None)],
            [["PO/27-28/001", D(2027, 4, 2), "SUP-0007", "Supplier A (masked)", D(2027, 4, 12), 30, "ISKO Factory", "PER-018", "OPEN", dt.datetime(2027, 4, 2, 11, 5), "PER-015", "PAD"],
             ["PO/27-28/002", D(2027, 4, 3), "SUP-0012", "Supplier B (masked)", D(2027, 4, 14), 45, "ISKO Factory", "PER-018", "OPEN", dt.datetime(2027, 4, 3, 16, 40), "PER-015", "PAD"]],
            "One row per PO (no requisition reference).", "Read only (posted)")
    b.sheet(POL, "reg", "PO LINE REGISTER - values only", "Received / pending are NOT stored here - see PO Status.",
            [("PO Line ID", "ca", 16, None), ("PO No", "ca", 13, None), ("Item ID", "ca", 10, None), ("Grade", "ca", 7, None), ("Thk mm", "ca", 7, None),
             ("L mm", "ca", 7, None), ("W mm", "ca", 7, None), ("Ordered kg", "ca", 10, None), ("Unit", "ca", 6, None), ("Rate Rs", "ca", 8, None), ("Line Value Rs", "ca", 12, None), ("Due Date", "ca", 10, None), ("Posted At", "ca", 13, None)],
            [["PO/27-28/001-01", "PO/27-28/001", "GPSP-1.6", "GPSP", 1.6, 2500, 1250, 3000, "KGS", 74.5, 223500, D(2027, 4, 12), dt.datetime(2027, 4, 2, 11, 5)],
             ["PO/27-28/001-02", "PO/27-28/001", "GPSP-1.2", "GPSP", 1.2, 2500, 1250, 2400, "KGS", 74.5, 178800, D(2027, 4, 12), dt.datetime(2027, 4, 2, 11, 5)],
             ["PO/27-28/002-01", "PO/27-28/002", "CRCA-1.0", "CRCA", 1.0, 2500, 1250, 1500, "KGS", 68.2, 102300, D(2027, 4, 14), dt.datetime(2027, 4, 3, 16, 40)]],
            "One row per PO item.", "Read only (posted)")
    b.sheet("PO Entry", "pad", "PO ENTRY PAD - one PO at a time", "Fill the header, then the lines. Post when every line reads READY. No requisition is needed.",
            [("Line", "lk", 6, None), ("PO Line ID", "ca", 16, "=IF(C{r}=\"\",\"\",$B$3&\"-\"&TEXT(A{r},\"00\"))"), ("Item ID", "in_", 10, None),
             ("Grade", "ca", 8, f"=IF(C{{r}}=\"\",\"\",IFERROR(INDEX('M_Material'!$C$5:$C$50,MATCH(C{{r}},'M_Material'!$A$5:$A$50,0)),\"?\"))"),
             ("Thk mm", "ca", 7, f"=IF(C{{r}}=\"\",\"\",IFERROR(INDEX('M_Material'!$D$5:$D$50,MATCH(C{{r}},'M_Material'!$A$5:$A$50,0)),\"\"))"),
             ("L mm", "in_", 7, None), ("W mm", "in_", 7, None), ("Ordered kg", "in_", 10, None), ("Unit", "ca", 6, "=IF(C{r}=\"\",\"\",\"KGS\")"), ("Rate Rs", "in_", 8, None),
             ("Line Value Rs", "ca", 12, "=IF(H{r}=\"\",\"\",ROUND(H{r}*J{r},2))"), ("Due Date", "in_", 10, None),
             ("Ready / issues", "cf", 30, f"=IF(C{{r}}=\"\",\"\",IF(D{{r}}=\"?\",\"ITEM NOT IN MASTER\",IF(N(H{{r}})<=0,\"QTY REQUIRED\",IF(N(J{{r}})<=0,\"RATE REQUIRED\",IF(L{{r}}<$B$4,\"WARN: DUE BEFORE PO DATE\",\"READY\")))))")],
            [[1, None, "GPSP-1.6", None, None, 2500, 1250, 2000, None, 74.5, None, D(2027, 4, 25)],
             [2, None, "HRCA-2.0", None, None, 2500, 1250, 800, None, None, None, D(2027, 4, 25)],
             [3, None, "GPSP-2.5", None, None, 2500, 1250, 500, None, 75, None, D(2027, 4, 25)],
             [4, None, "CRCA-1.0", None, None, 2500, 1250, 1000, None, 68.2, None, D(2027, 4, 1)]],
            "Create a PO (header + lines) with checks; post to PO Header and PO Line registers.", "Purchase / Stores",
            pre=[[("PO No", "h2"), (f"=\"PO/27-28/\"&TEXT(COUNTA({rng(POH,'A')})+1,\"000\")", "ca"), ("Supplier", "h2"), ("SUP-0007", "in_"), ("Name", "h2"),
                  (f"=IFERROR(INDEX('M_Supplier'!$B$5:$B$50,MATCH(D3,'M_Supplier'!$A$5:$A$50,0)),\"\")", "ca"), ("Pay days", "h2"),
                  (f"=IFERROR(INDEX('M_Supplier'!$E$5:$E$50,MATCH(D3,'M_Supplier'!$A$5:$A$50,0)),\"\")", "ca"), ("Approved by", "h2"), ("PER-018", "in_")],
                 [("PO Date", "h2"), (D(2027, 4, 15), "in_"), ("Delivery by", "h2"), (D(2027, 4, 25), "in_"), ("Location", "h2"), ("ISKO Factory", "in_")]],
            hdr=6, table=False, flagcol=13)
    ws = b.wb["PO Entry"]; ws["B4"].number_format = "dd-mmm-yy"; ws["D4"].number_format = "dd-mmm-yy"
    GRN = "GRN Register"
    b.sheet("PO Status", "bal", "PO STATUS - open PO lines (calculated)", "Ordered vs received; overdue lines feed the MIS exception list.",
            [("PO Line ID", "lk", 16, None), ("Supplier ID", "lk", 10, f"=INDEX({rng(POH,'C')},MATCH(LEFT(A{{r}},12),{rng(POH,'A')},0))"),
             ("Item ID", "lk", 10, f"=INDEX({rng(POL,'C')},MATCH(A{{r}},{rng(POL,'A')},0))"), ("Ordered kg", "lk", 10, f"=INDEX({rng(POL,'H')},MATCH(A{{r}},{rng(POL,'A')},0))"),
             ("Received kg", "ca", 11, f"=SUMIFS({rng(GRN,'N')},{rng(GRN,'F')},A{{r}})"), ("Pending kg", "ca", 10, "=MAX(0,D{r}-E{r})"),
             ("Due Date", "lk", 10, f"=INDEX({rng(POL,'L')},MATCH(A{{r}},{rng(POL,'A')},0))"), ("Overdue Days", "ca", 9, f"=IF(F{{r}}>0,MAX(0,{AS}-G{{r}}),0)"),
             ("Status", "cf", 18, "=IF(F{r}=0,\"CLOSED - fully received\",IF(H{r}>0,\"OVERDUE \"&H{r}&\" d\",\"OK\"))")],
            [["PO/27-28/001-01"], ["PO/27-28/001-02"], ["PO/27-28/002-01"]],
            "Received and pending quantity per open PO line; feeds GRN dropdown and MIS.", "Purchase (read only)", flagcol=9)
    # ---------------- GRN
    b.sheet(GRN, "reg", "GRN REGISTER - one row per lot, values only", "Receipt facts only. QC decisions are in Lot QC Register.",
            [("Lot ID", "ca", 13, None), ("GRN Date", "ca", 9, None), ("Supplier ID", "ca", 9, None), ("Supplier Inv No", "ca", 11, None), ("Invoice Date", "ca", 9, None),
             ("PO Line ID", "ca", 15, None), ("Material ID", "ca", 10, None), ("Grade Code", "ca", 6, None), ("Thk mm", "ca", 6, None), ("L mm", "ca", 6, None), ("W mm", "ca", 6, None),
             ("Form", "ca", 7, None), ("Sheets", "ca", 7, None), ("Received kg", "ca", 10, None), ("Theo kg", "ca", 9, None), ("Theo Var %", "ca", 7, None), ("Bill Rate Rs", "ca", 8, None),
             ("PO Rate Rs", "ca", 8, None), ("Excess kg", "ca", 8, None), ("Lot Value Rs", "ca", 11, None), ("Posted At", "ca", 12, None), ("Entry Source", "ca", 9, None)],
            [["HR-2728-0001", D(2027, 4, 1), "SUP-0012", "OPENING", D(2027, 4, 1), "NO-PO", "HRCA-2.0", "HR", 2.0, 2500, 1250, "SHEET", 9, 640.2, 441.6, None, 66.1, None, 0, 42317.22, dt.datetime(2027, 4, 1, 9, 0), "OPENING"],
             ["GP-2728-0001", D(2027, 4, 8), "SUP-0007", "INV/3381", D(2027, 4, 7), "PO/27-28/001-01", "GPSP-1.6", "GP", 1.6, 2500, 1250, "SHEET", 40, 1575.5, 1570, 0.0035, 74.5, 74.5, 0, 117374.75, dt.datetime(2027, 4, 8, 12, 10), "PAD"],
             ["GP-2728-0002", D(2027, 4, 8), "SUP-0007", "INV/3381", D(2027, 4, 7), "PO/27-28/001-02", "GPSP-1.2", "GP", 1.2, 2500, 1250, "SHEET", 60, 1750, 1766.25, -0.0092, 74.5, 74.5, 0, 130375, dt.datetime(2027, 4, 8, 12, 10), "PAD"],
             ["CR-2728-0001", D(2027, 4, 12), "SUP-0012", "B-77120", D(2027, 4, 12), "PO/27-28/002-01", "CRCA-1.0", "CR", 1.0, 2500, 1250, "SHEET", 30, 736, 735.94, 0.0001, 68.2, 68.2, 0, 50195.2, dt.datetime(2027, 4, 12, 15, 0), "PAD"],
             ["GP-2728-0003", D(2027, 4, 13), "SUP-0007", "INV/3402", D(2027, 4, 13), "PO/27-28/001-01", "GPSP-1.6", "GP", 1.6, 2500, 1250, "SHEET", 30, 1180, 1177.5, 0.0021, 74.5, 74.5, 0, 87910, dt.datetime(2027, 4, 13, 11, 30), "PAD"]],
            "Every received lot with its permanent Lot ID.", "Read only (posted)")
    b.sheet("GRN Entry", "pad", "GRN ENTRY PAD", "One row per lot received. Lot ID fills itself. The lot cannot be issued until Lot QC records ACCEPTED.",
            [("Lot ID", "ca", 13, f"=IF(C{{r}}=\"\",\"\",G{{r}}&\"-\"&{FY}&\"-\"&TEXT(COUNTIF({rng(GRN,'H')},G{{r}})+COUNTIF(G$4:G{{p}},G{{r}})+1,\"0000\"))"),
             ("GRN Date", "in_", 9, None), ("Supplier ID", "in_", 9, None), ("Supplier Inv No", "in_", 11, None), ("PO Line ID", "in_", 15, None),
             ("Material ID", "ca", 10, f"=IF(E{{r}}=\"\",\"\",IFERROR(INDEX({rng(POL,'C')},MATCH(E{{r}},{rng(POL,'A')},0)),\"\"))"),
             ("Grade Code", "ca", 6, "=IF(F{r}=\"\",\"\",INDEX('M_Material'!$B$5:$B$50,MATCH(F{r},'M_Material'!$A$5:$A$50,0)))"),
             ("Thk mm", "ca", 6, "=IF(F{r}=\"\",\"\",INDEX('M_Material'!$D$5:$D$50,MATCH(F{r},'M_Material'!$A$5:$A$50,0)))"),
             ("L mm", "in_", 6, None), ("W mm", "in_", 6, None), ("Form", "in_", 7, None), ("Sheets", "in_", 7, None), ("Received kg", "in_", 10, None),
             ("Theo kg", "ca", 9, f"=IF(OR(L{{r}}=\"\",H{{r}}=\"\"),\"\",ROUND(L{{r}}*I{{r}}*J{{r}}*H{{r}}*{DEN},3))"), ("Theo Var %", "ca", 7, "=IF(N{r}=\"\",\"\",M{r}/N{r}-1)"),
             ("Bill Rate Rs", "in_", 8, None), ("PO Rate Rs", "ca", 8, f"=IF(E{{r}}=\"\",\"\",IFERROR(INDEX({rng(POL,'J')},MATCH(E{{r}},{rng(POL,'A')},0)),\"\"))"),
             ("PO Pending kg", "lk", 9, f"=IF(E{{r}}=\"\",\"\",IFERROR(INDEX('PO Status'!$F$5:$F$100,MATCH(E{{r}},'PO Status'!$A$5:$A$100,0)),0)-SUMIFS(M$4:M{{p}},E$4:E{{p}},E{{r}}))"),
             ("Excess kg", "ca", 8, "=IF(R{r}=\"\",\"\",MAX(0,M{r}-R{r}))"),
             ("Ready / issues", "cf", 34, f"=IF(C{{r}}=\"\",\"\",IF(F{{r}}=\"\",\"NO OPEN PO LINE\",IF(N(M{{r}})<=0,\"RECEIVED KG REQUIRED\",IF(D{{r}}=\"\",\"INVOICE NO REQUIRED\",IF(S{{r}}>0,\"WARN: EXCESS \"&TEXT(S{{r}},\"0.0\")&\" kg - decide in Lot QC\",IF(ABS(O{{r}})>{TT},\"WARN: THEO VAR \"&TEXT(O{{r}},\"+0.0%\"),\"READY\"))))))")],
            [[None, D(2027, 4, 15), "SUP-0007", "INV/3419", "PO/27-28/001-01", None, None, None, 2500, 1250, "SHEET", 6, 245.0, None, None, 74.5],
             [None, D(2027, 4, 15), "SUP-0007", "INV/3419", "PO/27-28/001-02", None, None, None, 2500, 1250, "SHEET", 20, 590, None, None, 74.5],
             [None, D(2027, 4, 15), "SUP-0012", "B-77311", "PO/27-28/009-01", None, None, None, 2500, 1250, "SHEET", 10, 245.4, None, None, 68.2],
             [None, D(2027, 4, 15), "SUP-0012", "", "PO/27-28/002-01", None, None, None, 2500, 1250, "SHEET", 25, 613.2, None, None, 68.2]],
            "Receive lots against open PO lines; theoretical weight, rate and excess checks.", "Stores", flagcol=20, table=False)
    QC = "Lot QC Register"
    b.sheet(QC, "reg", "LOT QC REGISTER - QC and excess decisions (append-only)", "The latest decision per lot decides whether it can be issued.",
            [("QC Event ID", "ca", 15, None), ("Lot ID", "ca", 13, None), ("QC Date", "ca", 9, None), ("QC Status", "ca", 22, None), ("Rejected kg", "ca", 9, None),
             ("Rejected Sheets", "ca", 8, None), ("Excess Decision", "ca", 13, None), ("Decided By", "ca", 9, None), ("Accepted kg", "ca", 10, None), ("Accepted Sheets", "ca", 9, None), ("Posted At", "ca", 12, None)],
            [["QC-2728-000001", "HR-2728-0001", D(2027, 4, 1), "ACCEPTED", 0, 0, "", "PER-018", 640.2, 9, dt.datetime(2027, 4, 1, 9, 30)],
             ["QC-2728-000002", "GP-2728-0001", D(2027, 4, 8), "ACCEPTED", 0, 0, "", "PER-018", 1575.5, 40, dt.datetime(2027, 4, 8, 16, 0)],
             ["QC-2728-000003", "GP-2728-0002", D(2027, 4, 9), "ACCEPTED WITH REJECTION", 26.4, 1, "", "PER-018", 1723.6, 59, dt.datetime(2027, 4, 9, 10, 15)],
             ["QC-2728-000004", "CR-2728-0001", D(2027, 4, 12), "HOLD", 0, 0, "", "PER-018", 0, 0, dt.datetime(2027, 4, 12, 17, 0)],
             ["QC-2728-000005", "GP-2728-0003", D(2027, 4, 13), "ACCEPTED", 0, 0, "", "PER-018", 1180, 30, dt.datetime(2027, 4, 13, 15, 0)]],
            "QC status per lot (accepted / hold / rejected) and excess decisions.", "QC / Purchase")
    # ---------------- RM movements
    MV, ADJ = "RM Move Register", "Adjust Register"
    b.sheet(MV, "reg", "RM MOVEMENT REGISTER - values only", "ISSUE line = batch (Batch ID = Move ID). Returns and scrap quote the batch. Corrections = reversal rows.",
            [("Move ID", "ca", 15, None), ("Move Date", "ca", 9, None), ("Movement", "ca", 13, None), ("Batch ID", "ca", 15, None), ("Material ID", "ca", 9, None), ("Lot ID", "ca", 12, None),
             ("Form", "ca", 7, None), ("Sheets", "ca", 6, None), ("Weight kg", "ca", 9, None), ("Stock Effect kg", "ca", 10, None), ("Stock Effect Sh", "ca", 8, None), ("Rate Rs", "ca", 7, None),
             ("Value Rs", "ca", 10, None), ("Machine ID", "ca", 8, None), ("Issued To", "ca", 8, None), ("Purpose", "ca", 8, None), ("Avail kg Before", "ca", 9, None), ("Posted At", "ca", 12, None), ("Reverses ID", "ca", 10, None)],
            [["ISS-2728-000001", D(2027, 4, 12), "ISSUE", "ISS-2728-000001", "GPSP-1.6", "GP-2728-0001", "SHEET", 12, 471, -471, -12, 74.5, -35089.5, "MC-01", "PER-007", "PRD-A12", 1575.5, dt.datetime(2027, 4, 12, 9, 20), ""],
             ["ISS-2728-000002", D(2027, 4, 12), "ISSUE", "ISS-2728-000002", "GPSP-1.2", "GP-2728-0002", "SHEET", 15, 420, -420, -15, 74.5, -31290, "MC-03", "PER-011", "PRD-B07", 1723.6, dt.datetime(2027, 4, 12, 9, 25), ""],
             ["OFC-2728-000003", D(2027, 4, 13), "RETURN-OFFCUT", "ISS-2728-000001", "GPSP-1.6", "GP-2728-0001", "OFFCUT", 0, 18.4, 18.4, 0, 74.5, 1370.8, "", "PER-007", "", 1104.5, dt.datetime(2027, 4, 13, 17, 10), ""],
             ["ISS-2728-000004", D(2027, 4, 13), "ISSUE", "ISS-2728-000004", "HRCA-2.0", "HR-2728-0001", "SHEET", 3, 214.9, -214.9, -3, 66.1, -14204.89, "MC-02", "PER-003", "PRD-C03", 640.2, dt.datetime(2027, 4, 13, 10, 0), ""]],
            "All RM issues, returns, offcut/damaged returns, stores scrap, returns to supplier.", "Read only (posted)")
    b.sheet(ADJ, "reg", "AUTHORISED STOCK ADJUSTMENT REGISTER", "Approved corrections only - reason, reference, independent approver. Not a count module.",
            [("Adj ID", "ca", 13, None), ("Adj Date", "ca", 9, None), ("Stock Type", "ca", 9, None), ("Lot / Item ID", "ca", 12, None), ("Direction", "ca", 9, None), ("Adj kg / Qty", "ca", 9, None),
             ("Signed kg", "ca", 9, "=IF(E{r}=\"DECREASE\",-F{r},F{r})"), ("Rate Rs", "ca", 7, None), ("Value Rs", "ca", 9, "=ROUND(G{r}*H{r},2)"), ("Reason Code", "ca", 9, None), ("Reference", "ca", 16, None),
             ("Requested By", "ca", 9, None), ("Approved By", "ca", 9, None), ("Posted At", "ca", 12, None)],
            [["ADJ-2728-00001", D(2027, 4, 14), "RM-LOT", "GP-2728-0002", "DECREASE", 4.2, None, 74.5, None, "ADJ-WGT", "Weighbridge slip WB-118", "PER-015", "PER-018", dt.datetime(2027, 4, 14, 12, 0)]],
            "Approved stock corrections (RM lots and consumables).", "Read only (posted)")
    LB = "Lot Balance"
    b.sheet(LB, "bal", "LOT BALANCE - open lots, oldest first (calculated)", "Rebuilt after posting (Refresh All). Every issue check reads this sheet.",
            [("Lot ID", "lk", 13, None), ("Material ID", "lk", 9, f"=INDEX({rng(GRN,'G')},MATCH(A{{r}},{rng(GRN,'A')},0))"), ("GRN Date", "lk", 9, f"=INDEX({rng(GRN,'B')},MATCH(A{{r}},{rng(GRN,'A')},0))"),
             ("QC Status", "lk", 15, f"=IFERROR(INDEX({rng(QC,'D')},MATCH(A{{r}},{rng(QC,'B')},0)),\"PENDING\")"),
             ("Issue Eligible", "ca", 8, "=IF(OR(D{r}=\"ACCEPTED\",D{r}=\"ACCEPTED WITH REJECTION\"),\"Yes\",\"No\")"),
             ("Accepted kg", "ca", 10, f"=SUMIFS({rng(QC,'I')},{rng(QC,'B')},A{{r}})"), ("Accepted Sh", "ca", 8, f"=SUMIFS({rng(QC,'J')},{rng(QC,'B')},A{{r}})"),
             ("Movement kg", "ca", 10, f"=SUMIFS({rng(MV,'J')},{rng(MV,'F')},A{{r}})"), ("Movement Sh", "ca", 8, f"=SUMIFS({rng(MV,'K')},{rng(MV,'F')},A{{r}})"),
             ("Adjust kg", "ca", 8, f"=SUMIFS({rng(ADJ,'G')},{rng(ADJ,'D')},A{{r}})"), ("Available kg", "to", 10, "=ROUND(F{r}+H{r}+J{r},3)"), ("Available Sh", "to", 8, "=G{r}+I{r}"),
             ("Rate Rs", "lk", 7, f"=INDEX({rng(GRN,'Q')},MATCH(A{{r}},{rng(GRN,'A')},0))"), ("Value Rs", "ca", 11, "=ROUND(K{r}*M{r},2)"),
             ("Age Days", "ca", 7, f"={AS}-C{{r}}"), ("Status", "cf", 18, "=IF(K{r}<0,\"BREACH - negative\",IF(E{r}=\"No\",\"NOT ISSUE-ELIGIBLE\",IF(O{r}>90,\"WARN: over 90 days\",\"OK\")))")],
            [["HR-2728-0001"], ["GP-2728-0001"], ["GP-2728-0002"], ["CR-2728-0001"], ["GP-2728-0003"]],
            "Available kg and sheets per open lot, eligibility, value, ageing.", "Stores (read only)", flagcol=16,
            status=[("Rows at refresh", "h2"), (4, "lk"), ("Live rows", "h2"), (f"=COUNTA({rng(MV,'A')})", "lk"), ("=IF(D3=B3,\"OK\",\"STALE - Refresh All\")", "ca")])
    lb = b.wb[LB]
    for c in ("B1", "C1", "D1", "E1", "F1"):
        pass
    LBR = lambda col: rng(LB, col, 5, 100)
    b.sheet("RM Move Entry", "pad", "RM ISSUE & RETURN ENTRY PAD", "ISSUE creates a batch. Returns / offcuts / damaged / scrap must quote an open batch of the same lot. Max available is shown on every row.",
            [("Move ID", "ca", 15, f"=IF(C{{r}}=\"\",\"\",IF(C{{r}}=\"ISSUE\",\"ISS\",IF(C{{r}}=\"RETURN-OFFCUT\",\"OFC\",IF(C{{r}}=\"STORES-SCRAP\",\"SCR\",\"RET\")))&\"-\"&{FY}&\"-\"&TEXT(COUNTA({rng(MV,'A')})+ROW()-4,\"000000\"))"),
             ("Move Date", "in_", 9, None), ("Movement", "in_", 13, None), ("Batch ID", "ca", 15, "=IF(C{r}=\"ISSUE\",A{r},\"\")"), ("Material ID", "in_", 9, None),
             ("FIFO Lot", "lk", 12, None), ("Lot ID", "in_", 12, None), ("Override Reason", "in_", 9, None), ("Form", "in_", 7, None), ("Sheets", "in_", 6, None), ("Weight kg", "in_", 9, None),
             ("Avail kg", "lk", 9, f"=IF(G{{r}}=\"\",\"\",IFERROR(INDEX({LBR('K')},MATCH(G{{r}},{LBR('A')},0)),0)-SUMIFS(K$4:K{{p}},G$4:G{{p}},G{{r}},C$4:C{{p}},\"ISSUE\"))"),
             ("Avail Sh", "lk", 7, f"=IF(G{{r}}=\"\",\"\",IFERROR(INDEX({LBR('L')},MATCH(G{{r}},{LBR('A')},0)),0)-SUMIFS(J$4:J{{p}},G$4:G{{p}},G{{r}},C$4:C{{p}},\"ISSUE\"))"),
             ("Stock Effect kg", "ca", 10, "=IF(C{r}=\"\",\"\",IF(C{r}=\"ISSUE\",-K{r},IF(OR(C{r}=\"RETURN-USABLE\",C{r}=\"RETURN-OFFCUT\"),K{r},0)))"),
             ("Rate Rs", "ca", 7, f"=IF(G{{r}}=\"\",\"\",IFERROR(INDEX({rng(GRN,'Q')},MATCH(G{{r}},{rng(GRN,'A')},0)),\"\"))"),
             ("Machine ID", "in_", 8, None), ("Issued To", "in_", 8, None),
             ("Ready / issues", "cf", 40, f"=IF(C{{r}}=\"\",\"\",IF(IFERROR(INDEX({LBR('E')},MATCH(G{{r}},{LBR('A')},0)),\"No\")=\"No\",\"LOT NOT ISSUE-ELIGIBLE (QC)\",IF(AND(C{{r}}=\"ISSUE\",K{{r}}>L{{r}}),\"EXCEEDS AVAILABLE: max \"&TEXT(L{{r}},\"0.000\")&\" kg / \"&M{{r}}&\" sh\",IF(AND(C{{r}}<>\"ISSUE\",D{{r}}=\"\"),\"RETURN MUST QUOTE BATCH\",IF(AND(C{{r}}=\"ISSUE\",G{{r}}<>F{{r}},H{{r}}=\"\"),\"WARN: NOT FIFO - override reason required\",\"READY\")))))")],
            [[None, D(2027, 4, 15), "ISSUE", None, "GPSP-1.6", ("ARR", f"=IFERROR(INDEX({LBR('A')},MATCH(1,({LBR('B')}=E{{r}})*({LBR('E')}=\"Yes\")*({LBR('K')}>0),0)),\"\")"), "GP-2728-0001", "", "SHEET", 8, 300, None, None, None, None, "MC-01", "PER-007"],
             [None, D(2027, 4, 15), "ISSUE", None, "GPSP-1.6", ("ARR", f"=IFERROR(INDEX({LBR('A')},MATCH(1,({LBR('B')}=E{{r}})*({LBR('E')}=\"Yes\")*({LBR('K')}>0),0)),\"\")"), "GP-2728-0001", "", "SHEET", 25, 980, None, None, None, None, "MC-03", "PER-011"],
             [None, D(2027, 4, 15), "ISSUE", None, "CRCA-1.0", ("ARR", f"=IFERROR(INDEX({LBR('A')},MATCH(1,({LBR('B')}=E{{r}})*({LBR('E')}=\"Yes\")*({LBR('K')}>0),0)),\"\")"), "CR-2728-0001", "", "SHEET", 5, 122.6, None, None, None, None, "MC-02", "PER-003"],
             [None, D(2027, 4, 15), "ISSUE", None, "GPSP-1.6", ("ARR", f"=IFERROR(INDEX({LBR('A')},MATCH(1,({LBR('B')}=E{{r}})*({LBR('E')}=\"Yes\")*({LBR('K')}>0),0)),\"\")"), "GP-2728-0003", "", "SHEET", 4, 157.3, None, None, None, None, "MC-01", "PER-007"],
             [None, D(2027, 4, 15), "RETURN-OFFCUT", "ISS-2728-000002", "GPSP-1.2", "", "GP-2728-0002", "", "OFFCUT", 0, 12.5, None, None, None, None, "", "PER-011"]],
            "Issue RM to a machine (creates batch) and record returns against a batch; FIFO suggestion and availability hard stop.", "Stores", flagcol=18, table=False,
            status=[("Lot Balance:", "h2"), ("='Lot Balance'!E3", "ca")])
    ws = b.wb["RM Move Entry"]
    ws["D9"] = "ISS-2728-000002"; ws["D9"].fill = FILL["in_"]
    # ---------------- consumables
    CR, CI = "Cons Receipt Register", "Cons Issue Register"
    b.sheet(CR, "reg", "CONSUMABLE RECEIPT REGISTER - values only", "Purchases and the one-time OPENING lines.",
            [("Receipt ID", "ca", 14, None), ("Receipt Date", "ca", 9, None), ("Item ID", "ca", 9, None), ("Qty", "ca", 7, None), ("Bill Rate Rs", "ca", 8, None), ("Value Rs", "ca", 9, "=ROUND(D{r}*E{r},2)"),
             ("Receipt Type", "ca", 9, None), ("Supplier ID", "ca", 9, None), ("Bill No", "ca", 9, None), ("Posted At", "ca", 12, None)],
            [["CR-2728-000001", D(2027, 4, 1), "ABR-001", 120, 23.3, None, "OPENING", "", "", dt.datetime(2027, 4, 1, 9, 0)], ["CR-2728-000002", D(2027, 4, 1), "WLD-003", 80, 148, None, "OPENING", "", "", dt.datetime(2027, 4, 1, 9, 0)],
             ["CR-2728-000003", D(2027, 4, 1), "GAS-001", 150, 15, None, "OPENING", "", "", dt.datetime(2027, 4, 1, 9, 0)], ["CR-2728-000004", D(2027, 4, 9), "ABR-004", 50, 64, None, "PURCHASE", "SUP-0020", "C-5512", dt.datetime(2027, 4, 9, 14, 0)]],
            "Consumable receipts.", "Read only (posted)")
    b.sheet(CI, "reg", "CONSUMABLE ISSUE REGISTER - values only", "Issues and returns as separate rows. Rate = average at posting (frozen).",
            [("Cons Move ID", "ca", 14, None), ("Move Date", "ca", 9, None), ("Movement", "ca", 8, None), ("Item ID", "ca", 9, None), ("Qty", "ca", 7, None), ("Stock Effect", "ca", 8, None),
             ("Rate Rs", "ca", 7, None), ("Value Rs", "ca", 9, None), ("Issued To", "ca", 8, None), ("Where Used", "ca", 16, None), ("Posted At", "ca", 12, None)],
            [["CI-2728-000001", D(2027, 4, 10), "ISSUE", "ABR-001", 40, -40, 23.3, -932, "PER-007", "MC-01 SLTL Laser", dt.datetime(2027, 4, 10, 10, 0)],
             ["CI-2728-000002", D(2027, 4, 11), "ISSUE", "GAS-001", 120, -120, 15, -1800, "PER-030", "Fabrication Bay", dt.datetime(2027, 4, 11, 9, 0)],
             ["CI-2728-000003", D(2027, 4, 12), "RETURN", "ABR-001", 5, 5, 23.3, 116.5, "PER-007", "MC-01 SLTL Laser", dt.datetime(2027, 4, 12, 17, 0)]],
            "Consumable issues and returns.", "Read only (posted)")
    CB = "Cons Balance"
    b.sheet(CB, "bal", "CONSUMABLE BALANCE (calculated)", "Stock = receipts + net issues + approved adjustments. Re-order flag from minimum level.",
            [("Item ID", "lk", 9, None), ("Item Name", "lk", 20, "=INDEX('M_Consumable'!$B$5:$B$50,MATCH(A{r},'M_Consumable'!$A$5:$A$50,0))"), ("Unit", "lk", 6, "=INDEX('M_Consumable'!$D$5:$D$50,MATCH(A{r},'M_Consumable'!$A$5:$A$50,0))"),
             ("Received", "ca", 8, f"=SUMIFS({rng(CR,'D')},{rng(CR,'C')},A{{r}})"), ("Net Issued", "ca", 8, f"=SUMIFS({rng(CI,'F')},{rng(CI,'D')},A{{r}})"),
             ("Adjusted", "ca", 8, f"=SUMIFS({rng(ADJ,'G')},{rng(ADJ,'D')},A{{r}})"), ("Balance", "to", 8, "=D{r}+E{r}+F{r}"),
             ("Avg Rate Rs", "ca", 8, f"=IFERROR(SUMIFS({rng(CR,'F')},{rng(CR,'C')},A{{r}})/D{{r}},0)"), ("Value Rs", "ca", 9, "=ROUND(G{r}*H{r},2)"),
             ("Min Level", "lk", 7, "=INDEX('M_Consumable'!$E$5:$E$50,MATCH(A{r},'M_Consumable'!$A$5:$A$50,0))"), ("Status", "cf", 14, "=IF(G{r}<0,\"NEGATIVE\",IF(G{r}<=J{r},\"REORDER\",\"OK\"))")],
            [["ABR-001"], ["ABR-004"], ["WLD-003"], ["GAS-001"]], "Consumable stock, value and re-order status.", "Stores (read only)", flagcol=11)
    CBR = lambda col: rng(CB, col, 5, 100)
    b.sheet("Cons Entry", "pad", "CONSUMABLE ISSUE / RETURN PAD", "One row per item handed out or returned. Issue above balance is blocked.",
            [("Cons Move ID", "ca", 14, f"=IF(C{{r}}=\"\",\"\",\"CI-\"&{FY}&\"-\"&TEXT(COUNTA({rng(CI,'A')})+ROW()-4,\"000000\"))"), ("Move Date", "in_", 9, None), ("Movement", "in_", 8, None),
             ("Item ID", "in_", 9, None), ("Item Name", "ca", 20, f"=IF(D{{r}}=\"\",\"\",IFERROR(INDEX({CBR('B')},MATCH(D{{r}},{CBR('A')},0)),\"?\"))"), ("Qty", "in_", 7, None),
             ("Balance Before", "lk", 9, f"=IF(D{{r}}=\"\",\"\",IFERROR(INDEX({CBR('G')},MATCH(D{{r}},{CBR('A')},0)),0)-SUMIFS(F$4:F{{p}},D$4:D{{p}},D{{r}},C$4:C{{p}},\"ISSUE\"))"),
             ("Avg Rate Rs", "ca", 8, f"=IF(D{{r}}=\"\",\"\",IFERROR(INDEX({CBR('H')},MATCH(D{{r}},{CBR('A')},0)),0))"), ("Value Rs", "ca", 9, "=IF(F{r}=\"\",\"\",ROUND(IF(C{r}=\"ISSUE\",-1,1)*F{r}*H{r},2))"),
             ("Issued To", "in_", 8, None), ("Where Used", "in_", 16, None),
             ("Ready / issues", "cf", 30, "=IF(C{r}=\"\",\"\",IF(E{r}=\"?\",\"ITEM NOT IN MASTER\",IF(AND(C{r}=\"ISSUE\",F{r}>G{r}),\"EXCEEDS BALANCE: max \"&G{r},IF(K{r}=\"\",\"WHERE USED REQUIRED\",\"READY\"))))")],
            [[None, D(2027, 4, 15), "ISSUE", "ABR-001", None, 30, None, None, None, "PER-007", "MC-01 SLTL Laser"],
             [None, D(2027, 4, 15), "ISSUE", "GAS-001", None, 40, None, None, None, "PER-030", "Fabrication Bay"],
             [None, D(2027, 4, 15), "ISSUE", "WLD-003", None, 10, None, None, None, "PER-030", ""],
             [None, D(2027, 4, 15), "RETURN", "ABR-004", None, 2, None, None, None, "PER-003", "MC-02 Amada Punching"]],
            "Issue / return consumables with balance hard stop.", "Stores", flagcol=12, table=False)
    b.sheet("Adjust Entry", "pad", "AUTHORISED STOCK-ADJUSTMENT PAD", "Only for genuine corrections. Reason code, reference and an approver different from the requester are mandatory.",
            [("Adj ID", "ca", 13, f"=IF(C{{r}}=\"\",\"\",\"ADJ-\"&{FY}&\"-\"&TEXT(COUNTA({rng(ADJ,'A')})+ROW()-4,\"00000\"))"), ("Adj Date", "in_", 9, None), ("Stock Type", "in_", 9, None), ("Lot / Item ID", "in_", 12, None),
             ("Direction", "in_", 9, None), ("Adj kg / Qty", "in_", 9, None), ("Available", "lk", 9, f"=IF(D{{r}}=\"\",\"\",IF(C{{r}}=\"RM-LOT\",IFERROR(INDEX({LBR('K')},MATCH(D{{r}},{LBR('A')},0)),0),IFERROR(INDEX({CBR('G')},MATCH(D{{r}},{CBR('A')},0)),0)))"),
             ("Rate Rs", "ca", 7, f"=IF(D{{r}}=\"\",\"\",IF(C{{r}}=\"RM-LOT\",IFERROR(INDEX({LBR('M')},MATCH(D{{r}},{LBR('A')},0)),0),IFERROR(INDEX({CBR('H')},MATCH(D{{r}},{CBR('A')},0)),0)))"),
             ("Value Rs", "ca", 9, "=IF(F{r}=\"\",\"\",ROUND(IF(E{r}=\"DECREASE\",-1,1)*F{r}*H{r},2))"), ("Reason Code", "in_", 9, None), ("Reference", "in_", 18, None), ("Requested By", "in_", 9, None), ("Approved By", "in_", 9, None),
             ("Ready / issues", "cf", 32, f"=IF(C{{r}}=\"\",\"\",IF(J{{r}}=\"\",\"REASON CODE REQUIRED\",IF(K{{r}}=\"\",\"REFERENCE REQUIRED\",IF(OR(M{{r}}=\"\",M{{r}}=L{{r}}),\"INDEPENDENT APPROVER REQUIRED\",IF(AND(E{{r}}=\"DECREASE\",F{{r}}>G{{r}}),\"EXCEEDS AVAILABLE\",IF(ABS(I{{r}})>'Settings'!$B$13,\"WARN: above limit - MD approval\",\"READY\"))))))")],
            [[None, D(2027, 4, 15), "RM-LOT", "GP-2728-0003", "DECREASE", 2.6, None, None, None, "ADJ-WGT", "Weighbridge slip WB-131", "PER-015", "PER-018"],
             [None, D(2027, 4, 15), "CONSUMABLE", "GAS-001", "INCREASE", 10, None, None, None, "", "", "PER-015", "PER-015"],
             [None, D(2027, 4, 15), "RM-LOT", "HR-2728-0001", "DECREASE", 30, None, None, None, "ADJ-DMG", "Damage note DN-07", "PER-015", ""]],
            "Approved corrections with reason and independent approver (no counts).", "Stores + approver", flagcol=14, table=False)
    b.sheet("Posting Log", "ctl", "POSTING LOG", "One line per posting: what was posted and the pad control total, so pad = register.",
            [("Posted At", "in_", 13, None), ("Register", "in_", 20, None), ("Rows Posted", "in_", 9, None), ("Pad Total kg / Qty", "in_", 12, None), ("Register Total After", "ca", 13, None), ("Posted By", "in_", 9, None), ("Method", "in_", 10, None)],
            [[dt.datetime(2027, 4, 13, 17, 12), MV, 2, 233.3, 1124.3, "PER-015", "Macro"], [dt.datetime(2027, 4, 14, 12, 1), ADJ, 1, 4.2, 4.2, "PER-015", "Manual"]],
            "Evidence of every posting.", "Book owner")
    b.sheet("Control Totals", "ctl", "CONTROL TOTALS - read by MIS after refresh", "Rows and totals per register; MIS compares its extracts with these.",
            [("Register", "lk", 24, None), ("Rows", "ca", 7, None), ("Total kg / Qty", "ca", 12, None), ("Total Value Rs", "ca", 13, None)],
            [["GRN Register", f"=COUNTA({rng(GRN,'A')})", f"=SUM({rng(GRN,'N')})", f"=SUM({rng(GRN,'T')})"], [MV, f"=COUNTA({rng(MV,'A')})", f"=SUM({rng(MV,'J')})", f"=SUM({rng(MV,'M')})"],
             [CI, f"=COUNTA({rng(CI,'A')})", f"=SUM({rng(CI,'F')})", f"=SUM({rng(CI,'H')})"], [ADJ, f"=COUNTA({rng(ADJ,'A')})", f"=SUM({rng(ADJ,'G')})", f"=SUM({rng(ADJ,'I')})"]],
            "Row counts and totals per register for reconciliation.", "MIS / Finance")
    lists(b, [("Movement", ["ISSUE", "RETURN-USABLE", "RETURN-OFFCUT", "RETURN-DAMAGED", "STORES-SCRAP", "RETURN-TO-SUPPLIER"]), ("Form", ["SHEET", "OFFCUT", "COIL", "STRIP"]),
              ("QC Status", ["ACCEPTED", "ACCEPTED WITH REJECTION", "HOLD", "REJECTED"]), ("Excess Decision", ["ACCEPT EXCESS", "RETURN EXCESS"]), ("Stock Type", ["RM-LOT", "CONSUMABLE"]), ("Direction", ["INCREASE", "DECREASE"])])
    order = ["README", "PO Entry", "GRN Entry", "RM Move Entry", "Cons Entry", "Adjust Entry", "PO Header Register", "PO Line Register", "GRN Register", "Lot QC Register", MV,
             CR, CI, ADJ, "PO Status", LB, CB, "M_Material", "M_Grade", "M_Supplier", "M_Consumable", "M_Reason", "M_WhereUsed", "X_Product", "X_Machine", "X_People",
             "Posting Log", "Control Totals", "Settings", "Lists"]
    return b, order


# ============================================================ 2 PRODUCTION
def production():
    b = Book("Production", "2 ISKO Production FY2027-28 (prototype).xlsx")
    readme(b, "2 ISKO PRODUCTION FY2027-28 - cutting, bending, fabrication, rejection / rework, batch balance, WIP", [
        ("Daily routine", ""), ("Start of shift", "Refresh All: X_Batch (open batches from Stores). Check Batch Balance is current."),
        ("Through the shift", "Type in Prod Entry / Rej-Rework Entry. Post READY rows. A cutting line can never use more than the batch has left."),
        ("Batch finished", "Post a Batch Close row (variance above tolerance needs reason + approver)."), ("Never", "Count WIP - WIP is calculated from postings.")])
    settings(b, [["FY", "2027-28", ""], ["FYCode", "2728", ""], ["AsOnDate", D(2027, 4, 15), "Typed"], ["LockDate", D(2027, 3, 31), ""], ["BatchTol %", 0.02, "Batch variance tolerance"],
                 ["BatchTolMinKg", 1, "Minimum tolerance in kg"], ["DataFolder", "D:\\ISKO ERP\\FY2027-28\\", ""], ["StoresFile", "1 ISKO Stores FY2027-28.xlsx", "Source of X_Batch"]])
    FY = "'Settings'!$B$6"
    XB, PR, RJ, BC = "X_Batch", "Production Register", "Rej-Rework Register", "Batch Close Register"
    b.sheet(XB, "imp", "X_BATCH - open batches from Stores (values)", "Power Query extract of Stores ISSUE lines with returns and stores scrap per batch.",
            [("Batch ID", "lk", 15, None), ("Issue Date", "lk", 9, None), ("Machine ID", "lk", 8, None), ("Lot ID", "lk", 12, None), ("Material ID", "lk", 9, None), ("Issued kg", "lk", 9, None),
             ("Issued Sh", "lk", 7, None), ("Returned kg", "lk", 9, None), ("Stores Scrap kg", "lk", 9, None), ("Rate Rs", "lk", 7, None), ("Run ID", "lk", 15, None)],
            [["ISS-2728-000001", D(2027, 4, 12), "MC-01", "GP-2728-0001", "GPSP-1.6", 471, 12, 18.4, 0, 74.5, "RUN-270415-0800"],
             ["ISS-2728-000002", D(2027, 4, 12), "MC-03", "GP-2728-0002", "GPSP-1.2", 420, 15, 0, 0, 74.5, "RUN-270415-0800"],
             ["ISS-2728-000004", D(2027, 4, 13), "MC-02", "HR-2728-0001", "HRCA-2.0", 214.9, 3, 0, 0, 66.1, "RUN-270415-0800"]],
            "Open batches issued by Stores, with returns.", "Read only")
    b.sheet("X_Material", "imp", "X_MATERIAL - from Stores (values)", "Power Query extract.",
            [("Material ID", "lk", 10, None), ("Grade", "lk", 8, None), ("Thk mm", "lk", 7, None), ("Active", "lk", 7, None)],
            [["GPSP-1.6", "GPSP", 1.6, "Yes"], ["GPSP-1.2", "GPSP", 1.2, "Yes"], ["CRCA-1.0", "CRCA", 1.0, "Yes"], ["HRCA-2.0", "HRCA", 2.0, "Yes"]], "Material list for parts and BOM.", "Read only")
    b.sheet("X_Customer", "imp", "X_CUSTOMER - from Dispatch (values)", "Power Query extract.",
            [("Customer ID", "lk", 10, None), ("Customer", "lk", 18, None), ("Active", "lk", 7, None)], [["CUS-01", "Customer 1 (masked)", "Yes"], ["CUS-02", "Customer 2 (masked)", "Yes"]],
            "Customer list for the product master.", "Read only")
    b.sheet("M_Product", "mst", "PRODUCT MASTER - single master for all four books", "Product ID never changes; Stores, Dispatch and MIS import this sheet.",
            [("Product ID", "in_", 10, None), ("Product Name", "in_", 18, None), ("Customer ID", "in_", 10, None), ("Category", "in_", 10, None), ("BOM Model", "in_", 12, None),
             ("Coating Route", "in_", 13, None), ("Active", "in_", 7, None), ("Effective From", "in_", 10, None)],
            [["PRD-A12", "Model A12 ODU", "CUS-01", "ODU", "A12-ED2026B", "POWDER COAT", "Yes", D(2027, 4, 1)], ["PRD-B07", "Model B07 IDU", "CUS-01", "IDU", "B07-ED2026B", "POWDER COAT", "Yes", D(2027, 4, 1)],
             ["PRD-C03", "Model C03 frame", "CUS-02", "FRAME", "C03-ED2026A", "NO COATING", "Yes", D(2027, 4, 1)]],
            "Products / models with customer and coating route.", "Production engineering")
    MP = "M_Part"
    b.sheet(MP, "mst", "PART MASTER", "One row per Part ID and revision. Wt/pc is copied onto production lines at posting.",
            [("Part ID", "in_", 11, None), ("Product ID", "in_", 9, None), ("Part Name", "in_", 14, None), ("Drawing No", "in_", 10, None), ("Rev", "in_", 5, None), ("Material ID", "in_", 9, None),
             ("Thk mm", "in_", 6, None), ("Wt/pc kg", "in_", 8, None), ("Sq ft/pc", "in_", 7, None), ("Route", "in_", 7, None), ("Active", "in_", 6, None)],
            [["A12-BP-01", "PRD-A12", "Base panel", "DRG-1102", "C", "GPSP-1.6", 1.6, 6.215, 4.15, "R-CBF", "Yes"], ["A12-SP-04", "PRD-A12", "Side panel", "DRG-1107", "A", "GPSP-1.6", 1.6, 2.01, 1.6, "R-CB", "Yes"],
             ["B07-CB-02", "PRD-B07", "Coil bracket", "DRG-2210", "B", "GPSP-1.2", 1.2, 1.15, 0.9, "R-CB", "Yes"], ["C03-FR-01", "PRD-C03", "Frame", "DRG-3301", "A", "HRCA-2.0", 2.0, 12.4, 6.2, "R-CF", "Yes"]],
            "Drawing, revision, material, weight per piece, coating area, route.", "Production engineering")
    b.sheet("M_BOM_Model", "mst", "BOM EDITIONS", "New edition = new rows; old edition kept.",
            [("BOM Model", "in_", 13, None), ("Product ID", "in_", 9, None), ("Edition", "in_", 10, None), ("Effective From", "in_", 10, None), ("Review Due", "in_", 10, None), ("Set Wt kg", "ca", 9, None)],
            [["A12-ED2026B", "PRD-A12", "2026-B", D(2026, 10, 1), D(2027, 4, 1), 42.663], ["B07-ED2026B", "PRD-B07", "2026-B", D(2026, 10, 1), D(2027, 4, 1), 18.2], ["C03-ED2026A", "PRD-C03", "2026-A", D(2026, 4, 1), D(2026, 10, 1), 74.24]],
            "BOM edition headers.", "Production engineering")
    b.sheet("M_BOM", "mst", "BOM LINES", "Wt per set = blank weight x qty per set.",
            [("BOM Model", "in_", 13, None), ("SR No", "in_", 5, None), ("Parent SR", "in_", 6, None), ("Part ID", "in_", 10, None), ("Material ID", "lk", 9, None), ("Blank Wt kg", "in_", 9, None),
             ("Qty/Set", "in_", 6, None), ("Wt per Set kg", "ca", 9, "=ROUND(F{r}*G{r},3)"), ("Costable", "in_", 7, None)],
            [["A12-ED2026B", 1, "", "A12-BP-01", "GPSP-1.6", 6.215, 1, None, "Yes"], ["A12-ED2026B", 2, "", "A12-SP-04", "GPSP-1.6", 2.01, 2, None, "Yes"],
             ["B07-ED2026B", 1, "", "B07-CB-02", "GPSP-1.2", 1.15, 4, None, "Yes"], ["C03-ED2026A", 1, "", "C03-FR-01", "HRCA-2.0", 12.4, 2, None, "Yes"]],
            "BOM lines (drives MIS RM requirement).", "Production engineering")
    b.sheet("M_Stage_Route", "mst", "STAGE & ROUTE MASTER", "Stage sequence per route - this is what makes WIP calculable without counts.",
            [("Route", "in_", 8, None), ("Seq", "in_", 5, None), ("Stage", "in_", 13, None), ("Next Stage", "in_", 13, None), ("Ends In", "in_", 18, None)],
            [["R-CBF", 1, "CUTTING", "BENDING", ""], ["R-CBF", 2, "BENDING", "FABRICATION", ""], ["R-CBF", 3, "FABRICATION", "JWC", "Powder coating (JWC)"],
             ["R-CB", 1, "CUTTING", "BENDING", ""], ["R-CB", 2, "BENDING", "JWC", "Powder coating (JWC)"], ["R-CF", 1, "CUTTING", "FABRICATION", ""], ["R-CF", 2, "FABRICATION", "FG", "Finished goods (no coating)"]],
            "Stage order per part route; final stage and FG point.", "Production head")
    b.sheet("M_Machine", "mst", "MACHINE MASTER", "Process decides the stage of every production line.",
            [("Machine ID", "in_", 9, None), ("Machine", "in_", 16, None), ("Process", "in_", 12, None), ("Status", "in_", 9, None), ("Capacity kg/day (Phase 2)", "in_", 12, None)],
            [["MC-01", "SLTL Laser", "CUTTING", "Running", ""], ["MC-02", "Amada Punching", "CUTTING", "Running", ""], ["MC-03", "Divine Laser", "CUTTING", "Running", ""], ["MC-05", "Amada Bending", "BENDING", "Running", ""], ["MC-07", "Fabrication Bay", "FABRICATION", "Running", ""]],
            "Machines and their process.", "Production head")
    b.sheet("M_People", "mst", "PEOPLE MASTER - single list for all books", "Person ID; never delete, set Active = No.",
            [("Person ID", "in_", 9, None), ("Name (masked)", "in_", 14, None), ("Role", "in_", 13, None), ("Department", "in_", 12, None), ("Active", "in_", 6, None)],
            [["PER-003", "Person 03", "Operator", "Production", "Yes"], ["PER-007", "Person 07", "Operator", "Production", "Yes"], ["PER-011", "Person 11", "Operator", "Production", "Yes"],
             ["PER-015", "Person 15", "Stores", "Stores", "Yes"], ["PER-018", "Person 18", "QC / approver", "QC", "Yes"], ["PER-021", "Person 21", "Operator", "Production", "Yes"], ["PER-030", "Person 30", "Supervisor", "Production", "Yes"]],
            "Operators, supervisors, QC, stores, approvers.", "HR / Production")
    b.sheet("M_Reason", "mst", "REJECTION REASONS", "Reason group + reason codes.",
            [("Code", "in_", 9, None), ("Reason Group", "in_", 22, None), ("Reason", "in_", 30, None)],
            [["BND-DIM", "BENDING - DIMENSION / POSITION", "Wrong bending position"], ["CUT-MAT", "CUTTING / MATERIAL", "Cut in wrong thickness"], ["FAB-WLD", "FABRICATION - WELDING", "Weld spatter / burn through"]],
            "Rejection reasons.", "QC")
    b.sheet(PR, "reg", "PRODUCTION REGISTER - values only", "Posted lines never recalculate. Wt/pc, drawing revision and lot are frozen at posting.",
            [("Prod ID", "ca", 15, None), ("Prod Date", "ca", 9, None), ("Shift", "ca", 5, None), ("Machine ID", "ca", 8, None), ("Stage", "ca", 11, None), ("Batch ID", "ca", 15, None), ("Lot ID", "ca", 12, None),
             ("Product ID", "ca", 8, None), ("Part ID", "ca", 10, None), ("Drawing / Rev", "ca", 10, None), ("Qty", "ca", 5, None), ("Wt/pc kg", "ca", 7, None), ("Output kg", "ca", 9, None),
             ("Scrap kg", "ca", 7, None), ("Operator", "ca", 8, None), ("QC", "ca", 5, None), ("Dup Key", "ca", 30, None), ("Posted At", "ca", 12, None)],
            [["PRD-2728-000001", D(2027, 4, 12), "A", "MC-01", "CUTTING", "ISS-2728-000001", "GP-2728-0001", "PRD-A12", "A12-BP-01", "DRG-1102 / C", 60, 6.215, 372.9, 24.3, "PER-021", "OK", "46489|A|MC-01|ISS-2728-000001|A12-BP-01|60", dt.datetime(2027, 4, 12, 18, 0)],
             ["PRD-2728-000002", D(2027, 4, 12), "A", "MC-03", "CUTTING", "ISS-2728-000002", "GP-2728-0002", "PRD-B07", "B07-CB-02", "DRG-2210 / B", 300, 1.15, 345, 30, "PER-011", "OK", "46489|A|MC-03|ISS-2728-000002|B07-CB-02|300", dt.datetime(2027, 4, 12, 18, 0)],
             ["PRD-2728-000003", D(2027, 4, 13), "B", "MC-05", "BENDING", "", "", "PRD-A12", "A12-BP-01", "DRG-1102 / C", 60, 6.215, 372.9, 0, "PER-030", "OK", "46490|B|MC-05||A12-BP-01|60", dt.datetime(2027, 4, 13, 18, 0)],
             ["PRD-2728-000004", D(2027, 4, 14), "A", "MC-07", "FABRICATION", "", "", "PRD-A12", "A12-BP-01", "DRG-1102 / C", 50, 6.215, 310.75, 0, "PER-030", "OK", "46491|A|MC-07||A12-BP-01|50", dt.datetime(2027, 4, 14, 18, 0)],
             ["PRD-2728-000005", D(2027, 4, 13), "A", "MC-02", "CUTTING", "ISS-2728-000004", "HR-2728-0001", "PRD-C03", "C03-FR-01", "DRG-3301 / A", 15, 12.4, 186, 20.1, "PER-003", "OK", "46490|A|MC-02|ISS-2728-000004|C03-FR-01|15", dt.datetime(2027, 4, 13, 18, 0)],
             ["PRD-2728-000006", D(2027, 4, 14), "B", "MC-05", "BENDING", "", "", "PRD-B07", "B07-CB-02", "DRG-2210 / B", 280, 1.15, 322, 0, "PER-030", "OK", "46491|B|MC-05||B07-CB-02|280", dt.datetime(2027, 4, 14, 18, 0)]],
            "All production lines (cutting, bending, fabrication).", "Read only (posted)")
    b.sheet(RJ, "reg", "REJECTION & REWORK REGISTER - event lines", "REJECT opens a case on the source Prod ID; REWORK-OK returns qty to WIP; REWORK-SCRAP closes it as scrap.",
            [("Event ID", "ca", 14, None), ("Event Date", "ca", 9, None), ("Event Type", "ca", 12, None), ("Case ID", "ca", 14, None), ("Source Prod ID", "ca", 15, None), ("Stage", "ca", 10, None),
             ("Part ID", "ca", 10, None), ("Qty", "ca", 5, None), ("kg", "ca", 7, None), ("Reason Code", "ca", 9, None), ("Responsibility", "ca", 10, None), ("Disposition", "ca", 10, None), ("Posted At", "ca", 12, None)],
            [["REJ-2728-00001", D(2027, 4, 13), "REJECT", "REJ-2728-00001", "PRD-2728-000003", "BENDING", "A12-BP-01", 4, 24.86, "BND-DIM", "IN-HOUSE", "REWORK", dt.datetime(2027, 4, 13, 18, 30)],
             ["RWK-2728-00002", D(2027, 4, 14), "REWORK-OK", "REJ-2728-00001", "PRD-2728-000003", "BENDING", "A12-BP-01", 3, 18.645, "", "", "", dt.datetime(2027, 4, 14, 17, 0)],
             ["RWK-2728-00003", D(2027, 4, 14), "REWORK-SCRAP", "REJ-2728-00001", "PRD-2728-000003", "BENDING", "A12-BP-01", 1, 6.215, "", "", "SCRAP", dt.datetime(2027, 4, 14, 17, 0)]],
            "Rejection cases and rework outcomes, linked to the original production line.", "Read only (posted)")
    b.sheet(BC, "reg", "BATCH CLOSE REGISTER", "A batch closes only here. Variance above tolerance needs a reason and an approver.",
            [("Batch ID", "ca", 15, None), ("Close Date", "ca", 9, None), ("Net Batch kg", "ca", 9, None), ("Output kg", "ca", 9, None), ("Scrap kg", "ca", 8, None), ("Variance kg", "ca", 8, None),
             ("Var %", "ca", 6, None), ("Variance Reason", "ca", 16, None), ("Closed By", "ca", 8, None), ("Approved By", "ca", 8, None)],
            [], "Batch closures with frozen variance.", "Production + Stores")
    b.wb[BC].cell(5, 1, "(no batch closed yet in this sample)").font = Font(italic=True, color="5A6B82", size=9)
    BB = "Batch Balance"
    b.sheet(BB, "bal", "BATCH BALANCE - open batches (calculated)", "kg left = issued - returned - cut output - cut scrap - stores scrap. Cutting entries are blocked beyond kg left + tolerance.",
            [("Batch ID", "lk", 15, None), ("Issue Date", "lk", 9, f"=INDEX({rng(XB,'B')},MATCH(A{{r}},{rng(XB,'A')},0))"), ("Machine", "lk", 7, f"=INDEX({rng(XB,'C')},MATCH(A{{r}},{rng(XB,'A')},0))"),
             ("Lot ID", "lk", 12, f"=INDEX({rng(XB,'D')},MATCH(A{{r}},{rng(XB,'A')},0))"), ("Issued kg", "lk", 8, f"=INDEX({rng(XB,'F')},MATCH(A{{r}},{rng(XB,'A')},0))"),
             ("Returned kg", "lk", 8, f"=INDEX({rng(XB,'H')},MATCH(A{{r}},{rng(XB,'A')},0))"), ("Net kg", "ca", 8, "=E{r}-F{r}"),
             ("Cut Output kg", "ca", 9, f"=SUMIFS({rng(PR,'M')},{rng(PR,'F')},A{{r}},{rng(PR,'E')},\"CUTTING\")"), ("Cut Scrap kg", "ca", 8, f"=SUMIFS({rng(PR,'N')},{rng(PR,'F')},A{{r}},{rng(PR,'E')},\"CUTTING\")"),
             ("Stores Scrap kg", "lk", 8, f"=INDEX({rng(XB,'I')},MATCH(A{{r}},{rng(XB,'A')},0))"), ("kg Left", "to", 8, "=ROUND(G{r}-H{r}-I{r}-J{r},3)"), ("Var %", "ca", 6, "=IF(G{r}=0,0,K{r}/G{r})"),
             ("Variance Rs", "ca", 9, f"=ROUND(K{{r}}*INDEX({rng(XB,'J')},MATCH(A{{r}},{rng(XB,'A')},0)),2)"), ("Lines", "ca", 5, f"=COUNTIFS({rng(PR,'F')},A{{r}})"),
             ("Days Open", "ca", 6, "='Settings'!$B$7-B{r}"),
             ("Status", "cf", 18, f"=IF(COUNTIF({rng(BC,'A')},A{{r}})>0,\"CLOSED\",IF(N{{r}}=0,\"NOT STARTED\",IF(K{{r}}<-MAX('Settings'!$B$10,'Settings'!$B$9*G{{r}}),\"OVER-CONSUMED\",IF(ABS(K{{r}})<=MAX('Settings'!$B$10,'Settings'!$B$9*G{{r}}),\"READY TO CLOSE\",\"IN PROGRESS\"))))")],
            [["ISS-2728-000001"], ["ISS-2728-000002"], ["ISS-2728-000004"]],
            "Issued, returned, output, scrap, kg left, variance and status per batch.", "Production (read only)", flagcol=16)
    BBR = lambda col: rng(BB, col, 5, 100)
    b.sheet("Prod Entry", "pad", "PRODUCTION ENTRY PAD", "Cutting lines must quote an open batch of the machine. Output kg = qty x Wt/pc from the part master (frozen at posting).",
            [("Prod ID", "ca", 15, f"=IF(D{{r}}=\"\",\"\",\"PRD-\"&{FY}&\"-\"&TEXT(COUNTA({rng(PR,'A')})+ROW()-4,\"000000\"))"), ("Prod Date", "in_", 9, None), ("Shift", "in_", 5, None), ("Machine ID", "in_", 8, None),
             ("Stage", "ca", 11, "=IF(D{r}=\"\",\"\",INDEX('M_Machine'!$C$5:$C$50,MATCH(D{r},'M_Machine'!$A$5:$A$50,0)))"), ("Batch ID", "in_", 15, None),
             ("Lot ID", "ca", 12, f"=IF(F{{r}}=\"\",\"\",IFERROR(INDEX({rng(XB,'D')},MATCH(F{{r}},{rng(XB,'A')},0)),\"NOT OPEN\"))"),
             ("Part ID", "in_", 10, None), ("Product ID", "ca", 8, "=IF(H{r}=\"\",\"\",INDEX('M_Part'!$B$5:$B$50,MATCH(H{r},'M_Part'!$A$5:$A$50,0)))"),
             ("Drawing / Rev", "ca", 10, "=IF(H{r}=\"\",\"\",INDEX('M_Part'!$D$5:$D$50,MATCH(H{r},'M_Part'!$A$5:$A$50,0))&\" / \"&INDEX('M_Part'!$E$5:$E$50,MATCH(H{r},'M_Part'!$A$5:$A$50,0)))"),
             ("Qty", "in_", 5, None), ("Wt/pc kg", "ca", 7, "=IF(H{r}=\"\",\"\",INDEX('M_Part'!$H$5:$H$50,MATCH(H{r},'M_Part'!$A$5:$A$50,0)))"), ("Output kg", "ca", 9, "=IF(K{r}=\"\",\"\",ROUND(K{r}*L{r},3))"),
             ("Scrap kg", "in_", 7, None),
             ("Batch kg Left", "lk", 9, f"=IF(F{{r}}=\"\",\"\",IFERROR(INDEX({BBR('K')},MATCH(F{{r}},{BBR('A')},0)),0)-SUMIFS(M$4:M{{p}},F$4:F{{p}},F{{r}})-SUMIFS(N$4:N{{p}},F$4:F{{p}},F{{r}}))"),
             ("Operator", "in_", 8, None), ("Dup Key", "ca", 28, "=IF(D{r}=\"\",\"\",B{r}&\"|\"&C{r}&\"|\"&D{r}&\"|\"&F{r}&\"|\"&H{r}&\"|\"&K{r})"),
             ("Ready / issues", "cf", 36, f"=IF(D{{r}}=\"\",\"\",IF(AND(E{{r}}=\"CUTTING\",F{{r}}=\"\"),\"CUTTING NEEDS A BATCH\",IF(G{{r}}=\"NOT OPEN\",\"BATCH NOT OPEN\",IF(AND(E{{r}}=\"CUTTING\",N(M{{r}})+N(N{{r}})>N(O{{r}})+MAX('Settings'!$B$10,'Settings'!$B$9*N(O{{r}}))),\"OVER BATCH by \"&TEXT(N(M{{r}})+N(N{{r}})-N(O{{r}}),\"0.0\")&\" kg\",IF(COUNTIF({rng(PR,'Q')},Q{{r}})>0,\"WARN: POSSIBLE DUPLICATE - same key posted\",\"READY\")))))")],
            [[None, D(2027, 4, 15), "A", "MC-01", None, "ISS-2728-000001", None, "A12-SP-04", None, None, 20, None, None, 3.0, None, "PER-021"],
             [None, D(2027, 4, 15), "A", "MC-01", None, "ISS-2728-000001", None, "A12-SP-04", None, None, 10, None, None, 1.5, None, "PER-021"],
             [None, D(2027, 4, 13), "B", "MC-05", None, "", None, "A12-BP-01", None, None, 60, None, None, 0, None, "PER-030"],
             [None, D(2027, 4, 15), "A", "MC-03", None, "", None, "B07-CB-02", None, None, 40, None, None, 2, None, "PER-011"],
             [None, D(2027, 4, 15), "B", "MC-07", None, "", None, "B07-CB-02", None, None, 100, None, None, 0, None, "PER-030"]],
            "Daily production lines with batch hard stop and duplicate check.", "Production supervisors", flagcol=18, table=False)
    b.sheet("Rej-Rework Entry", "pad", "REJECTION & REWORK PAD", "Quote the production line the parts came from. Rejected qty cannot exceed what that line made minus earlier rejections.",
            [("Event ID", "ca", 14, f"=IF(C{{r}}=\"\",\"\",IF(C{{r}}=\"REJECT\",\"REJ\",\"RWK\")&\"-\"&{FY}&\"-\"&TEXT(COUNTA({rng(RJ,'A')})+ROW()-4,\"00000\"))"), ("Event Date", "in_", 9, None),
             ("Event Type", "in_", 12, None), ("Case ID", "in_", 14, None), ("Source Prod ID", "in_", 15, None),
             ("Part ID", "ca", 10, f"=IF(E{{r}}=\"\",\"\",IFERROR(INDEX({rng(PR,'I')},MATCH(E{{r}},{rng(PR,'A')},0)),\"NOT FOUND\"))"),
             ("Made Qty", "lk", 7, f"=IF(E{{r}}=\"\",\"\",IFERROR(INDEX({rng(PR,'K')},MATCH(E{{r}},{rng(PR,'A')},0)),0))"),
             ("Already Rejected", "lk", 8, f"=IF(E{{r}}=\"\",\"\",SUMIFS({rng(RJ,'H')},{rng(RJ,'E')},E{{r}},{rng(RJ,'C')},\"REJECT\"))"), ("Qty", "in_", 5, None),
             ("Reason Code", "in_", 9, None), ("Disposition", "in_", 10, None),
             ("Ready / issues", "cf", 34, "=IF(C{r}=\"\",\"\",IF(F{r}=\"NOT FOUND\",\"SOURCE PROD ID NOT FOUND\",IF(AND(C{r}=\"REJECT\",I{r}>G{r}-H{r}),\"MORE THAN MADE: max \"&(G{r}-H{r}),IF(AND(C{r}=\"REJECT\",J{r}=\"\"),\"REASON REQUIRED\",\"READY\"))))")],
            [[None, D(2027, 4, 15), "REJECT", "", "PRD-2728-000002", None, None, None, 6, "CUT-MAT", "SCRAP"],
             [None, D(2027, 4, 15), "REJECT", "", "PRD-2728-000005", None, None, None, 20, "CUT-MAT", "REWORK"],
             [None, D(2027, 4, 15), "REJECT", "", "PRD-2728-000099", None, None, None, 2, "BND-DIM", "REWORK"]],
            "Record rejections against the source line and rework outcomes.", "QC / Production", flagcol=12, table=False)
    # WIP
    WP = "WIP Balance"
    rows = [["A12-BP-01", "CUTTING"], ["A12-BP-01", "BENDING"], ["A12-BP-01", "FABRICATION"], ["B07-CB-02", "CUTTING"], ["B07-CB-02", "BENDING"], ["C03-FR-01", "CUTTING"]]
    b.sheet("X_JWCSent", "imp", "X_JWC SENT / FG TRANSFER - from Dispatch (values)", "Quantities that left WIP to a coater (JWC) by part.",
            [("Part ID", "lk", 10, None), ("Sent on JWC Qty", "lk", 10, None), ("Run ID", "lk", 15, None)], [["A12-BP-01", 40, "RUN-270415-0800"], ["B07-CB-02", 100, "RUN-270415-0800"]],
            "JWC quantities by part (moves WIP out of the last stage).", "Read only")
    b.sheet(WP, "bal", "WIP BALANCE - transaction-calculated (no counts)", "WIP = completed here - completed at next stage (or sent on JWC / to FG) - rejected here + rework OK.",
            [("Part ID", "lk", 10, None), ("Stage", "lk", 11, None), ("Route", "lk", 6, "=INDEX('M_Part'!$J$5:$J$50,MATCH(A{r},'M_Part'!$A$5:$A$50,0))"),
             ("Next Stage", "lk", 11, f"ARR:=INDEX('M_Stage_Route'!$D$5:$D$50,MATCH(1,('M_Stage_Route'!$A$5:$A$50=C{{r}})*('M_Stage_Route'!$C$5:$C$50=B{{r}}),0))"),
             ("Qty In", "ca", 7, f"=SUMIFS({rng(PR,'K')},{rng(PR,'I')},A{{r}},{rng(PR,'E')},B{{r}})"),
             ("Qty Moved On", "ca", 9, f"=IF(D{{r}}=\"JWC\",SUMIFS('X_JWCSent'!$B$5:$B$50,'X_JWCSent'!$A$5:$A$50,A{{r}}),SUMIFS({rng(PR,'K')},{rng(PR,'I')},A{{r}},{rng(PR,'E')},D{{r}}))"),
             ("Rejected (net)", "ca", 8, f"=SUMIFS({rng(RJ,'H')},{rng(RJ,'G')},A{{r}},{rng(RJ,'F')},B{{r}},{rng(RJ,'C')},\"REJECT\")-SUMIFS({rng(RJ,'H')},{rng(RJ,'G')},A{{r}},{rng(RJ,'F')},B{{r}},{rng(RJ,'C')},\"REWORK-OK\")"),
             ("WIP Qty", "to", 7, "=E{r}-F{r}-G{r}"), ("WIP kg", "ca", 8, "=ROUND(H{r}*INDEX('M_Part'!$H$5:$H$50,MATCH(A{r},'M_Part'!$A$5:$A$50,0)),3)"),
             ("Status", "cf", 22, "=IF(H{r}<0,\"BREACH - route / posting error\",IF(H{r}=0,\"OK\",\"IN WIP\"))")],
            rows, "WIP by part and stage from postings only.", "Production (read only)", flagcol=10)
    b.sheet("Daily Output", "rpt", "DAILY OUTPUT - pick a date", "Simple summary by machine for one day (PivotTable in the real build).",
            [("Machine ID", "lk", 9, None), ("Stage", "lk", 12, "=INDEX('M_Machine'!$C$5:$C$50,MATCH(A{r},'M_Machine'!$A$5:$A$50,0))"),
             ("Output kg", "ca", 10, f"=SUMIFS({rng(PR,'M')},{rng(PR,'D')},A{{r}},{rng(PR,'B')},$B$3)"), ("Scrap kg", "ca", 8, f"=SUMIFS({rng(PR,'N')},{rng(PR,'D')},A{{r}},{rng(PR,'B')},$B$3)"),
             ("Lines", "ca", 6, f"=COUNTIFS({rng(PR,'D')},A{{r}},{rng(PR,'B')},$B$3)")],
            [["MC-01"], ["MC-02"], ["MC-03"], ["MC-05"], ["MC-07"]], "Output by machine for a chosen date.", "Production head",
            pre=[[("Date", "h2"), (D(2027, 4, 12), "in_")]], hdr=5, table=False)
    b.wb["Daily Output"]["B3"].number_format = "dd-mmm-yy"
    b.sheet("Control Totals", "ctl", "CONTROL TOTALS - read by MIS", "Rows and totals per register.",
            [("Register", "lk", 24, None), ("Rows", "ca", 7, None), ("Total Qty", "ca", 10, None), ("Total kg", "ca", 12, None)],
            [[PR, f"=COUNTA({rng(PR,'A')})", f"=SUM({rng(PR,'K')})", f"=SUM({rng(PR,'M')})"], [RJ, f"=COUNTA({rng(RJ,'A')})", f"=SUM({rng(RJ,'H')})", f"=SUM({rng(RJ,'I')})"]],
            "Reconciliation totals.", "MIS / Finance")
    b.sheet("Refresh Log", "ctl", "REFRESH LOG", "One row per import per refresh.",
            [("Run ID", "lk", 15, None), ("Refreshed At", "lk", 13, None), ("Source File", "lk", 26, None), ("Source Saved", "lk", 13, None), ("Rows", "lk", 6, None), ("Distinct IDs", "lk", 8, None), ("Control kg", "lk", 10, None), ("Status", "cf", 10, None)],
            [["RUN-270415-0800", dt.datetime(2027, 4, 15, 8, 0), "1 ISKO Stores FY2027-28.xlsx", dt.datetime(2027, 4, 14, 18, 5), 3, 3, 1105.9, "OK"]],
            "Proof that imports are fresh and complete.", "Book owner", flagcol=8)
    lists(b, [("Shift", ["A", "B", "C", "GENERAL"]), ("QC", ["OK", "HOLD"]), ("Event Type", ["REJECT", "REWORK-OK", "REWORK-SCRAP"]), ("Disposition", ["REWORK", "SCRAP", "RETURN-TO-STORES", "USE-AS-IS"])])
    order = ["README", "Prod Entry", "Rej-Rework Entry", PR, RJ, BC, BB, WP, "Daily Output", XB, "X_Material", "X_Customer", "X_JWCSent", "M_Product", MP, "M_BOM_Model", "M_BOM",
             "M_Stage_Route", "M_Machine", "M_People", "M_Reason", "Control Totals", "Refresh Log", "Settings", "Lists"]
    return b, order


# ============================================================== 3 DISPATCH
def dispatch():
    b = Book("Dispatch", "3 ISKO Dispatch & Challan FY2027-28 (prototype).xlsx")
    readme(b, "3 ISKO DISPATCH & CHALLAN FY2027-28 - requirements, job work, coaters, finished goods, delivery", [
        ("Daily routine", ""), ("Start of shift", "Refresh All: X_ProdOutput, X_Product. Check Open JWC Balance and FG Balance status."),
        ("Job work", "JWC Entry: header + lines; post. Coater returns on Coater Move Entry."),
        ("Dispatch", "DC Entry: header + lines; a line can never exceed FG available (factory) or JWC pending (coater)."),
        ("Note", "Statutory challan documents are produced outside this book (Decision D-03). No print or mail here.")])
    settings(b, [["FY", "2027-28", ""], ["FYCode", "2728", ""], ["AsOnDate", D(2027, 4, 15), "Typed"], ["LockDate", D(2027, 3, 31), ""], ["DCPrefix", "DC/27-28/", ""],
                 ["JWCPrefix", "JWC/27-28/", ""], ["CoaterAge days", 15, "Ageing warning"], ["ProductionFile", "2 ISKO Production FY2027-28.xlsx", "Source of X_ProdOutput"]])
    AS = "'Settings'!$B$7"
    JH, JL, CM, DH, DL = "JWC Header Register", "JWC Line Register", "Coater Move Register", "DC Header Register", "DC Line Register"
    b.sheet("M_Customer", "mst", "CUSTOMERS", "Dispatch UOM is mandatory. Email = plain reference only.",
            [("Customer ID", "in_", 10, None), ("Customer Name", "in_", 20, None), ("State Code", "in_", 7, None), ("GSTIN", "in_", 18, None), ("Tally Ledger", "in_", 18, None), ("Dispatch UOM", "in_", 9, None), ("Active", "in_", 6, None)],
            [["CUS-01", "Customer 1 (masked)", "27", "27XXXXX0000X1Z5", "CUSTOMER 1 LTD", "QTY", "Yes"], ["CUS-02", "Customer 2 (masked)", "24", "24XXXXX0000X1Z8", "CUSTOMER 2", "KG", "Yes"]],
            "Customers with GSTIN, ledger, default dispatch UOM.", "Dispatch / Accounts")
    b.sheet("Customer Price & UOM", "mst", "CUSTOMER PRICE & UOM - effective-dated", "A new rate is a new row with a new Effective From; posted DC lines keep their rate.",
            [("Customer ID", "in_", 10, None), ("Product ID", "in_", 9, None), ("Part ID", "in_", 10, None), ("UOM", "in_", 6, None), ("Wt per Unit kg", "in_", 10, None), ("Rate Rs/kg", "in_", 9, None), ("Effective From", "in_", 10, None)],
            [["CUS-01", "PRD-A12", "A12-BP-01", "QTY", 6.215, 173.24, D(2027, 4, 1)], ["CUS-01", "PRD-B07", "B07-CB-02", "QTY", 1.15, 165.1, D(2027, 4, 1)], ["CUS-02", "PRD-C03", "C03-FR-01", "KG", 12.4, 92.5, D(2027, 4, 1)]],
            "Sale rate, unit and weight per unit per customer-product.", "Dispatch / Accounts")
    b.sheet("M_Vendor", "mst", "VENDORS / COATERS", "Identity only; rates are in Coater Rate.",
            [("Vendor ID", "in_", 9, None), ("Vendor Name", "in_", 18, None), ("GSTIN", "in_", 18, None), ("Status", "in_", 8, None)],
            [["VEN-001", "Coater 1 (masked)", "24XXXXX0000X1Z3", "Active"], ["VEN-002", "Coater 2 (masked)", "24XXXXX0000X1Z4", "Active"]], "Powder coaters / sub-contractors.", "Dispatch")
    b.sheet("Coater Rate", "mst", "COATER RATE - effective-dated", "JWC lines copy the rate in force on the JWC date.",
            [("Vendor ID", "in_", 9, None), ("Basis", "in_", 7, None), ("Rate Rs", "in_", 8, None), ("Effective From", "in_", 10, None)],
            [["VEN-001", "SQFT", 13.5, D(2027, 4, 1)], ["VEN-002", "KG", 27, D(2027, 4, 1)]], "Coating rate history.", "Dispatch / Accounts")
    b.sheet("X_Product", "imp", "X_PRODUCT / PART - from Production (values)", "Power Query extract.",
            [("Part ID", "lk", 10, None), ("Product ID", "lk", 9, None), ("Wt/pc kg", "lk", 8, None), ("Sq ft/pc", "lk", 7, None), ("Coating Route", "lk", 12, None)],
            [["A12-BP-01", "PRD-A12", 6.215, 4.15, "POWDER COAT"], ["A12-SP-04", "PRD-A12", 2.01, 1.6, "POWDER COAT"], ["B07-CB-02", "PRD-B07", 1.15, 0.9, "POWDER COAT"], ["C03-FR-01", "PRD-C03", 12.4, 6.2, "NO COATING"]],
            "Parts with weight and coating area.", "Read only")
    b.sheet("X_ProdOutput", "imp", "X_PROD OUTPUT - QC-OK parts available to send (values)", "From Production: completed at the last route stage, net of rejections, by batch.",
            [("Batch ID", "lk", 15, None), ("Part ID", "lk", 10, None), ("QC-OK Qty", "lk", 8, None), ("Already on JWC", "lk", 9, f"=SUMIFS({rng(JL,'G')},{rng(JL,'E')},A{{r}},{rng(JL,'F')},B{{r}})"),
             ("Available to Send", "ca", 9, "=C{r}-D{r}"), ("Run ID", "lk", 15, None)],
            [["ISS-2728-000001", "A12-BP-01", 49, None, None, "RUN-270415-0800"], ["ISS-2728-000002", "B07-CB-02", 280, None, None, "RUN-270415-0800"]],
            "What production can send on job work.", "Read only")
    b.sheet("Requirement Register", "reg", "CUSTOMER REQUIREMENT REGISTER - versioned", "A revision is a new row quoting the row it revises; the old one becomes SUPERSEDED.",
            [("Req ID", "ca", 14, None), ("Version", "ca", 6, None), ("Revises", "ca", 14, None), ("Entry Date", "ca", 9, None), ("Customer ID", "ca", 9, None), ("Product ID", "ca", 9, None),
             ("Required Qty", "ca", 8, None), ("Required Date", "ca", 9, None), ("Priority", "ca", 6, None), ("Status", "ca", 10, None),
             ("Dispatched Qty", "ca", 9, f"=SUMIFS({rng(DL,'J')},{rng(DL,'D')},A{{r}})"), ("Pending Qty", "to", 8, "=IF(J{r}=\"SUPERSEDED\",0,MAX(0,G{r}-K{r}))")],
            [["REQ-2728-00001", 1, "", D(2027, 4, 2), "CUS-01", "PRD-A12", 200, D(2027, 4, 30), 1, "SUPERSEDED"], ["REQ-2728-00002", 2, "REQ-2728-00001", D(2027, 4, 9), "CUS-01", "PRD-A12", 240, D(2027, 4, 30), 1, "OPEN"],
             ["REQ-2728-00003", 1, "", D(2027, 4, 5), "CUS-02", "PRD-C03", 30, D(2027, 4, 20), 2, "OPEN"]],
            "Customer schedules with versions; pending from DC lines quoting Req ID.", "Dispatch / Sales")
    b.sheet(JH, "reg", "JWC HEADER REGISTER - values only", "JWC numbers are sequential per FY, no gaps; cancellation is a status line.",
            [("JWC No", "ca", 14, None), ("JWC Date", "ca", 9, None), ("Vendor ID", "ca", 8, None), ("Vendor Name", "ca", 16, None), ("Vehicle No", "ca", 12, None), ("e-Way Bill", "ca", 12, None), ("Purpose", "ca", 16, None), ("Status", "ca", 8, None), ("Posted At", "ca", 12, None)],
            [["JWC/27-28/0001", D(2027, 4, 13), "VEN-001", "Coater 1 (masked)", "GJ06-XX-0001", "[number]", "FOR POWDER COATING", "OPEN", dt.datetime(2027, 4, 13, 16, 0)],
             ["JWC/27-28/0002", D(2027, 4, 14), "VEN-002", "Coater 2 (masked)", "GJ06-XX-0002", "[number]", "FOR POWDER COATING", "OPEN", dt.datetime(2027, 4, 14, 15, 0)]],
            "One row per job-work challan.", "Read only (posted)")
    b.sheet(JL, "reg", "JWC LINE REGISTER - values only", "Rate, basis, sq ft and lot frozen at posting. Returns and pending are in Open JWC Balance.",
            [("JWC Line ID", "ca", 17, None), ("JWC No", "ca", 14, None), ("JWC Date", "ca", 9, None), ("Vendor ID", "ca", 8, None), ("Batch ID", "ca", 15, None), ("Part ID", "ca", 10, None), ("Qty NOS", "ca", 7, None),
             ("Wt kg", "ca", 8, None), ("Sq ft", "ca", 7, None), ("Basis", "ca", 6, None), ("PC Rate", "ca", 7, None), ("PC Cost Rs", "ca", 9, None), ("Posted At", "ca", 12, None)],
            [["JWC/27-28/0001-01", "JWC/27-28/0001", D(2027, 4, 13), "VEN-001", "ISS-2728-000001", "A12-BP-01", 40, 248.6, 166, "SQFT", 13.5, 2241, dt.datetime(2027, 4, 13, 16, 0)],
             ["JWC/27-28/0002-01", "JWC/27-28/0002", D(2027, 4, 14), "VEN-002", "ISS-2728-000002", "B07-CB-02", 100, 115, 90, "KG", 27, 3105, dt.datetime(2027, 4, 14, 15, 0)]],
            "Parts sent to coaters.", "Read only (posted)")
    b.sheet(CM, "reg", "COATER MOVEMENT REGISTER - opening, returns, defects", "BACK-OK puts goods into FG; DEFECT-BACK needs a decision; OPENING only at cut-over.",
            [("Coater Move ID", "ca", 14, None), ("Move Date", "ca", 9, None), ("Move Type", "ca", 12, None), ("Vendor ID", "ca", 8, None), ("JWC Line ID", "ca", 17, None), ("Part ID", "ca", 10, None),
             ("Qty NOS", "ca", 7, None), ("Wt kg", "ca", 8, None), ("Whose Fault", "ca", 9, None), ("Decision", "ca", 10, None), ("Debit Rs", "ca", 7, None), ("Coater Challan", "ca", 10, None), ("Posted At", "ca", 12, None)],
            [["CRT-2728-00001", D(2027, 4, 1), "OPENING", "VEN-001", "", "A12-BP-01", 20, 124.3, "", "", "", "Statement 31-Mar", dt.datetime(2027, 4, 1, 9, 0)],
             ["CRT-2728-00002", D(2027, 4, 15), "BACK-OK", "VEN-001", "JWC/27-28/0001-01", "A12-BP-01", 30, 186.45, "", "", "", "UMB/C-881", dt.datetime(2027, 4, 15, 11, 0)],
             ["CRT-2728-00003", D(2027, 4, 15), "DEFECT-BACK", "VEN-001", "JWC/27-28/0001-01", "A12-BP-01", 2, 12.43, "COATER", "RE-SEND", "", "UMB/C-881", dt.datetime(2027, 4, 15, 11, 0)]],
            "Coater opening (cut-over), returns OK, defects with decision and debit.", "Read only (posted)")
    b.sheet(DH, "reg", "DC HEADER REGISTER - values only", "DC numbers sequential per FY, no gaps.",
            [("Challan No", "ca", 13, None), ("DC Date", "ca", 9, None), ("Customer ID", "ca", 9, None), ("Dispatch From", "ca", 10, None), ("Vehicle No", "ca", 12, None), ("LR No", "ca", 8, None),
             ("Invoice No", "ca", 9, None), ("e-Way Bill", "ca", 10, None), ("Status", "ca", 10, None), ("Posted At", "ca", 12, None)],
            [["DC/27-28/0001", D(2027, 4, 14), "CUS-01", "FACTORY", "GJ06-XX-0101", "[ref]", "[ref]", "[number]", "DISPATCHED", dt.datetime(2027, 4, 14, 17, 0)],
             ["DC/27-28/0002", D(2027, 4, 15), "CUS-01", "VEN-002", "GJ06-XX-0102", "[ref]", "[ref]", "[number]", "DISPATCHED", dt.datetime(2027, 4, 15, 12, 0)]],
            "One row per delivery challan.", "Read only (posted)")
    b.sheet(DL, "reg", "DC LINE REGISTER - values only", "UOM, weight per unit and rate frozen at posting.",
            [("DC Line ID", "ca", 15, None), ("Challan No", "ca", 13, None), ("DC Date", "ca", 9, None), ("Req ID", "ca", 14, None), ("Dispatch From", "ca", 9, None), ("JWC Line ID", "ca", 17, None),
             ("Part ID", "ca", 10, None), ("Batch ID", "ca", 15, None), ("Type", "ca", 6, None), ("Qty", "ca", 5, None), ("UOM", "ca", 5, None), ("Wt kg", "ca", 8, None), ("Challan Qty", "ca", 8, None),
             ("Rate Rs/kg", "ca", 8, None), ("Value Rs", "ca", 10, None)],
            [["DC/27-28/0001-01", "DC/27-28/0001", D(2027, 4, 14), "REQ-2728-00002", "FACTORY", "", "A12-BP-01", "ISS-2728-000001", "SALE", 25, "QTY", 155.375, 25, 173.24, 26917.17],
             ["DC/27-28/0002-01", "DC/27-28/0002", D(2027, 4, 15), "", "VEN-002", "JWC/27-28/0002-01", "B07-CB-02", "ISS-2728-000002", "SALE", 60, "QTY", 69, 60, 165.1, 11391.9]],
            "Delivery lines.", "Read only (posted)")
    OB = "Open JWC Balance"
    b.sheet(OB, "bal", "OPEN JWC / AT-COATER BALANCE (calculated)", "Pending = sent - back OK - defect back - direct dispatch. Ageing from JWC date.",
            [("JWC Line ID", "lk", 17, None), ("Vendor ID", "lk", 8, f"=INDEX({rng(JL,'D')},MATCH(A{{r}},{rng(JL,'A')},0))"), ("JWC Date", "lk", 9, f"=INDEX({rng(JL,'C')},MATCH(A{{r}},{rng(JL,'A')},0))"),
             ("Part ID", "lk", 10, f"=INDEX({rng(JL,'F')},MATCH(A{{r}},{rng(JL,'A')},0))"), ("Sent", "lk", 6, f"=INDEX({rng(JL,'G')},MATCH(A{{r}},{rng(JL,'A')},0))"),
             ("Back OK", "ca", 7, f"=SUMIFS({rng(CM,'G')},{rng(CM,'E')},A{{r}},{rng(CM,'C')},\"BACK-OK\")"), ("Defect Back", "ca", 7, f"=SUMIFS({rng(CM,'G')},{rng(CM,'E')},A{{r}},{rng(CM,'C')},\"DEFECT-BACK\")"),
             ("Direct DC", "ca", 7, f"=SUMIFS({rng(DL,'J')},{rng(DL,'F')},A{{r}})"), ("Pending Qty", "to", 8, "=E{r}-F{r}-G{r}-H{r}"),
             ("Pending kg", "ca", 8, f"=ROUND(I{{r}}*INDEX({rng(JL,'H')},MATCH(A{{r}},{rng(JL,'A')},0))/E{{r}},3)"), ("Days", "ca", 5, f"={AS}-C{{r}}"),
             ("Status", "cf", 20, "=IF(I{r}<0,\"BREACH - more back than sent\",IF(I{r}=0,\"CLOSED\",IF(K{r}>'Settings'!$B$11,\"WARN: over age\",IF(F{r}+G{r}+H{r}>0,\"PART RETURNED\",\"PENDING\"))))")],
            [["JWC/27-28/0001-01"], ["JWC/27-28/0002-01"]], "Material pending at each coater by JWC line with ageing.", "Dispatch (read only)", flagcol=12,
            status=[("Coater opening (cut-over) qty:", "h2"), (f"=SUMIFS({rng(CM,'G')},{rng(CM,'C')},\"OPENING\")", "to")])
    FB = "FG Balance"
    b.sheet(FB, "bal", "FINISHED-GOODS BALANCE - transaction-calculated (no counts)", "FG = opening + coater BACK-OK + uncoated final-stage completion - dispatch from factory +/- authorised adjustment.",
            [("Part ID", "lk", 10, None), ("FG Opening", "lk", 8, None), ("In: Back OK", "ca", 8, f"=SUMIFS({rng(CM,'G')},{rng(CM,'F')},A{{r}},{rng(CM,'C')},\"BACK-OK\")"),
             ("In: Uncoated", "lk", 8, None), ("Out: DC from Factory", "ca", 10, f"=SUMIFS({rng(DL,'J')},{rng(DL,'G')},A{{r}},{rng(DL,'E')},\"FACTORY\")"),
             ("FG Available", "to", 8, "=B{r}+C{r}+D{r}-E{r}"), ("Status", "cf", 16, "=IF(F{r}<0,\"BREACH - negative FG\",\"OK\")")],
            [["A12-BP-01", 0, None, 0], ["B07-CB-02", 0, None, 0], ["C03-FR-01", 0, None, 0]], "FG available by part for the dispatch hard stop.", "Dispatch (read only)", flagcol=7)
    FBR = lambda col: rng(FB, col, 5, 100)
    OBR = lambda col: rng(OB, col, 5, 100)
    b.sheet("JWC Entry", "pad", "JOB WORK CHALLAN PAD - header + lines", "Pick batch and part; qty cannot exceed what production has available to send. Rate comes from Coater Rate on the JWC date.",
            [("Line", "lk", 5, None), ("JWC Line ID", "ca", 17, "=IF(C{r}=\"\",\"\",$B$3&\"-\"&TEXT(A{r},\"00\"))"), ("Batch ID", "in_", 15, None), ("Part ID", "in_", 10, None),
             ("Avail to Send", "lk", 8, f"=IF(C{{r}}=\"\",\"\",SUMIFS('X_ProdOutput'!$E$5:$E$50,'X_ProdOutput'!$A$5:$A$50,C{{r}},'X_ProdOutput'!$B$5:$B$50,D{{r}})-SUMIFS(F$6:F{{p}},C$6:C{{p}},C{{r}},D$6:D{{p}},D{{r}}))"),
             ("Qty NOS", "in_", 7, None), ("Wt kg", "ca", 8, "=IF(F{r}=\"\",\"\",ROUND(F{r}*INDEX('X_Product'!$C$5:$C$50,MATCH(D{r},'X_Product'!$A$5:$A$50,0)),3))"),
             ("Sq ft", "ca", 7, "=IF(F{r}=\"\",\"\",ROUND(F{r}*INDEX('X_Product'!$D$5:$D$50,MATCH(D{r},'X_Product'!$A$5:$A$50,0)),1))"),
             ("Basis", "ca", 6, "=IF(C{r}=\"\",\"\",INDEX('Coater Rate'!$B$5:$B$50,MATCH($D$3,'Coater Rate'!$A$5:$A$50,0)))"),
             ("PC Rate", "ca", 7, "=IF(C{r}=\"\",\"\",INDEX('Coater Rate'!$C$5:$C$50,MATCH($D$3,'Coater Rate'!$A$5:$A$50,0)))"),
             ("PC Cost Rs", "ca", 9, "=IF(C{r}=\"\",\"\",ROUND(IF(I{r}=\"KG\",G{r},H{r})*J{r},2))"),
             ("Ready / issues", "cf", 34, "=IF(C{r}=\"\",\"\",IF(F{r}>E{r},\"MORE THAN AVAILABLE TO SEND: max \"&E{r},\"READY\"))")],
            [[1, None, "ISS-2728-000001", "A12-BP-01", None, 8], [2, None, "ISS-2728-000002", "B07-CB-02", None, 150], [3, None, "ISS-2728-000002", "B07-CB-02", None, 40]],
            "Send parts to a coater: header + lines with availability and rate snapshot.", "Dispatch", flagcol=12, table=False, hdr=6,
            pre=[[("JWC No", "h2"), (f"=\"JWC/27-28/\"&TEXT(COUNTA({rng(JH,'A')})+1,\"0000\")", "ca"), ("Vendor", "h2"), ("VEN-001", "in_"), ("Date", "h2"), (D(2027, 4, 15), "in_"), ("Vehicle", "h2"), ("GJ06-XX-0003", "in_"), ("Purpose", "h2"), ("FOR POWDER COATING", "in_")],
                 [("e-Way Bill", "h2"), ("[number]", "in_"), ("Shade", "h2"), ("RAL 7035", "in_")]])
    b.wb["JWC Entry"]["F3"].number_format = "dd-mmm-yy"
    b.sheet("Coater Move Entry", "pad", "COATER RETURN / DEFECT PAD", "Quote the JWC line. Returns above pending are blocked; defects need a decision.",
            [("Coater Move ID", "ca", 14, f"=IF(C{{r}}=\"\",\"\",\"CRT-2728-\"&TEXT(COUNTA({rng(CM,'A')})+ROW()-4,\"00000\"))"), ("Move Date", "in_", 9, None), ("Move Type", "in_", 12, None),
             ("JWC Line ID", "in_", 17, None), ("Pending", "lk", 7, f"=IF(D{{r}}=\"\",\"\",IFERROR(INDEX({OBR('I')},MATCH(D{{r}},{OBR('A')},0)),0)-SUMIFS(F$4:F{{p}},D$4:D{{p}},D{{r}}))"),
             ("Qty NOS", "in_", 7, None), ("Whose Fault", "in_", 9, None), ("Decision", "in_", 10, None), ("Coater Challan", "in_", 10, None),
             ("Ready / issues", "cf", 32, "=IF(C{r}=\"\",\"\",IF(F{r}>E{r},\"MORE THAN PENDING: max \"&E{r},IF(AND(C{r}=\"DEFECT-BACK\",H{r}=\"\"),\"DECISION REQUIRED\",\"READY\")))")],
            [[None, D(2027, 4, 15), "BACK-OK", "JWC/27-28/0001-01", None, 6, "", "", "UMB/C-890"], [None, D(2027, 4, 15), "BACK-OK", "JWC/27-28/0002-01", None, 50, "", "", "C2/551"],
             [None, D(2027, 4, 15), "DEFECT-BACK", "JWC/27-28/0001-01", None, 1, "COATER", "", "UMB/C-890"]],
            "Record coater returns OK and defects with decision.", "Dispatch", flagcol=10, table=False)
    b.sheet("DC Entry", "pad", "DELIVERY CHALLAN PAD - header + lines", "From FACTORY a line cannot exceed FG available; from a coater it must quote the JWC line and cannot exceed pending. UOM and rate from Customer Price & UOM.",
            [("Line", "lk", 5, None), ("DC Line ID", "ca", 16, "=IF(D{r}=\"\",\"\",$B$3&\"-\"&TEXT(A{r},\"00\"))"), ("Req ID", "in_", 14, None), ("Part ID", "in_", 10, None), ("Qty", "in_", 6, None),
             ("Available", "lk", 8, f"=IF(D{{r}}=\"\",\"\",IFERROR(INDEX({FBR('F')},MATCH(D{{r}},{FBR('A')},0)),0)-SUMIFS(E$6:E{{p}},D$6:D{{p}},D{{r}}))"),
             ("UOM", "ca", 5, "ARR:=IF(D{r}=\"\",\"\",IFERROR(INDEX('Customer Price & UOM'!$D$5:$D$50,MATCH(1,('Customer Price & UOM'!$A$5:$A$50=$D$3)*('Customer Price & UOM'!$C$5:$C$50=D{r}),0)),\"\"))"),
             ("Wt / Unit", "ca", 7, "ARR:=IF(D{r}=\"\",\"\",IFERROR(INDEX('Customer Price & UOM'!$E$5:$E$50,MATCH(1,('Customer Price & UOM'!$A$5:$A$50=$D$3)*('Customer Price & UOM'!$C$5:$C$50=D{r}),0)),0))"),
             ("Wt kg", "ca", 8, "=IF(E{r}=\"\",\"\",ROUND(E{r}*H{r},3))"), ("Challan Qty", "ca", 8, "=IF(E{r}=\"\",\"\",IF(G{r}=\"KG\",I{r},IF(G{r}=\"TON\",ROUND(I{r}/1000,3),E{r})))"),
             ("Rate Rs/kg", "ca", 8, "ARR:=IF(D{r}=\"\",\"\",IFERROR(INDEX('Customer Price & UOM'!$F$5:$F$50,MATCH(1,('Customer Price & UOM'!$A$5:$A$50=$D$3)*('Customer Price & UOM'!$C$5:$C$50=D{r}),0)),0))"),
             ("Value Rs", "ca", 10, "=IF(E{r}=\"\",\"\",ROUND(I{r}*K{r},2))"),
             ("Ready / issues", "cf", 34, "=IF(D{r}=\"\",\"\",IF(G{r}=\"\",\"NO UOM / PRICE FOR CUSTOMER + PART\",IF(E{r}>F{r},\"ABOVE FG AVAILABLE: max \"&F{r},IF(C{r}=\"\",\"WARN: no Req ID\",\"READY\"))))")],
            [[1, None, "REQ-2728-00002", "A12-BP-01", 4], [2, None, "REQ-2728-00002", "A12-BP-01", 10], [3, None, "", "C03-FR-01", 5]],
            "Dispatch to customer with FG / JWC hard stops and frozen UOM and rate.", "Dispatch", flagcol=13, table=False, hdr=6,
            pre=[[("Challan No", "h2"), (f"=\"DC/27-28/\"&TEXT(COUNTA({rng(DH,'A')})+1,\"0000\")", "ca"), ("Customer", "h2"), ("CUS-01", "in_"), ("Date", "h2"), (D(2027, 4, 15), "in_"), ("From", "h2"), ("FACTORY", "in_"), ("Vehicle", "h2"), ("GJ06-XX-0103", "in_")],
                 [("LR No", "h2"), ("[ref]", "in_"), ("Invoice", "h2"), ("[ref]", "in_"), ("e-Way Bill", "h2"), ("[number]", "in_"), ("Driver", "h2"), ("[name]", "in_")]])
    b.wb["DC Entry"]["F3"].number_format = "dd-mmm-yy"
    b.sheet("Control Totals", "ctl", "CONTROL TOTALS - read by MIS", "Rows and totals per register.",
            [("Register", "lk", 24, None), ("Rows", "ca", 7, None), ("Total Qty", "ca", 10, None), ("Total kg / Rs", "ca", 12, None)],
            [[JL, f"=COUNTA({rng(JL,'A')})", f"=SUM({rng(JL,'G')})", f"=SUM({rng(JL,'H')})"], [CM, f"=COUNTA({rng(CM,'A')})", f"=SUM({rng(CM,'G')})", f"=SUM({rng(CM,'H')})"],
             [DL, f"=COUNTA({rng(DL,'A')})", f"=SUM({rng(DL,'J')})", f"=SUM({rng(DL,'O')})"]], "Reconciliation totals.", "MIS / Finance")
    lists(b, [("Coater Move Type", ["OPENING", "BACK-OK", "DEFECT-BACK", "DEFECT-AT-COATER-SCRAP"]), ("Decision", ["REWORK AT COATER", "RE-SEND", "SCRAP", "DEBIT COATER"]),
              ("DC Type", ["SALE", "REPLACEMENT", "SAMPLE"]), ("UOM", ["QTY", "NOS", "SET", "KG", "TON"])])
    order = ["README", "JWC Entry", "Coater Move Entry", "DC Entry", "Requirement Register", JH, JL, CM, DH, DL, OB, FB, "M_Customer", "Customer Price & UOM", "M_Vendor", "Coater Rate",
             "X_Product", "X_ProdOutput", "Control Totals", "Settings", "Lists"]
    return b, order


# =================================================================== 4 MIS
def mis():
    b = Book("MIS", "4 ISKO MIS FY2027-28 (prototype).xlsx")
    readme(b, "4 ISKO MIS FY2027-28 - read-only reporting: KPIs, exceptions, trace", [
        ("Routine", ""), ("Before review (daily)", "Data > Refresh All. Read Refresh Log - every source must be OK. Then Control Tower and Exceptions."),
        ("Typing", "Only the period / filter boxes and the Lot / DC pick cells. No transaction is ever entered here."),
        ("If a source is STALE or MISSING", "The banner on the Control Tower turns coral; ask the book owner to save, then refresh again."),
        ("Detail extracts", "X_LotBal, X_RMMove, X_Prod, X_BatchBal, X_WIP, X_JWCLine, X_JWCBal, X_FGBal, X_DCLine load to the Power Query data model (not sheets); only the small summaries X_KPI_Month and X_Balances sit on sheets.")])
    settings(b, [["DataFolder", "D:\\ISKO ERP\\FY2027-28\\", "Typed once"], ["StoresFile", "1 ISKO Stores FY2027-28.xlsx", ""], ["ProductionFile", "2 ISKO Production FY2027-28.xlsx", ""],
                 ["DispatchFile", "3 ISKO Dispatch & Challan FY2027-28.xlsx", ""], ["Period", "Apr-27", "Month shown"], ["MaxAgeHours", 24, "Older source = STALE"], ["ArchiveFolder", "D:\\ISKO ERP\\Archive\\", "Earlier FY books"]])
    RL = "Refresh Log"
    b.sheet(RL, "ctl", "REFRESH LOG - is every number fresh and complete?", "Written by Power Query on Refresh All. Rows must equal distinct IDs; control totals must equal the source book's Control Totals.",
            [("Run ID", "lk", 15, None), ("Source Book", "lk", 12, None), ("Extract", "lk", 13, None), ("Refreshed At", "lk", 13, None), ("Source Saved", "lk", 13, None), ("Rows", "lk", 6, None),
             ("Distinct IDs", "lk", 8, None), ("Control kg", "lk", 10, None), ("Source Control kg", "lk", 10, None),
             ("Status", "cf", 24, "=IF(F{r}<>G{r},\"DUPLICATES\",IF(ROUND(H{r}-I{r},3)<>0,\"MISMATCH\",IF(D{r}-E{r}>'Settings'!$B$10/24,\"STALE - source older than \"&'Settings'!$B$10&\" h\",\"OK\")))")],
            [["RUN-270415-0915", "Stores", "X_LotBal", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 14, 18, 5), 5, 5, 3941.4, 3941.4],
             ["RUN-270415-0915", "Stores", "X_RMMove", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 14, 18, 5), 4, 4, -1087.5, -1087.5],
             ["RUN-270415-0915", "Production", "X_Prod", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 14, 18, 20), 6, 6, 1909.55, 1909.55],
             ["RUN-270415-0915", "Production", "X_BatchBal", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 14, 18, 20), 3, 3, 109.2, 109.2],
             ["RUN-270415-0915", "Dispatch", "X_DCLine", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 13, 17, 0), 2, 2, 224.375, 224.375],
             ["RUN-270415-0915", "Dispatch", "X_JWCLine", dt.datetime(2027, 4, 15, 9, 15), dt.datetime(2027, 4, 13, 17, 0), 2, 2, 363.6, 363.6]],
            "Proof of freshness and completeness of every extract.", "MIS owner", flagcol=10)
    # summary extract (monthly, values)
    b.sheet("X_KPI_Month", "imp", "X_KPI_MONTH - monthly summary built by Power Query (values)", "Small table (one row per month) - every tile and chart reads this, never a register. Earlier months come from archived books (sample values).",
            [("Month", "lk", 8, None), ("RM Issued kg", "lk", 10, None), ("Cut Output kg", "lk", 10, None), ("Net Batch kg (closed)", "lk", 11, None), ("Cut Output kg (closed)", "lk", 11, None),
             ("Scrap kg", "lk", 8, None), ("Scrap Rs", "lk", 9, None), ("Rejected kg", "lk", 8, None), ("Dispatched kg", "lk", 10, None), ("Dispatch Rs", "lk", 11, None), ("PC Cost Rs", "lk", 9, None), ("Sent to Coater kg", "lk", 10, None)],
            [["Jan-27", 121500, 109200, 118900, 108800, 9100, 678000, 210, 96100, 15950000, 268000, 9900], ["Feb-27", 128300, 115600, 125400, 114700, 9800, 730100, 185, 103800, 17130000, 281000, 10400],
             ["Mar-27", 133800, 121000, 131200, 120100, 10200, 760000, 240, 109200, 18050000, 296500, 11000], ["Apr-27", 1105.9, 903.9, 0, 0, 74.4, 5543, 24.86, 224.375, 38309.07, 5346, 363.6]],
            "Monthly KPI summary (Power Query output).", "Read only")
    b.sheet("X_Balances", "imp", "X_BALANCES - current positions from the three books (values)", "Lot, WIP, at-coater and FG balances as published by their owners.",
            [("Position", "lk", 22, None), ("Owner Book", "lk", 11, None), ("kg", "lk", 10, None), ("Rs", "lk", 11, None), ("Items", "lk", 6, None), ("Oldest Days", "lk", 8, None)],
            [["RM stock (open lots)", "Stores", 3941.4, 287102, 5, 14], ["RM stock older than 90 days", "Stores", 0, 0, 0, 0], ["WIP (calculated)", "Production", 202.29, 15070, 4, 3],
             ["At coaters (pending)", "Dispatch", 291.05, 0, 2, 2], ["FG available (calculated)", "Dispatch", 31.08, 0, 1, 0], ["Consumables", "Stores", 0, 21842, 4, 0]],
            "Current stock positions.", "Read only")
    b.sheet("Exceptions", "rpt", "EXCEPTIONS - one row per open problem, with owner and age", "Built on each refresh. BLOCKER rows mean a control was bypassed - same-day review.",
            [("Exception ID", "lk", 26, None), ("Code", "lk", 7, None), ("Exception", "lk", 36, None), ("Severity", "cf", 9, None), ("Owner", "lk", 11, None), ("Source Row ID", "lk", 17, None),
             ("First Seen", "lk", 9, None), ("Days Open", "lk", 6, None), ("Action Note", "in_", 26, None)],
            [["STK-03|CR-2728-0001", "STK-03", "Lot in QC HOLD more than 3 days", "HIGH", "QC", "CR-2728-0001", D(2027, 4, 12), 3, "Supplier visit 16-Apr"],
             ["PUR-01|PO/27-28/001-01", "PUR-01", "PO line past due with pending qty", "MEDIUM", "Purchase", "PO/27-28/001-01", D(2027, 4, 13), 2, ""],
             ["PRD-03|ISS-2728-000004", "PRD-03", "Batch variance above tolerance (in progress)", "MEDIUM", "Production", "ISS-2728-000004", D(2027, 4, 14), 1, ""],
             ["SYS-01|X_DCLine", "SYS-01", "Source book stale (Dispatch saved 13-Apr)", "BLOCKER", "MIS owner", "X_DCLine", D(2027, 4, 15), 0, "Asked Dispatch to save"]],
            "Open exceptions by owner and age.", "Each owner; MD reviews", flagcol=4)
    ex = b.wb["Exceptions"]
    ex.conditional_formatting = type(ex.conditional_formatting)()
    ex.conditional_formatting.add("D5:D20", FormulaRule(formula=['D5="BLOCKER"'], fill=FILL["ex"], stopIfTrue=True))
    ex.conditional_formatting.add("D5:D20", FormulaRule(formula=['D5="HIGH"'], fill=FILL["pe"], stopIfTrue=True))
    KM = "X_KPI_Month"
    ct = b.sheet("Control Tower", "rpt", "CONTROL TOWER - director view", "Period and filters at the top; every tile reads the small monthly summary or the balances - nothing scans a register.",
                 [("KPI", "h2", 30, None), ("Value", "to", 14, None), ("Unit", "lk", 9, None), ("Previous Month", "lk", 13, None), ("Owner", "lk", 13, None), ("Source", "lk", 16, None)],
                 [["Cutting output (month)", f"=INDEX('{KM}'!$C$5:$C$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/1000", "MT", f"=INDEX('{KM}'!$C$5:$C$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))/1000", "Production", "X_KPI_Month"],
                  ["Cutting yield (closed batches)", f"=INDEX('{KM}'!$E$5:$E$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/INDEX('{KM}'!$D$5:$D$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))", "%", f"=INDEX('{KM}'!$E$5:$E$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))/INDEX('{KM}'!$D$5:$D$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))", "Production", "X_BatchBal"],
                  ["Scrap loss", f"=INDEX('{KM}'!$G$5:$G$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/100000", "Rs lakh", f"=INDEX('{KM}'!$G$5:$G$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))/100000", "Production", "Q_MaterialFlow"],
                  ["Dispatched", f"=INDEX('{KM}'!$I$5:$I$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/1000", "MT", f"=INDEX('{KM}'!$I$5:$I$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))/1000", "MD", "X_DCLine"],
                  ["Realisation", f"=INDEX('{KM}'!$J$5:$J$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/INDEX('{KM}'!$I$5:$I$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))", "Rs/kg", f"=INDEX('{KM}'!$J$5:$J$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))/INDEX('{KM}'!$I$5:$I$20,MATCH(\"Feb-27\",'{KM}'!$A$5:$A$20,0))", "MD", "X_DCLine"],
                  ["Powder-coating cost", f"=INDEX('{KM}'!$K$5:$K$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))/INDEX('{KM}'!$L$5:$L$20,MATCH(\"Mar-27\",'{KM}'!$A$5:$A$20,0))", "Rs/kg", "", "Dispatch / Finance", "X_JWCLine"],
                  ["RM stock (open lots)", "=INDEX('X_Balances'!$C$5:$C$20,1)/1000", "MT", "", "Stores", "X_LotBal"], ["WIP (calculated)", "=INDEX('X_Balances'!$C$5:$C$20,3)/1000", "MT", "", "Production", "X_WIP"],
                  ["At coaters", "=INDEX('X_Balances'!$C$5:$C$20,4)/1000", "MT", "", "Dispatch", "X_JWCBal"], ["FG available (calculated)", "=INDEX('X_Balances'!$C$5:$C$20,5)/1000", "MT", "", "Dispatch", "X_FGBal"],
                  ["Open BLOCKER exceptions", "=COUNTIF('Exceptions'!$D$5:$D$200,\"BLOCKER\")", "count", "", "All", "Exceptions"],
                  ["Data freshness", "=IF(COUNTIF('Refresh Log'!$J$5:$J$200,\"OK\")=COUNTA('Refresh Log'!$A$5:$A$200),\"ALL SOURCES OK\",\"CHECK REFRESH LOG\")", "", "", "MIS owner", "Refresh Log"]],
                 "Director KPIs, data-freshness banner and dispatch trend.", "MD / heads", table=False, hdr=6,
                 pre=[[("Period", "h2"), ("Mar-27", "in_"), ("Customer", "h2"), ("All", "in_"), ("Material", "h2"), ("All", "in_")],
                      [("DATA STATUS", "h2"), ("=IF(COUNTIF('Refresh Log'!$J$5:$J$200,\"OK\")=COUNTA('Refresh Log'!$A$5:$A$200),\"ALL SOURCES OK\",\"ONE OR MORE SOURCES NOT OK - see Refresh Log\")", "pe"), None, None, None, None]])
    for r in range(7, 19):
        c = ct.cell(r, 2)
        c.number_format = "0.0%" if r == 8 else "#,##0.00"
    ch = BarChart(); ch.type = "col"; ch.title = "Dispatched kg by month (from X_KPI_Month)"; ch.height = 6.5; ch.width = 13; ch.legend = None
    ch.add_data(Reference(b.wb[KM], min_col=9, min_row=4, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(b.wb[KM], min_col=1, min_row=5, max_row=7))
    ct.add_chart(ch, "H3")
    b.sheet("Stock & WIP", "rpt", "STOCK & WIP SUMMARY - all transaction-calculated", "No physical-count block. Every position comes from its owning book.",
            [("Position", "lk", 26, None), ("kg", "to", 10, "='X_Balances'!C{r}"), ("Rs", "to", 11, "='X_Balances'!D{r}"), ("Items / lines", "lk", 9, "='X_Balances'!E{r}"), ("Oldest (days)", "lk", 9, "='X_Balances'!F{r}"), ("Owner", "lk", 11, "='X_Balances'!B{r}")],
            [["='X_Balances'!A{r}"] for _ in range(6)], "RM, WIP, at-coater, FG and consumable positions.", "Directors / Finance", table=False)
    b.sheet("Production & Yield", "rpt", "PRODUCTION & YIELD - by month", "From X_KPI_Month (PivotTable with slicers in the real build).",
            [("Month", "lk", 8, "='X_KPI_Month'!A{r}"), ("RM Issued kg", "ca", 11, "='X_KPI_Month'!B{r}"), ("Cut Output kg", "ca", 11, "='X_KPI_Month'!C{r}"), ("Yield (closed)", "to", 9, "=IFERROR('X_KPI_Month'!E{r}/'X_KPI_Month'!D{r},\"-\")"),
             ("Scrap kg", "ca", 9, "='X_KPI_Month'!F{r}"), ("Scrap Rs", "ca", 10, "='X_KPI_Month'!G{r}"), ("Rejected kg", "ca", 9, "='X_KPI_Month'!H{r}"), ("Rejection %", "to", 9, "=IFERROR(G{r}/C{r},0)")],
            [[] for _ in range(4)], "Output, yield, scrap, rejection trend.", "Production head", table=False)
    ws = b.wb["Production & Yield"]
    for r in range(5, 9):
        ws.cell(r, 4).number_format = "0.0%"; ws.cell(r, 8).number_format = "0.00%"
    b.sheet("Dispatch & JWC", "rpt", "DISPATCH & JOB WORK - by month", "Dispatch kg/value, realisation, coating cost.",
            [("Month", "lk", 8, "='X_KPI_Month'!A{r}"), ("Dispatched kg", "ca", 11, "='X_KPI_Month'!I{r}"), ("Dispatch Rs", "ca", 12, "='X_KPI_Month'!J{r}"), ("Rs / kg", "to", 8, "=IFERROR(C{r}/B{r},0)"),
             ("Sent to Coater kg", "ca", 11, "='X_KPI_Month'!L{r}"), ("PC Cost Rs", "ca", 10, "='X_KPI_Month'!K{r}"), ("PC Rs / kg", "to", 8, "=IFERROR(F{r}/E{r},0)")],
            [[] for _ in range(4)], "Dispatch and coating trend.", "Dispatch head", table=False)
    b.sheet("Lot Trace", "rpt", "LOT TRACE - pick a Lot ID in B3, then Refresh", "Power Query parameter query: PO -> GRN -> QC -> batches -> production -> JWC -> DC for one lot (no FILTER formulas).",
            [("Step", "h2", 12, None), ("ID", "lk", 17, None), ("Date", "lk", 9, None), ("Detail", "lk", 40, None), ("kg / Qty", "lk", 9, None), ("Book", "lk", 10, None)],
            [["PO line", "PO/27-28/001-01", D(2027, 4, 2), "Supplier A (masked), GPSP-1.6, ordered 3,000 kg", 3000, "Stores"], ["GRN / lot", "GP-2728-0001", D(2027, 4, 8), "Invoice INV/3381, 40 sheets", 1575.5, "Stores"],
             ["QC", "QC-2728-000002", D(2027, 4, 8), "ACCEPTED", 1575.5, "Stores"], ["Batch (issue)", "ISS-2728-000001", D(2027, 4, 12), "MC-01, 12 sheets", 471, "Stores"],
             ["Return", "OFC-2728-000003", D(2027, 4, 13), "Offcut back to lot", 18.4, "Stores"], ["Production", "PRD-2728-000001", D(2027, 4, 12), "CUTTING A12-BP-01 x 60", 372.9, "Production"],
             ["Job work", "JWC/27-28/0001-01", D(2027, 4, 13), "Coater 1, 40 nos, 166 sq ft", 40, "Dispatch"], ["Coater return", "CRT-2728-00002", D(2027, 4, 15), "BACK-OK 30 nos", 30, "Dispatch"],
             ["Delivery", "DC/27-28/0001-01", D(2027, 4, 14), "Customer 1, 25 nos, Req REQ-2728-00002", 25, "Dispatch"]],
            "End-to-end trace of one lot by IDs.", "Anyone", pre=[[("Lot ID", "h2"), ("GP-2728-0001", "in_"), ("Refreshed with run", "h2"), ("RUN-270415-0915", "lk")]], table=False, hdr=5)
    b.sheet("DC Trace", "rpt", "DC TRACE - pick a Challan No in B3, then Refresh", "Back from a delivery challan to JWC line, batch, lot, GRN and PO.",
            [("DC Line", "lk", 16, None), ("Part", "lk", 10, None), ("JWC Line", "lk", 17, None), ("Batch", "lk", 15, None), ("Lot", "lk", 12, None), ("GRN Date", "lk", 9, None), ("Supplier", "lk", 16, None), ("PO Line", "lk", 15, None)],
            [["DC/27-28/0002-01", "B07-CB-02", "JWC/27-28/0002-01", "ISS-2728-000002", "GP-2728-0002", D(2027, 4, 8), "Supplier A (masked)", "PO/27-28/001-02"]],
            "Backward trace from a delivery challan.", "Anyone", pre=[[("Challan No", "h2"), ("DC/27-28/0002", "in_")]], table=False, hdr=5)
    b.sheet("RM Requirement", "rpt", "RM REQUIREMENT - what to buy (replaces requisition)", "Open requirement x BOM set weight - RM stock - open PO pending (Power Query).",
            [("Material ID", "lk", 10, None), ("Required kg (open req x BOM)", "lk", 14, None), ("RM Stock kg", "lk", 10, None), ("On Order kg", "lk", 10, None), ("To Buy kg", "to", 10, "=MAX(0,B{r}-C{r}-D{r})"), ("Min Level kg", "lk", 9, None), ("Action", "cf", 18, "=IF(E{r}>0,\"WARN: raise PO\",\"OK\")")],
            [["GPSP-1.6", 10239, 2093.7, 1424.5, None, 1500], ["GPSP-1.2", 3640, 1281.3, 1640, None, 1000], ["HRCA-2.0", 2227, 425.3, 0, None, 600]],
            "Purchase need by material for PO decisions.", "Purchase", table=False, flagcol=7)
    b.sheet("MRM Monthly", "rpt", "MONTHLY REVIEW PACK (MRM)", "Same sections every month; read from the summaries. No mail, no print area.",
            [("Section", "h2", 30, None), ("This Month", "to", 13, None), ("Last Month", "lk", 13, None), ("Comment (typed in review)", "in_", 40, None)],
            [["1. Dispatch kg", "='X_KPI_Month'!I7", "='X_KPI_Month'!I6", ""], ["2. Cutting output kg", "='X_KPI_Month'!C7", "='X_KPI_Month'!C6", ""], ["3. Yield (closed batches)", "='Production & Yield'!D7", "='Production & Yield'!D6", ""],
             ["4. Scrap Rs", "='X_KPI_Month'!G7", "='X_KPI_Month'!G6", ""], ["5. PC cost Rs/kg", "='Dispatch & JWC'!G7", "='Dispatch & JWC'!G6", ""], ["6. Open BLOCKER exceptions", "='Control Tower'!B17", "", ""]],
            "Monthly management review pack.", "MD", table=False)
    order = ["README", "Control Tower", "Exceptions", RL, "Stock & WIP", "Production & Yield", "Dispatch & JWC", "RM Requirement", "Lot Trace", "DC Trace", "MRM Monthly", "X_KPI_Month", "X_Balances", "Settings"]
    return b, order


def build():
    books = [stores(), production(), dispatch(), mis()]
    meta = {}
    for b, order in books:
        names = [ws.title for ws in b.wb.worksheets]
        missing = [n for n in names if n not in order]
        order = [o for o in order if o in names] + missing
        b.wb._sheets = [b.wb[n] for n in order]
        p = b.save()
        info = {i["sheet"]: i for i in INFO[b.key]}
        meta[b.key] = dict(file=os.path.basename(p), sheets=[info[n] for n in order] + [i for i in INFO[b.key] if i["sheet"] not in order and i["sheet"] not in names])
        print(p, len(order), "sheets")
    json.dump(meta, open(os.path.join(OUT, "sheets.json"), "w"), indent=1, default=str)


if __name__ == "__main__":
    build()
