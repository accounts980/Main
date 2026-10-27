"""Generate HTML for the six illustrative Excel-layout mock-ups (sample / masked
data only). render.js screenshots them to assets/*.png.

Every image carries the label 'Illustrative Excel Layout - Mock-up' and
'sample data'. Times are shown as placeholders, never as live timestamps."""
import os, html

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "tools", "_build", "mock")
os.makedirs(OUT, exist_ok=True)

CSS = """
*{box-sizing:border-box} body{margin:0;background:#fff;font-family:Inter,'DejaVu Sans',Arial,sans-serif;color:#243247}
.win{display:inline-block;border:1px solid #C9D3DF;border-radius:8px;overflow:hidden;background:#fff;margin:6px}
.bar{background:#EEF2F7;border-bottom:1px solid #D9E0E8;padding:6px 10px;font-size:12px;display:flex;justify-content:space-between;gap:20px}
.bar b{font-weight:700}.tag{background:#EAE4F8;border:1px solid #CFC4EE;border-radius:10px;padding:1px 8px;font-size:11px;font-weight:600}
.fx{display:flex;border-bottom:1px solid #D9E0E8;font-size:11.5px}.fx .nm{width:90px;padding:4px 8px;border-right:1px solid #D9E0E8;color:#5A6B82}
.fx .f{padding:4px 8px;color:#5A6B82;font-family:'DejaVu Sans Mono',monospace;font-size:11px}
table.g{border-collapse:collapse;font-size:11.5px;table-layout:fixed}
table.g td{border:1px solid #D9E0E8;padding:3px 6px;height:22px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
table.g td.rh{background:#F1F4F8;color:#7A8AA0;text-align:center;width:30px;font-size:10.5px}
table.g tr.cl td{background:#F1F4F8;color:#7A8AA0;text-align:center;font-size:10.5px;height:18px}
.t{font-weight:700;font-size:13px;background:#fff}.s{color:#4A5A70;font-size:11px;background:#fff}
.h{background:#DCEAF7;font-weight:700}.h2{background:#EAE4F8;font-weight:700}
.in{background:#FFF4D6}.ca{background:#E2F0E8}.ex{background:#F8DCDD;color:#8A2F37;font-weight:600}
.pe{background:#FCECC9;color:#6B4E00;font-weight:600}.to{background:#D9F0EF;font-weight:700}.ok{background:#E2F0E8;color:#2E6B4A;font-weight:600}
.lk{background:#F7F9FC;color:#4A5A70}.r{text-align:right}.c{text-align:center}.w{background:#fff}
.tabs{display:flex;gap:2px;background:#EEF2F7;border-top:1px solid #D9E0E8;padding:4px 6px 0 36px;font-size:11px}
.tabs span{padding:4px 10px;background:#F7F9FC;border:1px solid #D9E0E8;border-bottom:none;border-radius:4px 4px 0 0;color:#5A6B82}
.tabs span.a{background:#fff;color:#243247;font-weight:700;border-bottom:2px solid #7FA7D6}
.legend{display:flex;gap:10px;padding:6px 10px;font-size:10.5px;color:#4A5A70;border-top:1px solid #D9E0E8;background:#fff;flex-wrap:wrap}
.legend i{display:inline-block;width:12px;height:12px;border:1px solid #C9D3DF;vertical-align:-2px;margin-right:4px}
.tile{display:inline-block;width:150px;margin:6px;padding:8px 10px;border-radius:8px;background:#D9F0EF;border:1px solid #C3E3E1;vertical-align:top}
.tile .k{font-size:10.5px;color:#4A5A70;font-weight:600}.tile .v{font-size:20px;font-weight:700;margin-top:2px}.tile .d{font-size:10px;color:#4A5A70}
.bars{display:inline-block;vertical-align:top;margin:6px 10px;padding:8px 10px;border:1px solid #D9E0E8;border-radius:8px;width:320px}
.bars .row{display:flex;align-items:center;font-size:10.5px;margin:3px 0}.bars .lab{width:90px}.bars .b{height:12px;background:#9DBFE3;border-radius:2px}
"""

LEG = ('<div class="legend"><span><i style="background:#FFF4D6"></i>user input</span><span><i style="background:#E2F0E8"></i>calculated / snapshot</span>'
       '<span><i style="background:#F7F9FC"></i>lookup (display)</span><span><i style="background:#F8DCDD"></i>hard stop</span>'
       '<span><i style="background:#FCECC9"></i>warning / attention</span><span><i style="background:#D9F0EF"></i>totals</span>'
       '<span><i style="background:#DCEAF7"></i>header</span><span style="margin-left:auto"><b>Illustrative Excel Layout - Mock-up</b> · sample / masked data · not a live screenshot</span></div>')


def colname(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def grid(widths, rows, start_row=1):
    """rows: list of lists of (text, cls, colspan) or plain strings."""
    h = ['<table class="g"><colgroup><col style="width:30px">' + "".join(f'<col style="width:{w}px">' for w in widths) + "</colgroup>"]
    h.append('<tr class="cl"><td class="rh"></td>' + "".join(f"<td>{colname(i)}</td>" for i in range(len(widths))) + "</tr>")
    for ri, r in enumerate(rows):
        cells = []
        used = 0
        for c in r:
            if isinstance(c, str):
                c = (c, "w", 1)
            elif len(c) == 2:
                c = (c[0], c[1], 1)
            t, cls, span = c
            used += span
            sp = f' colspan="{span}"' if span > 1 else ""
            cells.append(f'<td class="{cls}"{sp}>{html.escape(str(t))}</td>')
        while used < len(widths):
            cells.append('<td class="w"></td>'); used += 1
        h.append(f'<tr><td class="rh">{start_row + ri}</td>' + "".join(cells) + "</tr>")
    h.append("</table>")
    return "".join(h)


def page(name, title, book, fx_name, fx, body, tabs, active):
    t = "".join(f'<span class="{"a" if x == active else ""}">{html.escape(x)}</span>' for x in tabs)
    doc = (f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>"
           f"<div class='win' id='shot'><div class='bar'><span><b>{html.escape(book)}</b> - {html.escape(title)}</span>"
           f"<span class='tag'>Illustrative Excel Layout - Mock-up</span></div>"
           f"<div class='fx'><div class='nm'>{html.escape(fx_name)}</div><div class='f'>{html.escape(fx)}</div></div>"
           f"{body}<div class='tabs'>{t}</div>{LEG}</div></body></html>")
    open(os.path.join(OUT, name + ".html"), "w").write(doc)


I, CA, LK, EX, PE, H, H2, TO, OK = "in", "ca", "lk", "ex", "pe", "h", "h2", "to", "ok"


def grn():
    w = [104, 74, 76, 92, 100, 92, 44, 52, 52, 54, 52, 78, 72, 62, 64, 64, 66, 200]
    hdr = ["Lot ID", "GRN Date", "Supplier ID", "Supplier Inv No", "PO Line ID", "Material ID", "Thk", "Length", "Width", "Form", "Sheets",
           "Received kg", "Theo kg", "Theo var", "Bill Rs/kg", "PO Rs/kg", "Excess kg", "Ready / issues"]
    rows = [[("GRN ENTRY PAD  -  type in cream cells; post only rows that read READY", "t", 12), ("Balances: OK  ·  built from register rows = live rows", "ok", 6)],
            [("One row per lot received. QC decision is recorded afterwards on the Lot QC pad - the lot cannot be issued until QC is ACCEPTED.", "s", 18)],
            [(x, H) for x in hdr],
            [("GP-2728-0118", CA), ("12-Apr-27", I), ("SUP-0007", I), ("INV/3381", I), ("PO/27-28/014-02", I), ("GPSP-1.6", CA), ("1.6", CA, 1), ("2500", I), ("1250", I), ("SHEET", I), ("40", I),
             ("1,575.500", I), ("1,570.000", CA), ("+0.4%", CA), ("74.50", I), ("74.50", CA), ("0.000", CA), ("READY", OK)],
            [("GP-2728-0119", CA), ("12-Apr-27", I), ("SUP-0007", I), ("INV/3381", I), ("PO/27-28/014-03", I), ("GPSP-1.2", CA), ("1.2", CA), ("2500", I), ("1250", I), ("SHEET", I), ("60", I),
             ("1,860.000", I), ("1,766.250", CA), ("+5.3%", PE), ("74.50", I), ("74.50", CA), ("110.000", PE), ("THEO VAR > 3%; EXCESS - decision needed before issue", PE)],
            [("CR-2728-0042", CA), ("12-Apr-27", I), ("SUP-0012", I), ("B-77120", I), ("", EX), ("CRCA-1.0", I), ("1.0", CA), ("2500", I), ("1250", I), ("SHEET", I), ("30", I),
             ("736.000", I), ("735.940", CA), ("+0.0%", CA), ("68.20", I), ("", CA), ("", CA), ("NO OPEN PO LINE FOR SUPPLIER + ITEM", EX)],
            [("", I) for _ in range(17)] + [("", LK)],
            [("Pad totals", TO, 11), ("4,171.500", TO), ("", TO, 6)],
            ]
    page("mock_grn", "GRN Entry pad", "1 ISKO Stores", "L7", "=IF(K7=0,\"\",ROUND(K7*H7*I7*G7*Density,3))   (pad formula - posted as a value)",
         grid(w, rows), ["README", "PO Entry", "GRN Entry", "Lot QC Entry", "RM Move Entry", "GRN Register", "Lot Balance", "M_Material", "Settings"], "GRN Entry")


def rmissue():
    w = [118, 70, 98, 118, 84, 98, 98, 92, 56, 52, 70, 74, 60, 72, 60, 76, 92, 74, 210]
    hdr = ["Move ID", "Date", "Movement", "Batch ID", "Material ID", "FIFO lot", "Lot ID", "Override reason", "Form", "Sheets", "Weight kg",
           "Avail kg", "Avail sh", "Stock effect", "Rate", "Value Rs", "Machine", "Issued to", "Ready / issues"]
    rows = [[("RM ISSUE & RETURN PAD", "t", 6), ("Balances: OK  (register rows at refresh = live rows)", "ok", 7), ("Max for lot shown in Avail kg / Avail sh", "s", 6)],
            [("ISSUE creates a batch (Batch ID = Move ID). Returns, offcuts, damaged and stores scrap must quote an open batch of the same lot.", "s", 19)],
            [(x, H) for x in hdr],
            [("ISS-2728-004512", CA), ("14-Apr-27", I), ("ISSUE", I), ("ISS-2728-004512", CA), ("GPSP-1.6", I), ("GP-2728-0101", LK), ("GP-2728-0101", I), ("", I), ("SHEET", I), ("12", I),
             ("471.000", I), ("1,245.600", LK), ("52", LK), ("-471.000", CA), ("74.50", CA), ("-35,089.50", CA), ("MC-01", I), ("PER-007", I), ("READY", OK)],
            [("ISS-2728-004513", CA), ("14-Apr-27", I), ("ISSUE", I), ("ISS-2728-004513", CA), ("GPSP-1.6", I), ("GP-2728-0101", LK), ("GP-2728-0101", I), ("", I), ("SHEET", I), ("40", I),
             ("1,570.000", I), ("774.600", LK), ("40", LK), ("-1,570.000", CA), ("74.50", CA), ("", CA), ("MC-03", I), ("PER-011", I), ("EXCEEDS AVAILABLE: max 774.600 kg / 40 sheets", EX)],
            [("ISS-2728-004514", CA), ("14-Apr-27", I), ("ISSUE", I), ("ISS-2728-004514", CA), ("HRCA-2.0", I), ("HR-2728-0033", LK), ("HR-2728-0036", I), ("", PE), ("SHEET", I), ("3", I),
             ("214.900", I), ("640.200", LK), ("9", LK), ("-214.900", CA), ("66.10", CA), ("-14,204.89", CA), ("MC-02", I), ("PER-003", I), ("NOT FIFO - override reason required", PE)],
            [("OFC-2728-000318", CA), ("14-Apr-27", I), ("RETURN-OFFCUT", I), ("ISS-2728-004377", I), ("GPSP-1.2", CA), ("", LK), ("GP-2728-0097", CA), ("", I), ("OFFCUT", I), ("0", I),
             ("18.400", I), ("max 22.150", LK), ("", LK), ("+18.400", CA), ("74.50", CA), ("1,370.80", CA), ("", I), ("PER-007", I), ("READY", OK)],
            [("", I) for _ in range(18)] + [("", LK)],
            [("Pad totals by lot are subtracted from Avail kg, so two pad rows cannot together exceed stock", "s", 19)],
            ]
    page("mock_rmissue", "RM Issue & Return pad", "1 ISKO Stores", "L5",
         "=IFERROR(INDEX(LotBal[Available kg],MATCH(G5,LotBal[Lot ID],0)),0)-SUMIFS(K$4:K4,G$4:G4,G5)",
         grid(w, rows), ["README", "GRN Entry", "Lot QC Entry", "RM Move Entry", "RM Move Register", "Lot Balance", "Cons Issue Entry", "Settings"], "RM Move Entry")


def prod():
    w = [118, 70, 44, 66, 82, 118, 98, 70, 98, 96, 46, 60, 76, 62, 82, 70, 46, 230]
    hdr = ["Prod ID", "Date", "Shift", "Machine", "Stage", "Batch ID", "Lot ID", "Product", "Part ID", "Drawing / Rev", "Qty", "Wt/pc kg",
           "Output kg", "Scrap kg", "Batch kg left", "Operator", "QC", "Ready / issues"]
    rows = [[("PRODUCTION ENTRY PAD", "t", 5), ("Batch list: X_Batch refreshed at [start of shift] - OK", "ok", 7), ("Batch ISS-2728-004512: issued 471.000 · returned 0 · output 0 · left 471.000", "lk", 6)],
            [("Cutting lines must quote an open batch issued to this machine. Weight per piece and drawing revision are frozen at posting.", "s", 18)],
            [(x, H) for x in hdr],
            [("PRD-2728-018230", CA), ("14-Apr-27", I), ("A", I), ("MC-01", I), ("CUTTING", CA), ("ISS-2728-004512", I), ("GP-2728-0101", CA), ("PRD-A12", I), ("A12-BP-01", I), ("DRG-1102 / C", CA),
             ("60", I), ("6.215", CA), ("372.900", CA), ("24.300", I), ("471.000", LK), ("PER-021", I), ("OK", I), ("READY", OK)],
            [("PRD-2728-018231", CA), ("14-Apr-27", I), ("A", I), ("MC-01", I), ("CUTTING", CA), ("ISS-2728-004512", I), ("GP-2728-0101", CA), ("PRD-A12", I), ("A12-SP-04", I), ("DRG-1107 / A", CA),
             ("40", I), ("2.010", CA), ("80.400", CA), ("11.900", I), ("73.800", LK), ("PER-021", I), ("OK", I), ("OVER BATCH by 18.5 kg (left 73.800 incl. this pad)", EX)],
            [("PRD-2728-018232", CA), ("14-Apr-27", I), ("B", I), ("MC-05", I), ("BENDING", CA), ("", I), ("", CA), ("PRD-A12", I), ("A12-BP-01", I), ("DRG-1102 / C", CA),
             ("60", I), ("6.215", CA), ("372.900", CA), ("", I), ("", LK), ("PER-030", I), ("OK", I), ("POSSIBLE DUPLICATE - same key posted today", PE)],
            [("", I) for _ in range(17)] + [("", LK)],
            [("Pad totals", TO, 12), ("826.200", TO), ("36.200", TO), ("", TO, 4)],
            ]
    page("mock_prod", "Production Entry pad", "2 ISKO Production", "M5", "=ROUND(K5*IF(ISNUMBER(OverrideWt),OverrideWt,INDEX(M_Part[Wt/pc kg],MATCH(I5,M_Part[Part ID],0))),3)",
         grid(w, rows), ["README", "Prod Entry", "Rej-Rework Entry", "Batch Close Entry", "Production Register", "Batch Balance", "WIP Balance", "M_Part", "X_Batch"], "Prod Entry")


def jwc():
    w = [120, 88, 70, 70, 90, 116, 52, 54, 54, 56, 60, 70, 74, 48, 60, 120]
    hdr = ["JWC Line ID", "Coater", "JWC Date", "Product", "Part ID", "Batch ID", "Sent", "Back OK", "Defect", "Direct DC", "Pending", "Pending kg",
           "Pending sq ft", "Days", "Age", "Status"]
    rows = [[("OPEN JWC / AT-COATER BALANCE  (calculated - no entry on this sheet)", "t", 10), ("Rebuilt after each posting  ·  Open lines only", "ok", 6)],
            [("By coater", H2, 2), ("Open lines", H2), ("Pending qty", H2), ("Pending kg", H2), ("Pending sq ft", H2), ("> 15 days", H2), ("Opening (cut-over)", H2, 2), ("", "w", 7)],
            [("COATER-1", "w", 2), ("38", TO), ("1,214", TO), ("3,982.4", TO), ("8,315.0", TO), ("6", PE), ("0", "w", 2)],
            [("COATER-2", "w", 2), ("11", TO), ("402", TO), ("1,108.7", TO), ("2,240.6", TO), ("0", TO), ("0", "w", 2)],
            [(x, H) for x in hdr],
            [("JWC/27-28/0412-01", "w"), ("COATER-1", "w"), ("02-Apr-27", "w"), ("PRD-A12", "w"), ("A12-BP-01", "w"), ("ISS-2728-004377", "w"), ("120", "r"), ("80", "r"), ("4", "r"),
             ("0", "r"), ("36", CA), ("223.740", CA), ("498.6", CA), ("19", PE), ("16-30 d", PE), ("PART RETURNED", "w")],
            [("JWC/27-28/0412-02", "w"), ("COATER-1", "w"), ("02-Apr-27", "w"), ("PRD-A12", "w"), ("A12-SP-04", "w"), ("ISS-2728-004377", "w"), ("200", "r"), ("0", "r"), ("0", "r"),
             ("200", "r"), ("0", CA), ("0.000", CA), ("0.0", CA), ("", "w"), ("", "w"), ("CLOSED - drops off", "lk")],
            [("JWC/27-28/0419-01", "w"), ("COATER-2", "w"), ("09-Apr-27", "w"), ("PRD-B07", "w"), ("B07-CB-02", "w"), ("ISS-2728-004401", "w"), ("90", "r"), ("40", "r"), ("0", "r"),
             ("0", "r"), ("50", CA), ("146.500", CA), ("310.0", CA), ("12", "w"), ("0-15 d", "w"), ("PENDING", "w")],
            [("JWC/27-28/0421-03", "w"), ("COATER-1", "w"), ("11-Apr-27", "w"), ("PRD-A12", "w"), ("A12-BP-01", "w"), ("ISS-2728-004430", "w"), ("60", "r"), ("62", EX), ("0", "r"),
             ("0", "r"), ("-2", EX), ("-12.430", EX), ("-27.7", EX), ("10", "w"), ("0-15 d", "w"), ("BREACH - return > sent", EX)],
            [("Pending = Sent - Back OK - Defect back - Scrapped at coater - Direct dispatch (DC lines quoting the JWC line). Coater opening at cut-over is a one-time OPENING movement.", "s", 16)],
            ]
    page("mock_jwc", "Open JWC Balance", "3 ISKO Dispatch & Challan", "K6", "Pending Qty = [Sent] - [Back OK] - [Defect] - [Direct DC]   (Power Query group-by, rebuilt on posting)",
         grid(w, rows), ["README", "Req Entry", "JWC Entry", "Coater Move Entry", "DC Entry", "JWC Register", "Open JWC Balance", "FG Balance", "DC Register"], "Open JWC Balance")


def dc():
    w = [124, 74, 92, 116, 120, 76, 56, 58, 62, 76, 78, 64, 84, 78, 230]
    hdr = ["DC Line ID", "Product", "Part ID", "JWC Line ID", "Batch ID", "Type", "Qty", "UOM", "Wt/unit", "Wt kg", "Challan qty", "Rate", "Value Rs", "FG avail", "Ready / issues"]
    rows = [[("DELIVERY CHALLAN PAD  (one challan at a time: header, then lines)", "t", 9), ("FG Balance: OK  ·  JWC Balance: OK", "ok", 6)],
            [("Challan No", H2), ("DC/27-28/0191", CA), ("Date", H2), ("15-Apr-27", I), ("Customer", H2), ("CUST-01", I, 2), ("UOM", H2), ("QTY", CA), ("Dispatch from", H2), ("FACTORY", I), ("Vehicle", H2), ("GJ06-XX-0000", I), ("LR / Invoice", H2), ("[ref] / [ref]", I)],
            [("e-Way Bill", H2), ("[number]", I), ("Driver", H2), ("[name]", I), ("Status", H2), ("DISPATCHED", CA, 2), ("Lines", H2), ("3", TO), ("Header copied onto every line at posting (flat line register)", "s", 6)],
            [(x, H) for x in hdr],
            [("DC/27-28/0191-01", CA), ("PRD-A12", I), ("(set)", I), ("", I), ("", I), ("SALE", I), ("40", I), ("QTY", CA), ("41.950", CA), ("1,678.000", CA), ("40", CA), ("173.24", CA), ("2,90,696.72", CA), ("64", LK), ("READY", OK)],
            [("DC/27-28/0191-02", CA), ("PRD-B07", I), ("(set)", I), ("", I), ("", I), ("SALE", I), ("30", I), ("QTY", CA), ("28.400", CA), ("852.000", CA), ("30", CA), ("165.10", CA), ("", CA), ("22", LK), ("ABOVE FG AVAILABLE: max 22", EX)],
            [("DC/27-28/0191-03", CA), ("PRD-A12", I), ("A12-SP-04", I), ("JWC/27-28/0419-01", I), ("ISS-2728-004401", CA), ("SALE", I), ("20", I), ("QTY", CA), ("2.010", CA), ("40.200", CA), ("20", CA), ("173.24", CA), ("6,964.25", CA), ("50 at coater", LK), ("READY (from coater: JWC pending 50)", OK)],
            [("", I) for _ in range(14)] + [("", LK)],
            [("Challan totals", TO, 6), ("90", TO), ("", TO, 2), ("2,570.200", TO), ("", TO, 2), ("2,97,660.97", TO), ("", TO, 2)],
            ]
    page("mock_dc", "DC document pad", "3 ISKO Dispatch & Challan", "K5", "=IF(H5=\"KG\",J5,IF(H5=\"TON\",ROUND(J5/1000,3),G5))   UOM from Customer-Product table, frozen at posting",
         grid(w, rows), ["README", "Req Entry", "JWC Entry", "Coater Move Entry", "DC Entry", "DC Register", "FG Balance", "Open JWC Balance", "M_Customer"], "DC Entry")


def mis():
    w = [150, 110, 110, 70, 80, 90, 120]
    rf = [[("REFRESH STATUS", H2, 7)],
          [(x, H) for x in ["Source extract", "Refreshed", "Source saved", "Rows", "Distinct IDs", "Control kg", "Status"]],
          [("Stores - X_RMMove", "w"), ("[run time]", "w"), ("[file time]", "w"), ("6,118", "r"), ("6,118", "r"), ("412,906.2", "r"), ("OK", OK)],
          [("Production - X_Prod", "w"), ("[run time]", "w"), ("[file time]", "w"), ("23,544", "r"), ("23,544", "r"), ("1,206,331.0", "r"), ("OK", OK)],
          [("Dispatch - X_DCLine", "w"), ("[run time]", "w"), ("[file time]", "w"), ("1,302", "r"), ("1,302", "r"), ("611,870.5", "r"), ("STALE - source older than 24 h", PE)],
          [("Dispatch - X_JWCLine", "w"), ("[run time]", "w"), ("[file time]", "w"), ("7,015", "r"), ("7,015", "r"), ("402,118.9", "r"), ("OK", OK)]]
    tiles = [("RM stock (MT)", "132.4", "of which > 90 days: 11.2"), ("Cutting output (MT)", "108.9", "month"), ("Cutting yield", "91.6%", "closed batches"),
             ("Scrap loss (Rs lakh)", "3.42", "month"), ("WIP (MT)", "41.7", "calculated"), ("At coaters (MT)", "5.09", "6 lines > 15 days"),
             ("FG available (MT)", "18.3", "calculated"), ("Dispatched (MT)", "101.6", "Rs 1.68 cr · 165 Rs/kg")]
    tl = "".join(f"<div class='tile'><div class='k'>{a}</div><div class='v'>{b}</div><div class='d'>{c}</div></div>" for a, b, c in tiles)
    ex = [[("OPEN EXCEPTIONS BY OWNER", H2, 7)],
          [(x, H) for x in ["Code", "Exception", "Severity", "Owner", "Count", "Oldest (days)", ""]],
          [("PRD-04", "w"), ("Closed batch variance above tolerance", "w"), ("HIGH", PE), ("Production", "w"), ("3", "r"), ("2", "r"), ("", "w")],
          [("DSP-01", "w"), ("JWC line at coater > CoaterAge", "w"), ("MEDIUM", "w"), ("Dispatch", "w"), ("6", "r"), ("19", "r"), ("", "w")],
          [("STK-03", "w"), ("Lot in QC HOLD > 3 days", "w"), ("HIGH", PE), ("QC", "w"), ("1", "r"), ("4", "r"), ("", "w")],
          [("SYS-01", "w"), ("Source book stale (Dispatch)", "w"), ("BLOCKER", EX), ("MIS owner", "w"), ("1", "r"), ("0", "r"), ("", "w")]]
    bars = "".join(f"<div class='row'><span class='lab'>{m}</span><span class='b' style='width:{v}px'></span><span style='margin-left:6px'>{t}</span></div>"
                   for m, v, t in [("Jan", 150, "96.1"), ("Feb", 162, "103.8"), ("Mar", 171, "109.2"), ("Apr", 159, "101.6")])
    body = (f"<div style='padding:8px 10px'><div style='font-weight:700;font-size:14px'>MIS CONTROL TOWER</div>"
            f"<div style='font-size:11px;color:#4A5A70'>Period: [month]  ·  Customer: All  ·  Material: All  ·  Machine: All  ·  Product: All   (slicers)</div></div>"
            f"<div style='padding:0 6px'>{grid(w, rf)}</div>"
            f"<div style='padding:4px 4px 0'>{tl}</div>"
            f"<div style='display:flex;gap:8px;padding:4px 6px 8px'><div>{grid(w, ex)}</div>"
            f"<div class='bars'><div style='font-weight:700;font-size:11px;margin-bottom:4px'>Dispatched MT by month (summary table, 4 rows)</div>{bars}</div></div>")
    page("mock_mis", "Control Tower", "4 ISKO MIS", "B3", "Tiles read small Power Query summary tables; nothing on this sheet scans a register",
         body, ["README", "Control Tower", "Exceptions", "Refresh Log", "Stock & WIP", "Production & Yield", "Scrap & Yield", "Lot Trace", "DC Trace", "MRM Monthly"], "Control Tower")


if __name__ == "__main__":
    grn(); rmissue(); prod(); jwc(); dc(); mis()
    print("mockup html written to", OUT)
