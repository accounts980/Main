# Sheet-by-sheet rebuild decisions for every sheet in the four current workbooks.
# Key = (workbook, current sheet name). Measured figures (used range, rows, formulas)
# are merged in at build time from the audit JSON produced by tools/audit_*.py.
#
# Tuple: (purpose, decision, reason_and_control_impact, proposed_final, target_rows,
#         inputs_outputs, dependencies, risks_controls)

RP = "Removed from Phase 1"

S = {}

def add(wb, sheet, purpose, decision, reason, proposed, target, io, deps, risk):
    S[(wb, sheet)] = dict(purpose=purpose, decision=decision, reason=reason, proposed=proposed,
                          target=target, io=io, deps=deps, risk=risk)

# ---------------------------------------------------------------- 1 STORES
W = "Stores"
add(W, "START", "Navigation page with hyperlinks to every sheet and a process description.",
    "Simplify", "Keep a plain README; remove links to removed sheets (requisition, print, count) and the print area.",
    "README", "n/a", "Out: navigation only", "None", "Stale instructions if not updated with each release; version the README.")
add(W, "1.1 Requisition", "Purchase requisition register (requested by, item, kg, stores decision, PO link). 0 data rows; 2,401 formula cells pre-filled.",
    RP, "Scope override 1. PO is raised directly from need shown in MIS RM Requirement and the re-order flags. Control impact: purchase authorisation moves to the PO approval field; no PO depends on a requisition reference.",
    "None", "0", "None", "Referenced by 1.2 (Req No/Item/Line, Req kg), Lists AL:AR, names ReqFrom/RqItem/RqLen/RqWid/RqLine/TolReqPO, MIS D_Requisition, 4.10 Checks rows 36 and 38, 4.7 DC Trace chain.",
    "All dependants must be deleted in the same release or #REF! errors will appear.")
add(W, "1.2 Purchase Order", "PO lines with header fields repeated on the first line, requisition link, received/pending progress.",
    "Simplify", "Split into PO header and PO line registers; remove requisition columns and the 'Add to PO' mechanic; received/pending become a bounded PO Status view.",
    "S-PO Entry (pad) + tblPOHead + tblPOLine + PO Status (open lines)", "100,000 lines design limit (actual ~1-2k/yr expected)",
    "In: supplier, item, qty, rate, dates. Out: open PO lines to GRN; PO status to MIS.", "M_Supplier, M_Material, M_Consumable; GRN register for receipts.",
    "TODAY() in 202 rows and 1,597 running-range formulas today; target: no volatile functions, PO Line ID permanent.")
add(W, "1.3 Print PO", "PO print layout with Gmail compose link and 'Copy page' link.",
    RP, "Scope overrides 2 and 3 (print and mail). Control impact: none on stock; the PO register remains the record. Management decides how a PO document reaches the supplier outside the ERP.",
    "None", "0", "None", "Lists P:AA mail lists, MailTo4, POMailTo, Print_Area name, printerSettings part.", "Remove print area, printer settings, Gmail hyperlinks and mail names together.")
add(W, "1.4 RM Receipt GRN", "One row per received lot: size, sheets, kg, theoretical kg, rate, PO link, QC, excess, accepted kg/sheets.",
    "Simplify", "Keep receipt facts; move QC and excess decisions to an append-only QC decision log so decisions after receipt do not overwrite history; Lot ID becomes permanent.",
    "S-GRN Entry (pad) + tblGRN + tblLotQC + Lot Balance", "100,000 lots design limit",
    "In: receipt, bill, QC. Out: issue-eligible lots, PO receipt progress, lot rate.", "tblPOLine, M_Material, M_Supplier, Settings tolerances.",
    "RM ID built with a running COUNTIF (changes if rows are sorted/inserted); 696 running-range formulas; QC Status blank on all 187 lots.")
add(W, "1.5 Consumable Receipt", "Consumable receipts (item, qty, bill rate, supplier, opening lines).",
    "Simplify", "Entry pad + values register; opening balances posted once as OPENING movements.",
    "S-Cons Receipt (pad) + tblConsRcpt", "100,000", "In: item, qty, rate, bill. Out: consumable balance and average rate.",
    "M_Consumable, M_Supplier", "413 self-table aggregate formulas today; target none in the register.")
add(W, "1.6 RM Issue & Return", "RM issue, return, offcut return, scrap and sample movements with FIFO lot auto-pick, availability and checks.",
    "Simplify", "Largest Stores hotspot. Rebuild as entry pad + values-only movement register + bounded lot balance. Hard stops retained at entry; historic rows never recalculate.",
    "S-RM Move Entry (pad) + tblRMMove + Lot Balance + Batch list", "100,000 movements (actual ~6k/yr)",
    "In: material, lot, kg, sheets, machine, person, product. Out: batches to Production; lot balances; issue value.",
    "tblGRN, tblLotQC, tblAdjust, M_Machine (imported), M_Person.",
    "57,683 formula cells; 13,572 running-range and 13,572 self-table SUMIFS formulas (grow with the square of row count); TODAY() in 3,393 rows; FIFO lot pick recalculated for history.")
add(W, "1.6b Print RM Issue Slip", "Printable RM issue slip (batch slip) with NOW() and copy-page link.",
    RP, "Scope override 2. Control impact: the batch number is no longer carried on paper; Production picks the batch from the imported open-batch list by machine.",
    "None", "0", "None", "SlipIssue name, RMSlipList, Print_Area.", "Shop-floor must confirm batch identification without a slip (Decision D-07).")
add(W, "1.7 Consumable Issue", "Consumable issues/returns with BOM qty, variance, stock check, value.",
    "Simplify", "Entry pad + values register; BOM vs actual moves to MIS; stock check becomes a hard stop at entry against the consumable balance.",
    "S-Cons Issue (pad) + tblConsIssue", "100,000", "In: item, qty, returned qty, person, where used. Out: consumable balance, cost by where-used.",
    "M_Consumable, M_WhereUsed, M_Person", "1,724 TODAY() cells and 862 running-range formulas today; 734 pre-filled blank rows.")
add(W, "1.8 Print Issue Slip", "Consumable issue slip print layout with Gmail link.",
    RP, "Scope overrides 2 and 3. Control impact: none; the issue register is the record.", "None", "0", "None",
    "IssuePick, MailTo1, Print_Area, printerSettings.", "Remove with its names and list columns.")
add(W, "1.9 Stock Adjustment", "RM and consumable adjustments; current lines are count-driven corrections (217 lines, net -97,471 kg, no reason code or entered-by on any line).",
    "Simplify", "Redesign as an Authorised Stock-Adjustment movement (reason code, reference, requester, approver, value). Count-difference reasons are removed; this is not a count module.",
    "S-Adjust Entry (pad) + tblAdjust", "100,000 (expected low)", "In: lot/item, +/- qty, reason, approver. Out: lot and consumable balances.",
    "tblGRN, M_Consumable, M_Reason", "Today none of the 217 adjustments has a reason code or an approver; target makes both hard stops.")
add(W, "1.10 Stock Count", "Physical stock count entry (273 lines) compared with book stock on the dashboard.",
    RP, "Scope override 5. Control impact: stock is controlled through movement validation, QC eligibility and authorised adjustments only.",
    "None", "0", "None", "MIS D_Count/qCount, 4.3 Stock Summary, 4.8 MRM 'Counted (kg)', adjustment text.", "Historic count rows are archived read-only with the 2026-27 files, not migrated.")
add(W, "1.11 RM Stock", "Lot-by-lot balance, value, ageing and grade summary (spill formulas).",
    "Simplify", "Keep as the operational Lot Balance (needed for availability). Calculated for open lots only, refreshed on posting.",
    "Lot Balance (calculated)", "Open lots only (~300-1,000)", "In: GRN, QC, movements, adjustments. Out: available kg/sheets by lot; ageing.",
    "tblGRN, tblLotQC, tblRMMove, tblAdjust", "ANCHORARRAY/FILTER/SORT spill formulas need Excel 2021/365; replace with Power Query output or bounded SUMIFS.")
add(W, "1.12 Consumable Stock", "Consumable stock, value, reorder status; Gmail/Copy page.",
    "Simplify", "Keep as Consumable Balance with reorder and negative flags; remove mail and print area.",
    "Consumable Balance (calculated)", "Active items (~250-500)", "Out: stock, value, reorder flag.", "tblConsRcpt, tblConsIssue, tblAdjust, M_Consumable",
    "Negative stock possible today (2 items); target hard stop at issue.")
add(W, "1.13 Re-order Alert", "List of RM and consumables below minimum, with Gmail 'send alert'.",
    "Combine", "Reorder flags live on the two balance tables and on the MIS exception list; separate alert sheet and mail removed.",
    "Flags on Lot/Consumable Balance; MIS exception", "n/a", "Out: reorder exceptions", "Balances", "Mail-based alert removed (scope override 3).")
add(W, "1.14 Stores Dashboard", "Monthly Stores KPIs (stock, received, issued, turnover, ageing).",
    "Convert to report", "Move KPIs to the MIS Control Tower so there is one stock number; Stores keeps operational balances only.",
    "MIS Control Tower (Stores tiles)", "n/a", "Out: KPIs", "MIS extracts", "Duplicate KPI definitions removed.")
add(W, "M_Supplier", "Supplier master (Tally ledger mapping, GSTIN, terms, email).",
    "Keep", "Add a stable Supplier ID and Active flag; email kept only as an optional plain reference field (no hyperlink, no mail use).",
    "M_Supplier", "<1,000", "Out: supplier dropdowns, PO/GRN snapshots", "None", "Name edits must not change history: PO/GRN store Supplier ID and a name snapshot.")
add(W, "M_Material", "Material master with grade, thickness, standard size/weight, average issue, min level, live stock.",
    "Simplify", "Keep identity and standards; remove live stock/avg/reorder formula columns (moved to Lot Balance and MIS).",
    "M_Material", "<2,000", "Out: material IDs, standard sheet weight, min level", "M_Grade", "992 formula cells today on 69 materials.")
add(W, "M_Grade", "Grade codes used as the prefix of RM IDs.", "Keep", "Small stable list.", "M_Grade", "<100", "Out: grade code", "None", "Code changes must never re-key existing lots.")
add(W, "M_Consumable", "Consumable item master with live received/issued/stock formulas.",
    "Simplify", "Keep identity, unit, category, min level; live quantities move to Consumable Balance.",
    "M_Consumable", "<5,000", "Out: item list", "M_Cons_Category", "2,684 formula cells on 97 items today.")
add(W, "M_Cons_Category", "Consumable categories and code prefixes.", "Keep", "Small list.", "M_Cons_Category", "<100", "Out: category", "None", "None")
add(W, "M_Product", "Stores copy of products/jobs (249 rows; manual copy from Production on 03-Oct-2026).",
    "Combine", "Product master has three diverging copies (Stores 249, Production 269, Dispatch 278). Stores imports the Production-owned master.",
    "X_Product (import)", "<5,000", "In: Production product master", "Production M_Product", "24 products exist in Production but not in Stores today.")
add(W, "M_WhereUsed", "Machines and departments a consumable can be issued to.",
    "Keep", "Keep as the cost-centre list; machine names come from the imported Machine master to avoid a second spelling.",
    "M_WhereUsed", "<200", "Out: where-used list", "X_Machine", "Machine spelling drift between books.")
add(W, "M_Person", "People who receive issues.", "Combine", "One People master (owned by Production, imported into Stores and Dispatch) with Person ID and Active flag.",
    "X_People (import)", "<500", "In: Production M_Worker", "Production", "Person names typed differently in each book today.")
add(W, "M_Reason", "Reason list for rate, adjustment and other dropdowns.",
    "Simplify", "Keep rate and adjustment reasons; remove count-difference reasons.", "M_Reason", "<200", "Out: reason codes", "None", "Reason codes must be codes, not free text, for MIS grouping.")
add(W, "M_Contacts", "Email contacts and groups for Gmail links.", RP, "Scope override 3 (mail).", "None", "0", "None",
    "MailTo/MailPick names and Lists columns", "Delete together with all mail names.")
add(W, "Settings", "Thresholds, dates, prefixes, company print details, requisition settings.",
    "Simplify", "Keep tolerances, FY, prefixes, LockDate and a typed AsOnDate; remove print/mail fields (CoAddr, CoEmail), ReqFrom, TolReqPO and TODAY()-based cells.",
    "Settings", "<100", "Out: named parameters", "None", "Settings edits must be logged in the Change Log.")
add(W, "Lists", "Hidden helper lists built with spill formulas (open lots, POs, mail lists, requisition lists).",
    "Simplify", "Replace spill formulas with master-table columns and Power Query outputs; keep the sheet visible and labelled.",
    "Lists (visible, no spill formulas)", "<5,000", "Out: dropdown sources", "Masters, balances", "35 ANCHORARRAY-based names fail in Excel 2016/2019.")

# ---------------------------------------------------------------- 2 PRODUCTION
W = "Production"
add(W, "START", "Navigation page.", "Simplify", "Plain README; remove links to removed sheets.", "README", "n/a", "Navigation", "None", "Keep in step with releases.")
add(W, "2.1 Production Entry", "Every production line: stage, machine, batch, lot, product, part, qty, weight, scrap, batch status, QC.",
    "Simplify", "Largest hotspot in the system (203,113 formula cells). Rebuild as entry pad + values-only register + bounded batch balance. Snapshot Wt/pc and lot at posting.",
    "P-Prod Entry (pad) + tblProd + Batch Balance", "100,000 lines (actual ~23k/yr)",
    "In: shift, machine, batch, product, part, qty, scrap. Out: output kg, batch consumption, WIP movements, JWC availability.",
    "X_Batch (Stores), M_Part, M_Product, M_Machine",
    "37,614 self-table SUMIFS + 25,076 running-range formulas; 102,808 formulas read Power Query tables that refresh every 5 minutes; 1,132 pre-filled blank rows.")
add(W, "2.2 Rejection & Rework", "In-process rejections with reason, action, rework lot/kg/OK qty.",
    "Simplify", "Event register: REJECT, REWORK-OK, REWORK-SCRAP lines linked to the original Prod ID; original never overwritten.",
    "P-Rej Entry (pad) + tblRejRework", "100,000", "In: source Prod ID, qty, reason, disposition. Out: QC OK qty, WIP reduction, scrap.",
    "tblProd, M_Reason, M_Stage", "Only 13 of 68 rejections reference a Prod ID today; target makes it mandatory for in-process rejections.")
add(W, "2.3 Machine Log", "Breakdown/maintenance log with hours.", RP,
    "0 rows entered in six months and capacity figures are marked ASSUMED. Recommend Phase 2 unless Management requires downtime KPIs now (Decision D-09).",
    "None (Phase 2: tblDowntime)", "0", "None", "MIS D_MachineLog, 2.6, 2.7 downtime tiles", "Downtime KPIs disappear from dashboards until Phase 2.")
add(W, "2.4 WIP Count", "WIP physical count at month end (1,503 lines).", RP,
    "Scope override 4. WIP becomes transaction-calculated from posted production, rejection and JWC movements.",
    "WIP Balance (calculated) - see Section 13", "n/a", "None", "MIS D_WIPCount, 4.3 Stock Summary", "Historic count lines archived, not migrated.")
add(W, "2.5 Daily Production", "Output by machine for a picked date plus every line of that day; Gmail/Copy page.",
    "Simplify", "Keep as a simple daily view (PivotTable on the register filtered by date); remove mail and print area.",
    "Daily Output (PivotTable)", "n/a", "Out: daily view", "tblProd", "Spill formulas replaced with a PivotTable refreshed on demand.")
add(W, "2.6 Machine Report", "Machine output by month and utilisation against capacity.",
    "Remove from Phase 1", "Utilisation depends on assumed capacities; machine output by month is kept in MIS Production report. Utilisation returns in Phase 2 with real capacity data.",
    "MIS Production & Yield (output by machine)", "n/a", "None", "M_Machine capacity", "Decision D-09.")
add(W, "2.7 Production Dashboard", "Monthly stage output, rejection, downtime, utilisation tiles.",
    "Convert to report", "Move to MIS so production KPIs have one definition.", "MIS Control Tower / Production & Yield", "n/a", "KPIs", "MIS extracts", "Duplicate definitions removed.")
add(W, "2.8 Batch Balance", "Per batch: issued, returned, output, scrap, variance, status; batch drill-down.",
    "Keep", "Keep as the single owner of batch balance; add an append-only Batch Close register with variance reason and approver.",
    "Batch Balance (calculated) + tblBatchClose", "Open batches (~500-2,000)", "In: X_Batch, tblProd, tblBatchClose. Out: batch status to MIS.",
    "X_Batch (Stores), tblProd", "MIS currently recomputes batches separately (qBatch) - two versions of the same number.")
add(W, "M_Machine", "Machines, process (stage), capacity, status.", "Keep", "Machine ID and process drive Stage; capacity optional until Phase 2.",
    "M_Machine (owner for all books)", "<100", "Out: machine list to Stores, MIS", "None", "Single machine spelling across books.")
add(W, "M_Worker", "Operators, helpers, QC, maintenance.", "Keep", "Becomes the single People master (Person ID, role, active), imported by Stores and Dispatch.",
    "M_People", "<1,000", "Out: people list", "None", "Never delete; set Active = No.")
add(W, "M_Product", "Products/models with customer, category, BOM model, set weight, made kg.",
    "Keep", "Single Product master for all books (Product ID, Customer ID, category, coating route flag, BOM model, active). Remove 'Made kg' formula (report).",
    "M_Product (owner)", "<5,000", "Out: product list to Stores, Dispatch, MIS", "M_BOM_Model", "Three copies today with different row counts.")
add(W, "M_Part", "Weight per piece by Product + Part, drawing, material.",
    "Simplify", "Part master with Part ID, drawing no, revision, material, thickness, Wt/pc and effective date; weight snapshot taken at posting.",
    "M_Part", "<20,000", "Out: Wt/pc, drawing", "M_Product, M_Material (import)", "6,630 formula cells in the master today; Wt/pc edits currently change historic output kg.")
add(W, "M_BOM", "BOM lines: drawing, revision, material, size, blank weight, qty/set, costable, variant.",
    "Keep", "Keep with revision and effective date; 'Data Fix' migration column folded into Remarks.",
    "M_BOM", "<20,000", "Out: set weight, RM requirement", "M_Product, M_Material", "BOM revisions must be added, not overwritten.")
add(W, "M_BOM_Model", "BOM editions and review dates.", "Keep", "Edition header with effective-from; review date typed, status a flag.",
    "M_BOM_Model", "<500", "Out: edition", "M_BOM", "10 models overdue for review today (MIS check).")
add(W, "M_Reason", "Rejection reasons and groups.", "Keep", "Coded list.", "M_Reason", "<300", "Out: reasons", "None", "None")
add(W, "M_Stage", "Places a rejection can be found besides machines.", "Combine", "Merge into a Stage & Route master (stage sequence per part family, FG/coating flag) needed for calculated WIP.",
    "M_Stage_Route", "<500", "Out: stage sequence", "M_Part", "Route master is new and needs Production sign-off (Decision D-05).")
add(W, "M_Contacts", "Mail contacts.", RP, "Scope override 3.", "None", "0", "None", "MailTo1, Lists I:J", "Delete with mail names.")
add(W, "Settings", "AsOnDate, density, LockDate, CheckFrom, BatchFrom; CELL('filename') path.",
    "Simplify", "Keep density, tolerances, LockDate; replace CELL()/TODAY() with typed values; source path typed once.",
    "Settings", "<100", "Out: parameters", "None", "CELL() and TODAY() are volatile.")
add(W, "D_Material", "Hidden Power Query copy of Stores material master.", "Simplify", "Keep as controlled import X_Material (visible, labelled, read-only by convention).",
    "X_Material", "<2,000", "In: Stores export", "Stores tblMaterial", "Refresh on open/every 5 minutes today; target manual controlled refresh.")
add(W, "D_Lots", "Hidden Power Query lot list with issued kg.", "Combine", "Replaced by X_Batch (batch carries lot); lot detail not needed separately in Production.",
    "X_Batch", "n/a", "In: Stores export", "Stores", "None")
add(W, "D_Issues", "Hidden Power Query issue/batch list with returns and batch kg.", "Simplify",
    "Keep as X_Batch: open and recently closed batches only (not full history), with issued/returned/scrap kg.",
    "X_Batch", "Open + last 90 days (~2,000)", "In: Stores export", "Stores tblRMMove", "Full-history import grows without limit today (3,022 rows).")
add(W, "Lists", "Hidden dropdown lists (spill formulas, mail lists).", "Simplify", "Visible list sheet without spill formulas or mail lists.",
    "Lists", "<5,000", "Dropdown sources", "Masters", "FILTER/SORTBY/UNIQUE need Excel 2021/365.")

# ---------------------------------------------------------------- 3 DISPATCH
W = "Dispatch"
add(W, "START", "Navigation page.", "Simplify", "Plain README.", "README", "n/a", "Navigation", "None", "None")
add(W, "3.1 Customer Requirement", "Customer schedules with versions, priority, status, dispatched and pending qty.",
    "Keep", "Keep version/revises logic; dispatched qty computed from DC lines carrying Req ID (new optional link).",
    "D-Req Entry (pad) + tblRequirement", "100,000", "In: customer, product, qty, date, priority. Out: pending requirement to MIS RM Requirement.",
    "M_Customer, X_Product, tblDCLine", "Dispatched qty is matched by customer+product today, not by requirement (Decision D-12).")
add(W, "3.2 JWC Header", "Job-work challan header (vendor, vehicle, e-way bill, purpose, shade, status, coater bill).",
    "Simplify", "Document pad (header + lines) posting to header and line registers; JWC number permanent and sequential per FY.",
    "D-JWC Entry (pad) + tblJWCHead", "100,000", "In: vendor, date, vehicle, purpose. Out: JWC totals.", "M_Vendor", "2,617 formula cells for 152 headers today.")
add(W, "3.3 JWC Lines", "Parts sent to coaters: product, drawing, batch, lot, qty, kg, sq ft, rate, cost, returns, pending, ageing.",
    "Simplify", "Second-largest hotspot (88,661 formula cells). Lines register stores snapshots (rate, sq ft, lot); returns/pending move to the bounded Open JWC Balance.",
    "tblJWCLine + Open JWC Balance", "100,000 lines (actual ~7k/yr)",
    "In: product, part, batch, qty, kg, sq ft/pc. Out: pending at coater, PC cost.", "X_ProdOutput (Production), M_Vendor rates, tblCoaterReturn, tblDCLine",
    "Coater rate and basis looked up live from M_Vendor today (history changes when the rate changes).")
add(W, "3.4 Print JWC", "Job-work challan print layout with Gmail link.", RP,
    "Scope overrides 2 and 3. Statutory job-work challan document must be produced outside the ERP (Decision D-03).",
    "None", "0", "None", "JWCMailTo, MailTo2, Print_Area", "Statutory document risk - see Risk R-14.")
add(W, "3.5 Coater Opening", "One-time opening quantities at each coater on go-live (empty).",
    "Simplify", "Keep as a one-time cut-over Opening Balance movement confirmed from coater statements; not a recurring count.",
    "tblCoaterMove (type OPENING)", "<1,000", "In: coater, product, part, qty, kg. Out: opening at coater.", "M_Vendor, X_Product", "Must be entered before go-live or at-coater balances start wrong.")
add(W, "3.6 Coater Return", "Coated parts back OK / defects, fault, decision, re-sent JWC, debit, status, Mail Coater link.",
    "Simplify", "Keep as coater movement register (BACK-OK, DEFECT-BACK, DIRECT-DISPATCH via DC); debit/claim and closure fields kept; Mail Coater removed.",
    "D-Coater Return (pad) + tblCoaterMove", "100,000", "In: JWC line, qty, kg, type, decision. Out: pending at coater, FG receipts, debit notes.",
    "tblJWCLine", "Mail Coater (100 formula cells) removed.")
add(W, "3.7 At Coater", "Material at each coater, pending by age (spill formulas).",
    "Simplify", "Keep as Open JWC Balance by coater and line with ageing buckets; transaction-calculated.",
    "Open JWC Balance (calculated)", "Open lines only", "Out: pending qty/kg/sq ft by coater, age", "tblJWCLine, tblCoaterMove, tblDCLine", "Spill formulas replaced.")
add(W, "3.8 FG Count", "Finished goods physical count (450 lines).", RP,
    "Scope override 5. Replacement: transaction-calculated FG Balance (Add replacement).",
    "FG Balance (calculated)", "Products (~300-1,000)", "None", "MIS D_FGCount, 4.1 'FG counted' tile", "FG receipt rule needs sign-off (Decision D-06).")
add(W, "3.9 DC Header", "Delivery challan header (customer, dispatch from, vehicle, driver, LR, invoice, e-way bill, status).",
    "Simplify", "Document pad posting header + lines; DC number permanent and sequential per FY.",
    "D-DC Entry (pad) + tblDCHead", "100,000", "In: customer, from, transport refs. Out: DC totals.", "M_Customer, M_Vendor (dispatch-from coater)", "None")
add(W, "3.10 DC Lines", "Delivery lines: JWC link, product, part, batch, qty, unit, challan qty, weight, rate, value.",
    "Simplify", "Lines register with UOM, weight and rate snapshots; FG availability hard stop at entry.",
    "tblDCLine + FG Balance", "100,000 lines", "In: product, qty, JWC line (if from coater). Out: dispatch qty/kg/value, FG reduction.",
    "M_Customer UOM, Customer Price table, FG Balance, Open JWC Balance",
    "Sale rate and set weight looked up live from M_Product today; 33 lines show Unit '0'.")
add(W, "3.11 Print DC", "Delivery challan print layout with Gmail link.", RP,
    "Scope overrides 2 and 3. Statutory delivery challan document must be produced outside the ERP (Decision D-03).",
    "None", "0", "None", "DCMailTo, MailTo1, Print_Area", "Statutory document risk - see Risk R-14.")
add(W, "3.12 Customer Summary", "Dispatch kg by customer and month.", "Convert to report", "Move to MIS.", "MIS Dispatch report", "n/a", "KPIs", "MIS", "None")
add(W, "3.13 Dispatch Dashboard", "Dispatch and powder-coating KPIs.", "Convert to report", "Move to MIS Control Tower.", "MIS Control Tower", "n/a", "KPIs", "MIS", "None")
add(W, "M_Customer", "Customers with GSTIN, Tally ledger, dispatch UOM, email.",
    "Keep", "Add Customer ID and Active; Dispatch UOM mandatory (8 of 15 blank today); email optional plain reference only.",
    "M_Customer (owner)", "<1,000", "Out: customer list, UOM", "None", "Blank UOM must be completed before go-live.")
add(W, "M_Vendor", "Powder coaters with basis and rates.", "Simplify",
    "Keep vendor identity; move rates to an effective-dated Coater Rate table so a rate change never alters posted JWC cost.",
    "M_Vendor + tblCoaterRate", "<200 / <2,000", "Out: vendor list, rate at JWC date", "None", "Live rate lookup today.")
add(W, "M_Product", "Dispatch copy of products with HSN, unit, set weight, sale rate, dispatched kg.",
    "Combine", "Use the Production-owned product master (import) plus a Dispatch-owned Customer Price & UOM table (effective-dated).",
    "X_Product + tblCustPrice", "<5,000", "In: Production master", "Production M_Product", "9 products exist only here (scrap sale items).")
add(W, "M_Worker", "Dispatch people (0 rows).", "Combine", "Import the single People master.", "X_People", "n/a", "In: Production", "Production", "None")
add(W, "M_Contacts", "Mail contacts.", RP, "Scope override 3.", "None", "0", "None", "Mail names", "Delete with mail names.")
add(W, "Settings", "AsOnDate, FY, DC/JWC prefixes and next numbers, GST print fields.",
    "Simplify", "Keep FY, prefixes and counters, tolerances, LockDate; remove print-only GST fields and CELL()/TODAY().",
    "Settings", "<100", "Parameters", "None", "Counter edits must be logged.")
add(W, "D_ProdLots", "Hidden Power Query: cut kg by product/part/lot from Production.", "Combine",
    "Replace with X_ProdOutput (QC-OK quantity available to send by Batch, Product, Part).", "X_ProdOutput", "Open items", "In: Production export", "Production", "Full history import today (4,320 rows).")
add(W, "D_ProdBatch", "Hidden Power Query: batch output and QC OK by product/part (1 header row, no data).", "Combine",
    "Merged into X_ProdOutput.", "X_ProdOutput", "n/a", "In: Production export", "Production", "Currently returns no rows - JWC 'sent > produced' check cannot work.")
add(W, "Lists", "Hidden dropdown lists incl. mail lists.", "Simplify", "Visible lists without spill formulas or mail columns.", "Lists", "<5,000", "Dropdown sources", "Masters", "None")

# ---------------------------------------------------------------- 4 MIS
W = "MIS"
add(W, "START", "Navigation and refresh instructions.", "Simplify", "README with the refresh procedure and a link to the Refresh Log.", "README", "n/a", "Navigation", "None", "None")
add(W, "4.1 Control Tower", "Five-filter KPI tiles, charts and tables; Gmail/Copy page.",
    "Simplify", "Rebuild as one director dashboard fed by small Power Query summary tables and PivotTables; refresh status panel; 'FG counted' tile replaced by calculated FG.",
    "MIS Control Tower", "n/a", "Out: KPIs, exceptions", "X_ extracts, Refresh Log", "51 spill references and FILTER formulas need Excel 2021/365.")
add(W, "4.2 RM Requirement", "Open requirements x BOM set weight less stock and on-order = what to buy.",
    "Keep", "Keep as the purchase-need report that replaces the requisition step; computed in Power Query; no mail.",
    "MIS RM Requirement", "n/a", "Out: buy list by material/thickness", "X_Requirement, X_BOM, X_LotBal, X_POStatus", "2 requirement lines with no BOM today.")
add(W, "4.3 Stock Summary", "RM, consumables, WIP and FG summary incl. 'latest physical count'.",
    "Simplify", "All balances transaction-calculated (RM, consumables, WIP, at coater, FG); physical-count block removed.",
    "MIS Stock & WIP Summary", "n/a", "Out: stock positions", "X_LotBal, X_ConsBal, X_WIP, X_AtCoater, X_FGBal", "Physical-count dependency removed.")
add(W, "4.4 Consumables", "Consumable use by category/where/person and BOM vs actual.", "Simplify", "PivotTable over X_ConsMove; BOM vs actual kept as optional report.",
    "MIS Consumables", "n/a", "Out: usage", "X_ConsMove", "None")
add(W, "4.5 Production & Yield", "Output by stage by month, yield, rejection by reason.", "Keep", "Power Query monthly summary + PivotTable; adds output by machine (from 2.6).",
    "MIS Production & Yield", "n/a", "Out: yield KPIs", "X_Prod, X_RejRework, X_BatchBal", "None")
add(W, "4.6 Lot Trace", "Trace one lot from PO to customer.", "Simplify", "Parameter-driven Power Query (pick Lot ID, Refresh) returning the chain; no FILTER formulas; no requisition link.",
    "MIS Lot Trace", "n/a", "Out: trace chain", "X_ extracts", "Trace limited to data captured (batch links mandatory from go-live).")
add(W, "4.7 DC Trace", "Trace a delivery challan back to JWC, batch, lot, GRN, PO and requisition.", "Simplify",
    "Parameter-driven Power Query; chain ends at PO (requisition removed).", "MIS DC Trace", "n/a", "Out: trace chain", "X_ extracts", "None")
add(W, "4.8 MRM Monthly", "Monthly review pack; Gmail/Copy page; 'Counted (kg)'.", "Simplify", "Keep the pack without mail, print area or count columns.",
    "MIS MRM Monthly", "n/a", "Out: monthly pack", "MIS summaries", "None")
add(W, "4.9 MIS Dashboard", "Second monthly KPI dashboard duplicating Control Tower tiles.", "Combine", "Merged into the Control Tower (one KPI definition each).",
    "MIS Control Tower", "n/a", "KPIs", "MIS", "Duplicate tiles removed.")
add(W, "4.10 Checks", "Cross-book checks with result, status, where-to-fix and 'Mail the team'.", "Keep",
    "Rebuild as the Exception table (one row per open exception, owner, age); mail column, requisition and count checks removed.",
    "MIS Exceptions (tblException)", "Open exceptions (<5,000)", "Out: exceptions by owner", "X_ extracts, Refresh Log", "MIS last refreshed 06-Oct-2026 while Stores holds later PO lines - stale numbers today.")
add(W, "4.11 Scrap Dashboard", "Scrap and loss by month, customer, model, part, material, process, batch.", "Simplify",
    "Keep as PivotTable with slicers over a Power Query material-flow table.", "MIS Scrap & Yield", "n/a", "Out: scrap KPIs", "Q_MaterialFlow", "60 spill references today.")
add(W, "M_Contacts", "Mail contacts.", RP, "Scope override 3.", "None", "0", "None", "MailTo names, Checks 'Mail the team'", "Delete with mail names.")
add(W, "Settings", "DataFolder (absolute path), file names, AsOnDate, months.", "Simplify",
    "Typed DataFolder and file names, period parameters; no CELL()/TODAY().", "Settings", "<100", "Parameters", "None", "Absolute D:\\ path breaks when the folder moves.")
add(W, "Lists", "Hidden lists for Control Tower dropdowns and mail.", "Simplify", "Visible lists without spill formulas or mail.", "Lists", "<5,000", "Dropdowns", "X_ extracts", "None")

MIS_D = {
 "D_RMReceipt": ("Import of Stores GRN table (all 38 columns).", "Simplify", "X_GRN with the columns MIS needs only.", "X_GRN", "Stores"),
 "D_RMIssue": ("Import of Stores RM issue table (all 30 columns).", "Simplify", "X_RMMove (values register export).", "X_RMMove", "Stores"),
 "D_ConsReceipt": ("Import of consumable receipts.", "Combine", "X_ConsMove (receipts and issues in one extract).", "X_ConsMove", "Stores"),
 "D_ConsIssue": ("Import of consumable issues.", "Combine", "X_ConsMove.", "X_ConsMove", "Stores"),
 "D_Adjust": ("Import of stock adjustments.", "Simplify", "X_Adjust (authorised adjustments).", "X_Adjust", "Stores"),
 "D_Count": ("Import of physical stock count.", RP, "Scope override 5.", "None", "Stores"),
 "D_Material": ("Import of material master.", "Simplify", "X_Material.", "X_Material", "Stores"),
 "D_Consumable": ("Import of consumable master.", "Simplify", "X_ConsItem.", "X_ConsItem", "Stores"),
 "D_POLine": ("Import of PO lines (2 rows while Stores holds 3 - stale).", "Simplify", "X_POStatus (open PO lines with received/pending).", "X_POStatus", "Stores"),
 "D_Production": ("Import of all production lines (11,406 rows x 38 columns).", "Simplify", "X_Prod with required columns only; monthly summary built in Power Query.", "X_Prod", "Production"),
 "D_MachineLog": ("Import of machine log (empty).", RP, "Machine log deferred to Phase 2.", "None", "Production"),
 "D_Rejection": ("Import of rejections.", "Simplify", "X_RejRework.", "X_RejRework", "Production"),
 "D_WIPCount": ("Import of WIP physical count.", RP, "Scope override 4.", "None", "Production"),
 "D_Machine": ("Import of machine master.", "Simplify", "X_Machine.", "X_Machine", "Production"),
 "D_BOM": ("Import of BOM lines.", "Combine", "X_BOM (lines with edition).", "X_BOM", "Production"),
 "D_BOMModel": ("Import of BOM editions.", "Combine", "X_BOM.", "X_BOM", "Production"),
 "D_ProductP": ("Import of Production product master.", "Combine", "X_Product (single master).", "X_Product", "Production"),
 "D_DCHead": ("Import of DC headers.", "Simplify", "X_DCHead.", "X_DCHead", "Dispatch"),
 "D_DCLine": ("Import of DC lines.", "Simplify", "X_DCLine.", "X_DCLine", "Dispatch"),
 "D_JWCHead": ("Import of JWC headers.", "Simplify", "X_JWCHead.", "X_JWCHead", "Dispatch"),
 "D_JWCLine": ("Import of JWC lines (3,508 x 34).", "Simplify", "X_JWCLine.", "X_JWCLine", "Dispatch"),
 "D_Requirement": ("Import of customer requirements.", "Simplify", "X_Requirement.", "X_Requirement", "Dispatch"),
 "D_FGCount": ("Import of FG physical count.", RP, "Scope override 5; replaced by X_FGBal (calculated in Dispatch).", "X_FGBal", "Dispatch"),
 "D_CoaterReturn": ("Import of coater returns (empty).", "Combine", "X_CoaterMove.", "X_CoaterMove", "Dispatch"),
 "D_CoaterOpening": ("Import of coater opening (empty).", "Combine", "X_CoaterMove (type OPENING).", "X_CoaterMove", "Dispatch"),
 "D_Customer": ("Import of customer master.", "Simplify", "X_Customer.", "X_Customer", "Dispatch"),
 "D_Vendor": ("Import of vendor master.", "Simplify", "X_Vendor.", "X_Vendor", "Dispatch"),
 "D_ProductD": ("Import of Dispatch product master.", "Combine", "X_Product (single master) + X_CustPrice.", "X_Product", "Dispatch"),
 "D_Requisition": ("Import of requisitions (empty).", RP, "Scope override 1.", "None", "Stores"),
 "D_Flow": ("Power Query material-flow table built from production, rejection and stores scrap (11,474 rows).", "Simplify",
            "Keep as Q_MaterialFlow (staging query feeding Scrap & Yield); load to the Data Model or a summary only.", "Q_MaterialFlow", "MIS (derived)"),
 "D_Batch": ("Power Query batch balance recomputed inside MIS (413 rows).", "Combine",
            "Import Production's Batch Balance (X_BatchBal) instead of recomputing - one owner for batch variance.", "X_BatchBal", "Production"),
}
for sh, (purpose, dec, reason, prop, src) in MIS_D.items():
    add("MIS", sh, purpose + " Hidden.", dec, reason, prop, "Matches source" if prop != "None" else "0",
        f"In: {src} workbook (Power Query)" if prop != "None" else "None",
        f"{src} workbook table" , "Refresh on open today; target controlled refresh with Refresh Log row counts." if prop != "None" else "Delete query, connection and sheet together.")
