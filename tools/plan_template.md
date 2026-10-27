---
title: "ISKO Excel ERP Rebuild Plan - 100,000-Row Resilience Blueprint"
subtitle: "Planning document for management approval - Version 1.0 - 27 October 2026"
lang: en-GB
---

<div class="cover">
<div class="cover-band"></div>
<p class="cover-kicker">ISKO Engineering - Excel ERP</p>
<h1 class="cover-title">Excel ERP Rebuild Plan</h1>
<p class="cover-sub">100,000-Row Resilience Blueprint</p>
<table class="cover-meta">
<tr><td>Prepared for</td><td>Managing Director and department heads (Stores, Production, Dispatch, Finance)</td></tr>
<tr><td>Scope</td><td>Four workbooks: 1 Stores, 2 Production, 3 Dispatch &amp; Challan, 4 MIS (2026 files)</td></tr>
<tr><td>Version / date</td><td>1.0 - 27 October 2026</td></tr>
<tr><td>Status</td><td>Planning blueprint for approval. No production workbook has been rebuilt or overwritten.</td></tr>
<tr><td>Performance status</td><td><b>Not yet benchmark-tested in Excel.</b> Targets in Section 17 are provisional.</td></tr>
</table>
<div class="cover-legend">
<span class="lbl current">Current State</span>
<span class="lbl rec">Recommendation</span>
<span class="lbl removed">Removed from Phase 1</span>
<span class="lbl risk">Risk</span>
<span class="lbl decision">Decision Required</span>
<span class="lbl untested">Not Yet Tested</span>
</div>
</div>

<div class="toc-page">
<h2 class="no-num">Contents</h2>
{{TOC}}
</div>

## 2. Executive summary

::: {.callout .current}
**Current State - observed in the attached files.** The four 2026 workbooks contain **120 sheets** and **430,095 formula cells** (Stores 91,698; Production 217,775; Dispatch & Challan 119,138; MIS 1,484). Three sheets hold **81%** of all formulas: *2.1 Production Entry* (203,113), *3.3 JWC Lines* (88,661) and *1.6 RM Issue & Return* (57,683). The books hold about six months of history (April-September 2026): 11,406 production lines, 3,023 RM movements, 3,508 JWC lines and 626 DC lines.
:::

**Key findings**

1. **The present design cannot reach 100,000 rows.** 53,011 formulas sum their own table on every row and 47,369 use expanding "first row to this row" ranges. That work grows with the *square* of the row count: taking *2.1 Production Entry* from 12,538 to 100,000 rows multiplies rows by 8 but this calculation work by about 64. Today the book is usable because it is small.
2. **Record IDs are not permanent.** Issue, production, JWC-line, DC-line, PO-line and consumable IDs are built from `ROW()`; RM lot IDs and batch IDs from running counts. Sorting or inserting a row renumbers them, and every link that quotes them breaks silently.
3. **History can change when a master changes.** Coater rate and basis, customer sale rate, set weight, part weight per piece and dispatch UOM are live lookups on posted rows. Editing a master re-prices or re-weighs old transactions.
4. **About 61% of formulas (≈264,000) depend on functions added in Excel 2019 or later**, mostly Excel 2021 / Microsoft 365 (`XLOOKUP`, `LET`, `FILTER`, `SORTBY`, `UNIQUE`, `SEQUENCE`), and 109 defined names use spill references. The installed Excel version is not recorded in the files and must be confirmed (Decision D-01).
5. **Numbers already disagree between books.** There are three product masters (Stores 249, Production 269, Dispatch 278 rows). MIS was last refreshed on 06-Oct-2026: it shows 2 PO lines where Stores has 3, and reports 442 RM issues with a lot problem where Stores now shows 72.
6. **Controls exist on paper but are weak in practice.** All 217 stock adjustments (net −97,471 kg) have no reason code and no approver; QC status is blank on all 187 lots; validation is bypassed by pasting (form values `SHET`, `SHEE`; 33 DC lines with unit `0`).
7. **The four scope removals touch every book**: purchase requisition (1 sheet, 5 PO columns, 6 names, list columns, MIS import and 3 checks), print (5 print sheets, 15 print areas, 19 printer-settings parts), mail (Gmail links on 10 sheets, 4 contact masters, mail names, the *Mail Coater* and *Mail the team* columns), and physical counts (3 count sheets, 3 MIS imports, count tiles and count-driven adjustments).

::: {.callout .rec}
**Recommendation - target architecture.** Keep four Excel books and the same business flow, but change *how* each transaction table works:

- **Entry pad → append-only values register → bounded balance.** Users type into a small entry pad (about 200 rows) that holds all lookups and checks. Posting copies valid rows into the register **as values** with a permanent ID. Registers contain no formulas, so 100,000 rows cost little to recalculate.
- **Balances only for open items** (open lots, open batches, open JWC lines, FG by product), rebuilt on posting with Power Query or bounded `SUMIFS`. Every hard stop reads these balances.
- **Excel 2016 function set** (`INDEX`/`MATCH`, `SUMIFS`, `COUNTIFS`), Excel Tables, data validation, PivotTables and Power Query. No volatile functions, no spill formulas, no sheet protection, no freeze panes, light pastel styling.
- **Controlled transfer:** each book publishes values-only extracts; consumers refresh at defined points; every refresh writes a **Refresh Log** row (file, modified time, rows, distinct IDs, control totals, status).
- **MIS is read-only:** director KPIs from small Power Query summaries, an exception table with owners, and parameter-driven trace queries.
- **FY-wise files** keep each live register inside the tested volume; open balances carry forward.
:::

**Decisions required before build** (full list in Section 20): (1) Excel version on every PC and whether a signed "Post" macro is allowed (D-01, D-02); (2) how the statutory delivery challan and job-work challan documents will be produced once print is removed (D-03); (3) the stage/route and FG rules that make WIP and FG transaction-calculated, and the approved cut-over opening balances (D-05, D-06, D-08).

::: {.callout .decision}
**Do not build until agreed:** the posting mechanism (depends on D-01/D-02), the WIP and FG balances (D-05/D-06), the opening-balance load (D-08) and any document-number sequence for DC/JWC (D-03). Masters, IDs and the Stores pilot can start once D-01 and D-02 are taken.
:::

::: {.callout .untested}
**Not Yet Tested.** No Excel benchmark has been run. The 100,000-row figure is a design target to be proven with the test plan in Section 17 on the company's own PCs.
:::

## 3. What the current workbooks tell us

All figures in this section were measured from the attached files by reading the workbook XML (every cell, formula, table, defined name, validation and Power Query definition). They describe the files as received.

### 3.1 Inventory

{{TBL_INVENTORY}}

### 3.2 Where the formulas are

{{TBL_HOTSPOTS}}

### 3.3 Formula patterns that limit scale

{{TBL_PATTERNS}}

*Self-table aggregate* = a row formula that runs `SUMIFS`/`COUNTIFS` over its own table. *Expanding range* = `INDEX(col,1):[this row]`. *Position-dependent* = formula uses `ROW()` (IDs and running ranges). *Reads imports* = formula reads a Power Query table that refreshes on open and every 5 minutes.

### 3.4 Growth and calculation load

{{TBL_GROWTH}}

::: {.callout .current}
**Calculation-load estimate (analytical, not a benchmark).** For a row formula that aggregates its own table, a full recalculation evaluates roughly *rows²* comparisons per aggregate. Moving *2.1* from 12,538 to 100,000 rows multiplies that by (100,000 / 12,538)² ≈ **64**; *1.6* from 3,393 to 100,000 rows by ≈ **869**; *3.3* from 4,030 by ≈ **616**. Whatever the time is today, the same design at 100,000 rows is far slower. This is why the plan removes per-row aggregates instead of simply extending formulas.
:::

### 3.5 Compatibility, duplication and design issues observed

- **Excel version:** files were last saved by Excel 16.x and contain calculation features for `LET` and dynamic arrays. About 264,000 formula cells use post-2016 functions; 109 names use `ANCHORARRAY` spill references. These books will show `#NAME?` in Excel 2016/2019.
- **Power Query in use:** Production reads Stores (3 queries), Dispatch reads Production (2), MIS reads all three (31 query tables). Production and Dispatch queries refresh on open and every 5 minutes; MIS refreshes on open. The source folder is an absolute path (`D:\Claude code folder\ISKO ERP\...`) held in *Settings*.
- **Duplicated data:** production history is stored in Production and again in MIS (296,551 cells); RM movements in Stores, Production (52,091 cells) and MIS (79,943 cells); JWC lines in Dispatch and MIS (102,388 cells). MIS imports pre-filled blank rows (3,393 rows for 3,023 issues).
- **Duplicated logic:** batch balance is computed in Production (*2.8*) and again inside MIS (*qBatch*); product, people and contact lists exist in several books.
- **Sheet protection and freeze panes:** 84 of 120 sheets are protected and 44 have freeze panes; all four books lock workbook structure. The target removes both, as instructed, and replaces protection with values-only registers and visible checks.
- **No VBA:** all four files are `.xlsx`; no macro project exists. Print and mail features are formula hyperlinks (Gmail compose links built with `ENCODEURL`) and print areas.
- **Migrated history:** most history was loaded from older registers (Source text such as "Part Production sheet"); 3 production rows are balancing lines ("Yield sheet says ... not itemised"), and MIS reports 190,137 kg of production not itemised. Shift and QC status are blank on all production history.

### 3.6 Measured facts versus assumptions

| Statement | Basis |
|---|---|
| Sheet, formula, table, name, validation and query counts | **Observed in the attached files** |
| Data-row counts and monthly volumes (Apr-Sep 2026) | **Observed** (non-blank date cells in each table) |
| Growth projections (rows per year, years to 100,000) | **Assumption requiring confirmation** - straight-line from six months |
| Calculation-load multipliers | **Analytical estimate** - not measured in Excel |
| Installed Excel version and bitness | **Unknown** - go-live check (D-01) |
| File size of a 100,000-row values register (14.8 MB) | **Measured** with a synthetic file written by Python/openpyxl, not by Excel |
| Open/save/recalculation times | **Not Yet Tested** |

## 4. Design principles and scope

### 4.1 Phase 1 scope

Phase 1 covers **Purchase Order → RM receipt/GRN → RM lot stock → RM issue/return → Cutting → Bending → Fabrication → Rejection/rework → transaction-calculated WIP → Powder coating/job work → Finished goods → Dispatch/challan → Management MIS**, across the four existing books.

::: {.callout .removed}
**Removed from Phase 1 (explicit exclusions)**

- **Purchase requisition** - no sheet, table, field, ID, approval queue, KPI or link. A PO is created on its own; purchase need is shown by the MIS RM Requirement report and re-order flags. Removed: *1.1 Requisition*; PO columns *Req No, Req Item, Req Line, Req kg, Req left kg*; names *ReqFrom, TolReqPO, RqItem, RqLen, RqWid, RqLine, ReqFirst, ReqNoList, ReqOpenList, ReqItemByReq, ReqItmId, ReqItmKey*; *Lists* columns AL:AR; MIS *D_Requisition* and query; checks "Requisitions waiting", "PO lines without a requisition", "PO lines ordering more than their requisition"; the requisition step in *4.7 DC Trace* and the START pages.
- **All print functionality** - *1.3 Print PO, 1.6b Print RM Issue Slip, 1.8 Print Issue Slip, 3.4 Print JWC, 3.11 Print DC*; all 15 print areas and 6 print-title names; 19 printer-settings parts; "Copy page" links; print-only settings (company address/email/GST print fields). No replacement templates.
- **All mailing / email-sending functionality** - "Gmail to" and "Open Gmail" links on 1.3, 1.8, 1.12, 1.13, 2.5, 3.4, 3.11, 4.1, 4.2, 4.8; *M_Contacts* in all four books; *MailTo/MailPick/POMailTo/DCMailTo/JWCMailTo* names; mail list columns; *Mail Coater* (3.6) and *Mail the team* (4.10). Supplier, customer and vendor email fields stay only as optional plain reference text.
- **All physical-count and count-reconciliation features** - *1.10 Stock Count, 2.4 WIP Count, 3.8 FG Count*; MIS *D_Count, D_WIPCount, D_FGCount*; "latest physical count" (4.3), "FG counted" (4.1) and "Counted (kg)" (4.8); count-difference adjustment reasons. RM, consumable, WIP and FG balances are calculated from valid transactions only. A necessary correction uses the **authorised stock-adjustment movement** (reason, reference, requester, independent approver, value).
- **Machine log and utilisation** (recommended for Phase 2, Decision D-09).
:::

### 4.2 Design principles

1. **Enter once, at source.** Each record has one owning book; others import values.
2. **Registers are values.** A posted row never recalculates; corrections are reversal rows.
3. **Permanent IDs** written at posting: prefix + financial year + sequence.
4. **Snapshot what history needs** (rates, weights, UOM, names on documents); look up only what may follow the current master.
5. **Bounded calculation:** formulas in pads and open-item balances only; no whole-column ranges, no volatile functions, no spill formulas.
6. **Controls at entry, exceptions in MIS.** Hard stops in the pad; MIS lists what slipped through and who owns it.
7. **Calm, readable sheets:** light pastel palette, one header row, no merged cells in tables, no freeze panes, no protected sheets.
8. **Version-safe:** Excel 2016 functions unless D-01 approves more.

### 4.3 Assumptions

- Windows desktop Excel, one editor per book at a time, books in one shared folder on a local network or PC.
- Power Query (Get & Transform) is available in the installed Excel - it is already used by three of the books, which is evidence but not proof for every PC.
- Typical office PC: 4-core CPU, 8-16 GB RAM, SSD, 64-bit Excel recommended (go-live check).

### 4.4 Non-negotiable transaction controls

Unique permanent IDs; no negative stock through ordinary issue; QC-accepted lots only; return only against a batch; cutting cannot exceed batch kg; dispatch cannot exceed FG or JWC pending; customer UOM mandatory; no posted-row editing (reversal only); period lock; every adjustment reasoned and independently approved.

## 5. Four-workbook target architecture

![Four-workbook target architecture: ownership, extracts, refresh points and IDs](assets/architecture.svg){.fig-wide}

**How a transaction flows.** A user fills the entry pad in the owning book. The pad shows lookups, available balances and a *Ready* flag. *Post* copies ready rows into the register as values with a permanent ID, then the book's balances are refreshed. At defined points the downstream book refreshes its import (X_ tables) and MIS refreshes everything, writing the Refresh Log.

![Entry pad, values register and bounded balance](assets/pattern.svg){.fig-mid}

| Layer | What it holds | Formulas? | Size |
|---|---|---|---|
| Entry pad | Today's unposted rows, lookups, checks, Ready flag | Yes - all entry logic | ~200 rows |
| Register | Posted transactions, values only, append-only | **None** | up to 100,000 rows per FY file |
| Balance | Open lots / batches / JWC lines / FG / WIP | Power Query group-by or bounded SUMIFS | open items only |
| Extract (X_) | Values from another book + lineage columns | None (Power Query output) | as source, needed columns only |
| Report | KPIs, exceptions, trace | PivotTables / small summaries | hundreds of rows |

## 6. End-to-end process map

![End-to-end process map with exception and reversal paths](assets/process_map.svg){.fig-wide}

**Exception and reversal pathways**

- **GRN:** excess over PO → *Excess decision* (accept / return) in the QC log before the lot can be issued; QC HOLD/REJECTED → lot not issue-eligible; supplier return → *RETURN-TO-SUPPLIER* movement.
- **RM movement:** wrong issue → reversal row quoting the Move ID, then a correct row; usable material back → *RETURN-USABLE* / *RETURN-OFFCUT* against the batch; damaged → *RETURN-DAMAGED* (separate bucket, not issue-eligible); stores scrap → *STORES-SCRAP* against the batch.
- **Production:** over-consumption is blocked; wrong line → reversal; rejection → REJECT case → REWORK-OK (back to WIP) or REWORK-SCRAP; batch closes only through a Batch Close row with reason/approver above tolerance.
- **Coater:** defects back → DEFECT-BACK with decision (re-send, scrap, debit coater); direct dispatch from coater → DC line quoting the JWC line.
- **Dispatch:** cancelled challan → DC status line CANCELLED plus reversal lines; FG cannot go negative.
- **Corrections to stock** outside these paths → authorised stock-adjustment movement only.

Purchase requisition, print/mail steps and physical-count/reconciliation flows are not part of the map.

## 7. Workbook responsibility matrix

| Workbook | Purpose | Owns (system of record) | Imports (values) | Must not do |
|---|---|---|---|---|
| **1 Stores** | Purchase, receipt, QC, RM and consumable stock | PO header/line, GRN/lot, Lot QC, RM movements, consumable receipts/issues, adjustments; masters: Material, Grade, Supplier, Consumable, Category, Where-used | X_Product, X_Machine, X_People (from Production) | Hold production, dispatch or KPI history |
| **2 Production** | Make parts, consume batches, control yield | Production lines, rejection/rework events, batch close; Batch Balance, WIP Balance; masters: Product, Part, BOM, Stage/Route, Machine, People, Reasons | X_Batch, X_Material (from Stores); X_Customer (from Dispatch) | Re-key stores issues; compute dispatch |
| **3 Dispatch & Challan** | Requirements, job work, coaters, FG, delivery | Requirements, JWC header/lines, coater movements, DC header/lines; Open JWC Balance, FG Balance; masters: Customer, Customer price & UOM, Vendor, Coater rate | X_ProdOutput, X_Product, X_People (from Production) | Hold production history |
| **4 MIS** | Director reporting and cross-book checks | KPI summaries, exception table, Refresh Log, trace queries | X_ extracts from all three books | Accept transaction entry |

### 7.1 Source-of-truth matrix (one row per data entity)

{{TBL_ENTITY}}

<div class="landscape">

## 8. Sheet-by-sheet rebuild decision register

Every one of the 120 current sheets is listed with its measured size and the recommended action. *Rows* = non-blank records in the sheet's table (where a table exists); *Formulas* = formula cells; *Used range* = the sheet's dimension. Decisions: **Keep / Simplify / Combine / Removed from Phase 1 / Convert to report**. The full register with inputs/outputs is also in the decision-register workbook.

{{TBL_SHEETS}}

</div>

<div class="landscape">

## 9. Column-by-column design blueprint

**Legend.** Class: **IN** user input · **DD** dropdown / master selection · **SN** transaction-time snapshot · **CE** calculated at entry (posted as value) · **LK** lookup / display only (pad) · **CF** exception / control flag · **IM** imported / read-only · **RP** report-only. Mandatory: **M** mandatory · **O** optional · **C** conditional. Validation: **H** hard stop · **W** warning (post with authorised explanation).

**Standard audit columns** appended to every register (marked "+ audit" below):

{{TBL_AUDIT}}

{{TBL_BLUEPRINTS}}

</div>

## 10. Master data and unique-ID standards

### 10.1 Master ownership

| Master | Owner book | Key | Rules |
|---|---|---|---|
| Material, Grade | Stores | Material ID (e.g. GPSP-1.6), Grade code (GP) | Never re-key; Active flag; standard size and sheet weight |
| Supplier | Stores | Supplier ID (SUP-0001) | Tally ledger name kept as reference; email optional text |
| Consumable item, Category | Stores | Item ID (ABR-001) | Category prefix + sequence |
| Product / model | Production | Product ID (short code) | One master for all books (replaces 3 copies); Customer ID; coating route |
| Part | Production | Part ID (product + part code) | Drawing no, revision, material, thickness, Wt/pc, route |
| BOM model / edition / lines | Production | Model + Edition + SR No | New edition = new rows; effective-from dates |
| Stage & route | Production | Stage code; Route ID | Sequence per part family; FG point (D-05) |
| Machine | Production | Machine ID (MC-01) | Process = stage; capacity Phase 2 |
| People | Production | Person ID (PER-001) | Role; Active; used by all books |
| Customer, Customer price & UOM | Dispatch | Customer ID (CUS-001) | Dispatch UOM mandatory; price effective-dated |
| Vendor / coater, Coater rate | Dispatch | Vendor ID (VEN-001) | Rate basis and rate effective-dated |
| Reasons | Owning book | Reason code | Codes, not free text; count reasons deleted |

### 10.2 ID standard

| Record | Format | Example | Generated by |
|---|---|---|---|
| PO / PO line | PO/yy-yy/nnn ; PO No-nn | PO/27-28/014-02 | Settings prefix + next number |
| Lot | Grade code-FY-nnnn | GP-2728-0118 | Grade code + COUNTIF of grade in register + 1 |
| QC event | QC-FY-nnnnnn | QC-2728-000231 | Sequence |
| RM movement / batch | ISS/RET/OFC/DMG/SCR/RTS-FY-nnnnnn | ISS-2728-004512 (= Batch ID) | Sequence shared by the register |
| Consumable receipt / issue | CR-/CI-FY-nnnnnn | CI-2728-000877 | Sequence |
| Adjustment | ADJ-FY-nnnnn | ADJ-2728-00012 | Sequence |
| Production line | PRD-FY-nnnnnn | PRD-2728-018230 | Sequence |
| Rejection / rework | REJ-/RWK-FY-nnnnn | REJ-2728-00341 | Sequence |
| Requirement | REQ-FY-nnnnn + version | REQ-2728-00045 v2 | Sequence |
| JWC / JWC line | JWC/yy-yy/nnnn ; JWC No-nn | JWC/27-28/0412-03 | Settings prefix + next number |
| Coater movement | CRT-FY-nnnnn | CRT-2728-00090 | Sequence |
| DC / DC line | DC/yy-yy/nnnn ; DC No-nn | DC/27-28/0191-01 | Settings prefix + next number (current DCNext = 191) |
| Extract run | RUN-yymmdd-hhmm | RUN-270415-0915 | Power Query |

Rules: IDs are written as values at posting and never recalculated; the FY element prevents collisions after archival; sequences are `COUNTA(register key) + pad row number`, which is safe only because registers are append-only and one person posts at a time; statutory document numbers (DC, JWC) must have no gaps - a cancelled number stays in the register with status CANCELLED.

### 10.3 Naming and revision rules

Names are labels, IDs are keys. A renamed product keeps its ID. A new drawing revision is a new Part row with a new effective date; the old row is set inactive, never overwritten. Master deletions are not allowed; set *Active = No*.

## 11. Formula and performance architecture

### 11.1 What stays a formula

| Location | Formulas | Rule |
|---|---|---|
| Entry pads (≈200 rows) | Lookups (one MATCH per key + INDEX), calculated values, checks, Ready flag | Recalculate on every edit - cheap because small |
| Balance tables | One aggregate per open item (or Power Query) | Rebuilt on posting / refresh |
| Registers | **None** | Values only |
| Masters | Only simple row-local helpers (e.g. sheet weight) | No aggregates over transactions |
| MIS | Small summaries over Power Query outputs, PivotTables | Refresh only |

**Formula budget:** ≤ 40,000 formula cells per book at 100,000 register rows (today: up to 217,775 at 12,538 rows).

### 11.2 Calculated once at posting, not on every edit

IDs; snapshot rates, weights, UOM and names; stock effect; value; available-before evidence; duplicate key. These are pad formulas whose *results* are posted.

### 11.3 Posting mechanism - VBA versus no macros

| | **A. Signed "Post" macro (preferred if D-02 allows)** | **B. Manual paste-values (fallback / no macros)** |
|---|---|---|
| How | Button validates Ready rows, writes values to the register, stamps Posted At/By, clears the pad, refreshes balances | User selects Ready rows, *Copy → Paste Values* into the next register rows, types Posted At/By, clears pad, *Refresh All* |
| Control strength | Hard stops cannot be bypassed; consistent | Depends on procedure; pad control total must equal register delta |
| Security | `.xlsm`, digitally signed, macros enabled only for that signature; one small module (~150 lines) | No macros |
| Backup / audit | Macro writes a posting log row | Poster records the control total in the posting log |
| Compatibility | Excel 2016+ Windows | Any Excel |
| Maintenance | Needs someone able to read VBA | None |

Both use identical pads and registers, so a book can move from B to A later.

### 11.4 Why this scales better

{{TBL_SCALE}}

## 12. Cross-workbook integration and MIS refresh design

### 12.1 Method comparison

| Method | Pros | Cons | Use |
|---|---|---|---|
| Live external-link formulas | Instant | Fragile paths, slow, break on rename; recalc chains | **Not used** |
| Copy / paste-as-values export | Works in any Excel | Manual; duplicate risk | Fallback when Power Query unavailable |
| **Power Query from the saved source book** | Already used; replace-not-append; no formulas; lineage columns | Reads last *saved* version; needs Excel with Get & Transform | **Recommended** |
| VBA-assisted refresh | Can add logging and checks | Needs macros | Optional with D-02 |

### 12.2 Extract catalogue

{{TBL_EXTRACTS}}

Each extract also carries the standard lineage columns (Section 9, *X_ extract*).

### 12.3 Refresh points

| When | Who | What |
|---|---|---|
| Start of shift | Production / Dispatch | Refresh X_Batch, X_ProdOutput (after Stores/Production saved) |
| After each posting | Book owner | Refresh own balances (automatic with macro) |
| End of shift | Each owner | Save, back up, write day control totals |
| Before MIS review (daily 10:00 suggested) | MIS owner | Refresh All in MIS, read Refresh Log, review exceptions |
| Month close | Finance | Lock date, close totals, archive copy |

### 12.4 How a user knows a refresh succeeded

The MIS header and every operational book show a **Refresh status panel**: last run time, each source with *OK / STALE / MISSING FILE / MISMATCH / DUPLICATES*, rows read and control totals. Anything except *OK* turns the panel coral and the Control Tower shows a banner; KPIs from that source are labelled *stale*.

### 12.5 Duplicates, corrections and file problems

- **Duplicates:** extracts *replace* their table on each refresh (never append) and filter on a non-blank primary key; the Refresh Log compares rows with distinct IDs.
- **Corrections:** happen only in the owning book as reversal rows, so the next refresh carries them automatically.
- **File closed / moved / renamed / missing:** Power Query reads the saved file (closed or open); a missing file gives *MISSING FILE* and the previous extract is kept, clearly marked stale.
- **Updated by another person:** only the owner edits a book; a consumer reading during an unsaved edit gets the last saved state and the *Source Modified* time shows it.
- **Reconciliation after refresh:** control totals (rows, kg, Rs) of each extract equal the source register totals written in the source book's *Control Totals* table.
- **Delay tolerance:** KPIs, trends and trace may be up to one shift old. Stock availability, batch kg left, JWC pending and FG available must be validated in the owning book at the moment of entry - never from MIS.

## 13. Inventory and production control design

### 13.1 RM lot balance and availability

Lot available kg = accepted kg (latest QC decision) + approved adjustments + signed stock effect of all movements. Sheets follow the same rule. Balances are rebuilt for open lots only after each posting; the pad subtracts quantities already typed in the pad for the same lot, so two pad rows cannot together exceed stock. A **stale guard** compares the register row count stored in the balance with the live count; a difference blocks entry until refresh.

### 13.2 Issue and return rules

| Movement | Lot effect | Must quote | Hard stops |
|---|---|---|---|
| ISSUE | − kg, − sheets | Machine, person, purpose | Lot eligible; kg ≤ available; sheets ≤ available |
| RETURN-USABLE | + kg, + sheets | Batch | ≤ batch returnable; same lot |
| RETURN-OFFCUT | + kg (sheets 0) | Batch | ≤ batch returnable |
| RETURN-DAMAGED | + kg to DAMAGED bucket (not issue-eligible) | Batch | ≤ batch returnable |
| STORES-SCRAP | 0 to lot; batch consumption | Batch | ≤ batch returnable |
| RETURN-TO-SUPPLIER | − kg | Lot, supplier document | ≤ available |

### 13.3 Batch balance, closure, yield and scrap

Batch kg left = issued − returned − cut output − cut scrap − stores scrap. A cutting line cannot exceed kg left plus tolerance (Settings, today 2% with a 1 kg minimum). Closing a batch posts a *Batch Close* row that freezes variance kg, variance % and variance value; above tolerance it needs a reason and an approver. Yield = cut output / net batch kg on closed batches.

### 13.4 Rework

Rejections are cases: REJECT (from a source Prod ID) → REWORK-OK (quantity returns to WIP at the detection stage) or REWORK-SCRAP. The original production line is never edited; QC-OK quantity = produced − rejected + reworked OK.

### 13.5 Transaction-calculated WIP (no counts)

For each Product + Part + Stage on the part's route: **WIP = completed at this stage − completed at the next route stage − sent on JWC from this stage − transferred to FG − rejected at this stage + rework OK returned to this stage.** Opening WIP at cut-over is a one-time approved migration entry derived from closing book figures (D-08), not a count. Negative WIP is an exception (route or posting error), never a reason to post a count. The route master is new and needs Production sign-off (D-05).

## 14. JWC, subcontractor and dispatch controls

![Illustrative Excel Layout - Mock-up: Open JWC / coater balance (sample data)](assets/mock_jwc.png){.shot}

- **Pending at coater** per JWC line = sent − BACK-OK − DEFECT-BACK − scrapped at coater − direct dispatch (DC lines quoting the JWC line). Coater opening at cut-over is a one-time OPENING movement confirmed from the coater's statement (not a count). Returns and direct dispatch above pending are blocked.
- **Return / debit logic:** every defect carries *whose fault* and a *decision* (rework at coater, re-send, scrap, debit coater); a debit needs an amount; a case closes when a later movement quotes it.
- **Cost:** coater rate and basis come from an effective-dated rate table at the JWC date and are frozen on the line (today they re-price history).
- **Customer UOM:** the dispatch UOM comes from a Customer-Product UOM table, else the customer default; QTY / NOS / SET use quantity, KG uses weight, TON uses weight ÷ 1,000. Blank UOM blocks posting (8 of 15 customers are blank today, D-18).
- **FG availability:** FG = opening + coater BACK-OK + uncoated final-stage completion − dispatch from factory − final-QC rejection ± authorised FG adjustment. A DC line from the factory cannot exceed FG available; from a coater it cannot exceed JWC pending.
- **Traceability:** DC line → JWC line → batch → lot → GRN → PO, using IDs on every row.

::: {.callout .risk}
**Risk R-14 - statutory documents.** Removing print removes the only place the ERP produces the delivery challan and job-work challan documents. Goods movements still need those documents. Management must decide how they are produced (Decision D-03) before go-live; the ERP will keep the challan number, lines and status as the record.
:::

## 15. Management dashboards and exception reporting

![Illustrative Excel Layout - Mock-up: MIS Control Tower (sample data)](assets/mock_mis.png){.shot}

PivotTables and small summary tables built by Power Query are preferred over formula dashboards: they work in Excel 2016, refresh in one step and do not scan registers on every edit. Charts read only the summaries (hundreds of rows). Filters: period, customer, material, machine, product (slicers).

### 15.1 KPI definitions

{{TBL_KPI}}

### 15.2 Exception catalogue

{{TBL_CHECKS}}

Exceptions are one row per open problem with an owner and age, rebuilt on each MIS refresh. Hard stops in the books should keep BLOCKER exceptions at zero; any BLOCKER appearing means a control was bypassed and is reviewed the same day.

## 16. Data migration, archival and go-live

### 16.1 History and archival policy

| Option | For | Against |
|---|---|---|
| Keep 100,000+ rows in one live register | Simple search | Slower open/save, larger backups, more risk per file |
| **FY-wise files (recommended)** | Each register holds one year (Production ≈ 23k rows/yr) - well inside the 100,000 target; clean statutory numbering | Needs carry-forward of open balances |
| Half-year files | Fallback if a PC fails the test | More files |

**Policy:** at FY close, archive the four books read-only (file attribute and a separate folder; a checksum is stored in the Close Log). Carry forward as OPENING rows: open lots (kg, sheets, rate), open batches, WIP by product/part/stage, at-coater balances by JWC line, FG by product, open POs, open requirements. IDs include the FY so they never collide. MIS reads current and archived years through a folder query (*Folder.Files*), so history stays searchable and trace queries cross years. Nothing is deleted because it is old.

### 16.2 Migration and go-live steps

1. **Freeze and back up** the current books (read-only archive with checksums).
2. **Clean masters:** one product master (merge 249/269/278 rows), Person IDs, Customer UOM (8 blanks), coater rates with effective dates, route master.
3. **Map** current columns to target columns (Appendix B); history either migrated as values with *Entry Source = MIGRATION* or left in the archived books (D-17).
4. **Opening balances at cut-over** (D-08): RM lots, consumables, WIP, at coater, FG, open POs, open requirements - taken from the closing figures of the current books, reviewed by Finance and posted as OPENING movements. This is a one-time migration, not a count module.
5. **Validate:** totals by lot, batch, coater and product match the approved closing figures; zero BLOCKER exceptions.
6. **Parallel run** one month (March 2027 suggested) with daily comparison of key totals.
7. **Cut-over** on 1 April 2027 (start of FY 2027-28) so numbering, archival and opening balances align.
8. **Hypercare** for two months with daily exception review.

## 17. Stress-test and acceptance plan

::: {.callout .untested}
**Planned Validation - Not Yet Tested.** No Excel timing has been measured. The only measurement made is file size: a synthetic register of 100,000 rows × 32 columns written as values by Python (openpyxl) is **14.8 MB** (10k: 1.5 MB; 25k: 3.7 MB; 50k: 7.4 MB). Today's Production book is 7.2 MB at 12,538 rows; growing the current design linearly would give roughly 58 MB at 100,000 rows (estimate).
:::

**Test environment to record:** PC model, CPU, RAM, disk type, Windows version, Excel version/build and bitness, calculation mode (automatic), add-ins, file location (local / network), file size, row counts per register, elapsed time (median of three), memory/CPU from Task Manager, errors, pass/fail. Run on the **slowest** PC that will be used.

**Volumes:** each major register at 10,000 / 25,000 / 50,000 / 100,000 rows of synthetic or masked data, and a combined system of 400,000+ register rows across the four books with masters and extracts.

### 17.1 Provisional acceptance thresholds

{{TBL_THRESH}}

These are **targets requiring validation on the company's computers.** After the pilot, the MIS owner records actual medians; Management keeps or tightens thresholds that pass with margin, and for any that fail chooses: shorter archive period, fewer extract columns, Data Model instead of sheet tables, or a Phase 2 database.

### 17.2 Test matrix

{{TBL_TESTS}}

Sign-off: MIS owner (execution), department head of each book (functional), Finance (totals), MD (go-live).

## 18. Risk register

{{TBL_RISKS}}

## 19. Implementation roadmap

{{TBL_ROADMAP}}

Each gate passes only when its tests pass and the department head signs. A failed stress test stops the next pilot, not the whole programme.

## 20. Management decisions required

{{TBL_DECISIONS}}

**Top three:** D-01/D-02 (platform: Excel version and macro policy), D-03 (statutory challan documents after print removal), D-05/D-06/D-08 (WIP/FG rules and cut-over opening balances).

### 20.1 Business requirements - decisions

{{TBL_REQS}}

<div class="landscape">

## 21. Appendices

### Appendix A - Full sheet inventory (measured)

{{TBL_APPX_A}}

### Appendix B - Current column inventory and actions

Every column of every Excel Table in the Stores, Production and Dispatch books, with the recommended action. MIS import tables (D_*) mirror these source tables column for column and follow the same actions; they are listed in full in the decision-register workbook.

{{TBL_APPX_B}}

</div>

### Appendix C - Formula and compatibility notes

| Current function / feature | Minimum Excel | Target replacement |
|---|---|---|
| XLOOKUP, XMATCH | 2021 / 365 | INDEX + MATCH (one MATCH helper per key) |
| LET | 2021 / 365 | Helper columns in the pad |
| FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, spill references (#, ANCHORARRAY) | 2021 / 365 | Power Query outputs, PivotTables, validation lists from Tables |
| TEXTJOIN | 2019 | Separate flag columns; one *Issues* text built with `&` |
| MAXIFS / MINIFS | 2019 | Power Query group-by, or sorted lookup |
| ENCODEURL + HYPERLINK (Gmail compose) | 2013 | **Removed** (mail) |
| TODAY, NOW, CELL, OFFSET, INDIRECT | any (volatile) | Typed As-on date; INDEX ranges; typed folder path |
| ROW()-based IDs, running COUNTIF IDs | any | IDs posted as values |
| Power Query (Get & Transform) | 2016 (built in) | Kept, controlled refresh |
| ISFORMULA (integrity check) | 2013 | Used in pad headers |

### Appendix D - Terms

| Term | Meaning |
|---|---|
| Entry pad | Small entry area (≈200 rows) with all formulas and checks; rows are posted, then cleared |
| Register | Append-only Excel Table of posted transactions, values only |
| Balance | Calculated position for open items (lots, batches, JWC lines, FG, WIP) |
| Snapshot | Value copied at posting so later master edits do not change history |
| Extract (X_) | Values-only import of another book's register or balance |
| Refresh Log | One row per extract per refresh with status and control totals |
| Batch | One RM ISSUE movement; its Move ID is the Batch ID |
| Hard stop / warning | Entry blocked / entry allowed with an authorised explanation |
| Authorised stock adjustment | Approved correction movement with reason and reference; not a count |
| FY | Financial year (April-March) |

### Appendix E - Test checklist

{{TBL_TESTCHECK}}
