# Note for Claude - Inventory Register (coil, slitting, CTL) - DRAFT

Status: **draft requirements note, 31 Oct 2026.** Captured from a team discussion. Nothing is built yet.
Use this note as the brief for the next build session. Items marked **?** must be confirmed before building.

Context: this extends the ISKO Excel ERP design in `ISKO_ERP_Sheetwise_Guide.pdf` (four books: Stores, Production,
Dispatch & Challan, MIS). Its rules still apply: one entry pad per register, posted rows saved as values with
permanent IDs, balances calculated from transactions, no purchase requisition, no print or mail features, and no
physical-count sheets. The earlier "remove columns, don't add any" instruction applied to the PO sheets only; this
note asks for new columns.

---

## 1. Raw notes (as said in the meeting)

1. If there is anything said CTL then in customer that customer's RM is going to be cut on CTL
2. Add a column - is it WIP, FG or CTL
3. Down weight ma add column slitting, CTL
4. Coil ma khami
5. Order j etlo j che
6. We want to track the coil - if rejected then it shall show in our master data this many coils were rejected
7. Scrap % in dashboard
8. Circle ma circle dia ketlu nikdu - they want agad
9. Check agad rakhvu dispatch ma ane coil number par
10. Open stock ya toh FG - then we have customer name
11. Juno stock padyo che ey kadhvo che
12. GRN - Dispatch - closing
13. We have to make an inventory register

## 2. What each point means (interpretation)

| # | Interpretation | Confidence | What to build |
|---|---|---|---|
| 1 | Customer master gets a **CTL flag**. If a customer is marked CTL, their raw material (coil) is processed on the CTL (cut-to-length) line. | Medium | `Process Route` column in Customer master: CTL / SLITTING / FABRICATION; it defaults the route on that customer's coils and orders. |
| 2 | Every stock line needs a **stage column**: WIP, FG or CTL. | High | `Stock Stage` column (dropdown: WIP / FG / CTL; add SLIT if slitting stock is held separately **?**). |
| 3 | "Down weight" = the weight section of the register. Add **separate weight columns for slitting and CTL** output. | Medium **?** | `Slitting Output kg` and `CTL Output kg` columns next to input weight; each process's scrap is input minus output. Confirm whether "down weight" means *weight below* (columns under the weight heading) or *down-graded weight*. |
| 4 | Record **coil defects**. | High | `Coil Defect` (dropdown: rust, edge crack, thickness variation, scratch, wavy, other) + `Defect Remarks` on GRN / coil QC. |
| 5 | "The order is only this much": the order quantity is fixed, so processing and dispatch must not exceed it. | Medium **?** | `Order Qty` on the customer order / coil; hard stop when processed or dispatched quantity exceeds `Order Qty` (tolerance **?**). |
| 6 | **Track rejected coils by coil number**, and show the count of rejected coils in master data (per supplier / customer / grade). | High | Coil-level QC status (ACCEPTED / REJECTED / HOLD) in the coil register; a calculated `Rejected Coils` count in the Supplier (and Customer **?**) master; MIS exception for rejected coils not returned. |
| 7 | Show **Scrap %** on the dashboard. | High | MIS tile: Scrap % = scrap kg / input kg, by month, process (slitting / CTL / cutting), customer and coil. |
| 8 | For **circles (round blanks)**: how many circles of what diameter came out, carried **forward** (agad) to the next stage / dispatch. | Medium | Production columns `Circle Dia mm` and `Circle Qty (nos)`, carried as snapshots onto JWC / DC lines. |
| 9 | Keep the check **forward into dispatch, by coil number**: dispatch must quote the coil number and be checked against what that coil produced. | High | `Coil No` on every production, JWC and DC line; hard stop when dispatched qty > produced qty for that coil. |
| 10 | **Opening stock and FG must carry the customer name** (customer-owned or customer-allocated stock). | High | `Customer ID` column on opening-stock rows and on the FG balance; FG reported by customer. |
| 11 | **Find old stock lying unused** ("juno stock padyo che"). | High | Ageing on the inventory register (0-30 / 31-60 / 61-90 / >90 days) + MIS "old stock" list by coil, customer and stage. |
| 12 | Register flow: **GRN in → Dispatch out → Closing**. | High | The register shows Opening + GRN − Dispatch ± process moves = Closing, per coil and per customer. |
| 13 | Build a **coil-wise inventory register**. | High | See section 3. |

## 3. Draft inventory register (coil-wise)

One row per coil (or per coil part after slitting / CTL). Closing is calculated from transactions; there is no
physical count.

| Column | Type | Source | Note |
|---|---|---|---|
| Coil No | ID | GRN | Permanent; traced to production, JWC and DC lines (point 9) |
| Parent Coil No | ID | Production | Set when slitting splits a coil into child coils **?** |
| GRN No / GRN Date | ID / date | GRN | |
| Supplier ID | ID | GRN | Rejected-coil count rolls up here (point 6) |
| Customer ID | ID | GRN / order | Mandatory for customer-owned coils, opening stock and FG (point 10) |
| Process Route | CTL / SLITTING / OTHER | Customer master default | Point 1 |
| Grade / Thk mm / Width mm | Code / decimal | GRN | |
| Order Qty | kg or nos | Customer order | Point 5 |
| Opening kg | decimal | Cut-over | Once only, with Customer ID |
| GRN kg (in) | decimal | GRN | |
| Slitting Output kg | decimal | Production | Point 3 |
| CTL Output kg | decimal | Production | Point 3 |
| Circle Dia mm / Circle Qty | decimal / nos | Production | Point 8 |
| Scrap kg | decimal | Production | Input − output |
| Scrap % | % | Calculated | Feeds the dashboard (point 7) |
| Dispatched kg / nos | decimal | DC lines by Coil No | Point 12 |
| Closing kg | decimal | Calculated | Opening + GRN − output moved on − scrap − dispatch |
| Stock Stage | WIP / FG / CTL | Calculated or picked **?** | Point 2 |
| Coil QC Status | ACCEPTED / HOLD / REJECTED | Lot QC | Point 6 |
| Coil Defect / Remarks | dropdown / text | Lot QC | Point 4 |
| Age Days / Age Bucket | integer / code | Calculated | Point 11 |

**Dashboard additions (MIS):** Scrap % (by process, customer, month); rejected coils (count and kg, by supplier);
old stock > 90 days (by customer and stage); closing stock by customer and stage (WIP / FG / CTL).

## 4. Questions to answer before building

1. Is "CTL" a **process** (cut-to-length line), a **stock stage**, or both? Point 2 lists it next to WIP and FG.
2. "Down weight": columns under the weight heading, or down-graded / secondary material weight?
3. "Order j etlo j che": is the order a fixed cap with zero tolerance, or is a ± tolerance allowed? In kg or nos?
4. After slitting, does each slit get its own coil number (parent → child), or does it keep the parent number?
5. Are customer-supplied coils (job work for the customer) and ISKO-purchased coils both in this register?
6. Circles: is diameter one value per production line, or can one coil produce several diameters?
7. Which book owns this register: Stores (coil stock) or Production (process output)? Suggested: Stores owns GRN and
   closing; Production posts slitting / CTL / circle output; Dispatch posts DC lines by Coil No.
8. Should rejected coils count against the supplier only, or also against the customer when the coil is customer-owned?
