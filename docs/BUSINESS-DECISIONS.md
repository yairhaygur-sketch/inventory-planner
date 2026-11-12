# BUSINESS-DECISIONS.md

Version: 1.0
Status: Approved
Owner: Yair Gur

---

# Purpose

Inventory Planner is a spare-parts inventory planning and decision-support tool.

The system exists to help identify:

- Immediate shortages
- Shortage prevention opportunities
- Open purchase order follow-up
- Excess inventory
- Dead inventory
- Slow-moving inventory
- Working-capital release opportunities

The system is not intended to forecast sales revenue, profitability, or future commercial income.

---

# Core Principle

Inventory decisions must be based on inventory reality, demand history, lead time, open purchase orders and working capital.

Inventory Planner is designed to answer:

> How much money is tied up in inventory?

It is not designed to answer:

> How much revenue might be generated in the future?

---

# Financial Source of Truth

## Approved Decision

Column BZ ("מחיר FOB") is the only approved financial value for inventory calculations.

BZ represents:

- Net manufacturer cost
- FOB part price
- Actual inventory investment value

All financial calculations must use BZ.

---

# Currency

## Approved Decision

Column BE ("מטבע FOB") is the currency field.

Current usage:

- Display only
- No exchange-rate calculations
- No currency conversions

Future versions may use BE for reporting purposes.

---

# Selling Price

## Approved Decision

Column BX ("מחיר מכירה") is not part of inventory-planning logic.

BX must not influence:

- Inventory recommendations
- Safety stock decisions
- Reorder point calculations
- Excess inventory calculations
- Dead inventory calculations
- Working-capital calculations

BX may remain available for future commercial reporting but is excluded from inventory decision making.

---

# Financial Metrics

## Primary Financial KPI

Working Capital

Definition:

Working Capital = Quantity × BZ

This KPI is used for:

- Excess inventory valuation
- Dead inventory valuation
- Slow-moving inventory valuation
- Inventory value calculations
- Capital release opportunities

---

# Revenue Risk

## Approved Decision

Revenue Risk is removed from the system.

The following concepts are deprecated:

- Revenue at Risk
- Potential Revenue Exposure
- Missing Selling Price Alerts

The planner focuses on inventory value rather than theoretical sales value.

---

# Missing Price Logic

## Approved Decision

The system validates only FOB price availability.

Approved alert:

"X items missing FOB price"

Deprecated alerts:

- Missing selling price
- Missing selling price or cost
- Revenue exposure without price

---

# Future Development Rule

Any new feature must answer the following question:

"Does this information help make a better inventory decision?"

If the answer is no, the feature should not become part of the inventory decision engine.

---

# Protected Business Logic

The following business logic remains unchanged:

- Three-year inventory policy
- Dead inventory classification
- Slow-moving inventory classification
- Safety Stock calculations
- Reorder Point calculations
- Lead Time logic
- Open Purchase Order coverage
- Immediate shortage detection
- Shortage prevention logic
- PD item handling
- Warranty analysis

Changes to these areas require explicit business approval.

## Approved changes (log)

- **2026-10-31 · ROP/SS policy A′, stage 1 (PR #145).** For items on the
  three-year-rate path, when the recent rate (EWMA, 4-month half-life) is
  below 25% of the three-year rate, the Reorder Point is calculated on the
  recent rate. Applied only when it changes the suggested ROP, not when there
  is a sale in the current month, and not when annual and monthly columns
  are inconsistent. Safety Stock, SAP `ssMin`, `rate`, coverage and the SAP
  field mapping are unchanged. Every SAP file row keeps SS ≥ ssMin.
- **2026-10-31 · SAP ssMin is not changed by the tool.** The tool shows one
  official, SAP-applicable recommendation. A reduction potential for stale
  SAP floors may be shown separately, never as a second recommendation.

- **2026-10-31 · Two-tab safety (PR #146).** Item marks and decision
  memory are no longer lost when the tool is open in two browser tabs. No
  change to calculations or exports.

- **2026-11-01 · Base month and age reference (BASE_MONTH, PR א׳).** The
  same file, with the same data and settings, must produce the same
  recommendations on any day it is loaded. Approved:
  - The base month (last full month covered by consumption) is determined in
    this order: a manual choice saved for the same file content; the data
    («צר.השנה» = the current-year months in the window + «צר.החודש», with
    confidence thresholds, and the year verified separately from activity
    dates); an explicit, reliable date above the header; otherwise the month
    is **unverified**.
  - When the data and the header disagree, the data decides, with a warning
    that separates the report's extraction date from the base month.
  - Stock age (days since last sale/entry) is counted from the report's
    reliable state date; without one, from the end of the base month — shown
    as a calculation assumption. Not from the day the tool is opened.
  - Activity after the reference date is real, not an error: its age is 0,
    it is flagged and explained, and it does not make an item dead or slow.
  - Unverified month: results are shown as provisional with a clear warning;
    no run history, no profile history, no implementation tracking and no
    SAP export until the user confirms or chooses a month.
  - Operational checks (ETA, mark age) still use the current date.
  - Measured on the real report, against main loaded in the correct month:
    ROP/SS 0 changes, SAP and slow-stock files 0 changes; ages counted from
    30.9 instead of the load day; 5 items move from «מלאי מת» to «נמכר פעם,
    לא זז מאז». Details: `docs/DECISION-RULES.md` section 5ת3.

- **2026-10-28 · Phase 2 — PD→ZM-C and missing lead time (shared engine, both
  tools).** Approved in the minimal-tool conversation. PR #139 merged into
  #138's branch after #138 itself was merged, so it never reached main; it is
  brought onto main in the minimal-tool PR of 2026-11-01.
  - PD→ZM-C when net consumption over 12 calendar months (11 full + the
    current month, returns included) is ≥ 3 units. 04/14 excluded by code
    and text; a missing monthly column means no recommendation (not zero);
    tools/accessories get the recommendation with a manual-review warning.
  - ZM-A/B/C downgrades keep the H rules, waiting period and safeguards; the
    lowest target is ZM-C. No downgrade to PD.
  - Missing planned LT (`אספ.מתוכנ.` blank or 0): «חסר זמן אספקה — יש לעדכן
    ב-SAP». No SS/ROP recommendation (not zero, not a SAP value); no
    replacement LT. Slow stock is marked as an estimate; tied-up money goes to
    «אין מה לשנות».
  - SAP file «זמן אספקה מתוכנן» = the planned LT only, not planned + GR.
  - `ssMin||0` in the SAP file is **not changed** — open until the SAP import
    semantics of a blank cell are verified.
  - Full-tool Excel: no new columns (headers locked by PR-B); the profile
    screen shows ROP/SS after ZM-C on screen only.
- **2026-11-01 · Minimal tool (4 topics) — structure and ETA.** Each
  subcategory is a shared-engine selector, no new rule or threshold:
  customer orders without coverage (no purchase / partial — demand vs. free
  stock + open PO); consumption change (rising / falling — the full tool's
  trend doors); stock 0 now / expected to run out this month (the full tool's
  shortage list, including dated ETA arriving by month end); data cleanup
  (MRP profile change / SS-ROP update / missing LT). The ETA report loads in
  the minimal tool through the full tool's own loader.

- **2026-11-02 · Approved, not yet implemented.**
  - **PDI / customer order against non-free stock** (implemented 8.11, see below). When non-free stock
    (total − free) covers an open customer order, the order is not shown as
    «without coverage». It gets its own state («non-free stock against the
    order — check: reserved for PDI or blocked») and no automatic purchase
    recommendation. The report cannot tell blocked from reserved stock, so
    the order is not treated as covered. Measured on the real report and
    shown to the owner before implementation; separate PR.
  - **MRP-profile exclusion.** The planner can exclude an item from the
    profile-change candidates permanently, until undone manually (one-off or
    on-demand items). Items marked «campaign», «done» or with an exclusion
    note are not shown as candidates. The excluded items have their own view
    for future review. Changes the profile-screen counter — documented in
    its PR.
  - **Blocked stock** (total > 0, free 0) stays a shortage, as today. No
    extra explanation line in the card.

- **2026-11-02 · Minimal tool, round 2 (approved in the minimal-tool
  conversation).**
  - **Consumption change:** the trend stays on the cleaned series (stock-out
    month → median, spike → twice the median, returns → 0). The raw sum from
    the report is shown next to it, and the item drawer says which month was
    adjusted and why. Display only.
  - **MRP-profile exclusion — implemented in part.** A shared «no profile
    change» neutralization: stored separately from marks
    (`planner_profoff_v1`), included in backup, permanent until undone
    manually. The item stays in every other topic. Button in the item drawer
    of both tools; list with undo in the minimal tool; count on the full
    tool's profile screen. **Still open (full-tool conversation):** items
    marked «campaign» / «done» not shown as candidates, and a dedicated
    review view in the full tool.
  - **PD→ZM-C minimum reorder point:** at least 1 unit on the recommendation
    (simulation). Tools/accessories stay manual review without a number.
  - **«Will run out before the next delivery»** (topic 3, requires ETA):
    days to zero = free ÷ monthly rate × 30, from today, before the nearest
    ETA delivery date — also when that is after month end.

- **2026-11-08 · MRP-profile exclusion — completed (UX6, full tool).**
  - An item marked «campaign» or «done», or with any mark that carries a
    note, is not a profile-change candidate while the mark is valid. An
    expired mark does not exclude. Same effect as the manual «no profile
    change»: the recommendation is cancelled in both tools (`pdUpgrade`,
    `profInfo`). The profile-screen counter drops by the number of such
    items with an active recommendation; with no marks it is unchanged
    (real report, clean browser: 0 changes).
  - The full tool's profile screen lists the excluded items (manual and by
    mark, with the reason) above the table; manual ones have «undo», mark
    ones open the item card. Its Excel export adds a sheet «ללא שינוי
    פרופיל»: part number · description · current profile · recommended
    profile · reason · exclusion date · note. No excluded items, no sheet.
  - «MRP profile change» moved from «data quality» to the planning domain,
    under a separate «periodic control» heading.

- **2026-11-08 · PDI — customer order against non-free stock (implemented).**
  - Non-free stock = total − free. When it covers what the order still lacks
    after free stock and open purchase orders, the item gets its own state:
    «מלאי לא פנוי מול ההזמנה — לבדוק: שמור ל-PDI או חסום». No automatic
    purchase recommendation (action: check in SAP whether the stock is
    reserved or blocked; if blocked, release or order). Not «without
    coverage»: not in «אין רכש כלל», «כיסוי חלקי», the minimal tool's
    «customer orders without coverage», or the expedite report's «no order
    at all». Not treated as covered either: the shortage stays, the item
    stays in «חוסרים ולקוחות», same severity, and in the work center under
    «customer waiting, purchase does not cover». Tables show a short label;
    the full state is in the card, the tooltip and the Excel export.
  - Partial coverage (non-free stock smaller than the gap): stays «without
    coverage», with one info line in the card.
  - Measured on the real report (8.11): 0 items in the new state, 1 with the
    info line; queues, screens, counters and exports otherwise unchanged.
    Synthetic report: 6 items move, «כיסוי חלקי» 134 → 128. In the Excel
    export those items' «uncovered customer order» and its value are blank.

- **2026-11-09 · «ROP short of lead time» on the same rate as the
  recommended ROP (approved in principle 1.11, simulation approved 9.11).**
  For an item where RATE_DECAY applies, the screen's check of the current SAP
  reorder point, its «demand during lead time» and «months without stock»
  columns, and the card line «…months without stock» use the recent rate —
  the same rate the recommended ROP is built on. Every other item: unchanged.
  Recommended ROP/SS, classification and the SAP file: unchanged. Measured on
  the real report: the screen 18 → 11 (7 decay items leave, 0 others),
  «ROP below recommended» 221 unchanged, every computed field identical.

- **2026-11-10 · ROP/SS stage 2 — «reduction potential» in the item card
  (approved 10.11).** For an item where RATE_DECAY applies and the SAP ssMin
  floor holds the safety stock above what the recent demand needs, the card
  shows one information line: the extra units, the target-stock value in the
  item's own currency (no conversion), and what the ROP would be without the
  floor — worded as «not an additional recommendation; the tool does not
  change the floor». No money threshold ($250 not adopted). Card only: not in
  the engine's reasons, not in any export, counter or classification. The
  official recommendation and the SAP file are unchanged. Real report: 9
  items, 88 units, about $2,000 and ₪4,000; business-guard 0 differences.

- **2026-11-11 · «אספקות פת.» (OPEN_DEL, column DB) is a customer delivery,
  not purchase in transit (approved 11.11).** When a customer delivery is
  opened in SAP, its quantity leaves «כמות בהז.פ» and waits in «אספקות פת.»
  until it is issued. Until now the tool added it to the open purchase
  orders — so a delivery *to* a customer counted as coverage, as «purchase in
  transit toward surplus», and against the ETA report. Now it leaves the open
  purchase orders and is added to the customer demand (DA + DB). The item
  card shows one info line with the quantity. Data-signature version 2 → 3,
  so the next upload reads as «migrated», not as a new report.
  - Every consequence follows the existing rules on the corrected numbers:
    an order that looked covered by «purchase» (which was the delivery
    itself) becomes «customer waiting · no purchase»; a status-04 item with
    an open delivery says so; an on-demand accessory with an open delivery
    is no longer «nothing to decide» and is judged by the existing accessory
    rule.
  - Real report (11.11): 89 items change. «purchase in transit toward
    surplus» 89 → 59 (about $123,000 → $31,000). Queues: follow 138 → 137,
    immediate 2 → 3, prevent 33 → 34; agenda 38 → 40. Two items change
    classification: one «supply follow-up» whose «purchase» was the delivery
    itself becomes «shortage prevention», and one never-sold item with an
    open delivery gets the PDI state (non-free stock against the order).
    The expedite report shrinks from 696 to 627 lines (−267 units): 69 items
    were listed for expediting with no real purchase order behind them.
    Every SAP update file is unchanged.
  - Synthetic report (DB on every 60th item, as in the real report's ~1%):
    15 items change; «אין רכש כלל» 127 → 128; «purchase in transit toward
    surplus» 107 → 105; one accessory moves from «on demand» to «surplus».
    The synthetic report also gained real partial purchase orders (every
    30th «burning» item) and real open orders on dead items (every 20th):
    until now those cases existed in it only through «אספקות פת.».

- **2026-11-12 · Safety stock covers supplier delay, and active items get at
  least 95% service (SS_LTVAR, approved 12.11).** The statistical part of
  safety stock was ⌈Z × σ × √LT⌉: demand variation only, assuming the supplier
  always arrives exactly on time. It is now ⌈Z × √(LT·σ² + (rate · σ_LT)²)⌉,
  where σ_LT — the lead-time deviation — is the planner's default until it is
  measured from ETA history: 25% of the lead time for imports (Z001), 10% for
  local purchase (Z004), unknown material type as import. An active item
  (5 units a month or more) gets at least 95% service; a higher SAP level is
  kept; below the threshold the SAP level applies as before. The SAP ssMin
  floor, the low-demand policy floor and the max between them are unchanged,
  so no recommended safety stock goes down. The item card shows the supplier
  risk and the service level in force.
  - Exception (approved 12.11): an item with no consumption in the last 11
    months (rate from 3-year history only) gets no supplier risk — otherwise
    33 such items got a statistical requirement and 9 of them (mostly slow,
    5 in decay) a higher safety stock, against RATE_DECAY.
  - Same logic (approved 12.11): for an item where the tool detected demand
    decay, supplier risk and the active-item threshold use the recent rate,
    as the reorder point does (one item: safety stock would jump 1 → 11).
  - Real report (12.11, 10,159 items): 51 items, +724 units, about $37,000
    and ₪11,200. Items selling 30+ a month: median statistical cushion about
    21 → 64 days of sales.
  - Alternatives simulated: delays only (38 items, about $6,000); delays +
    95% for all (462 items, about $117,000 and ₪86,000 — slow items over-
    cushioned); delays + 95% from 3 a month (90 items). The planner chose 5.

Details, numbers and open items: `docs/PROJECT-STATUS.md`.

---

# Summary

Inventory Planner is a working-capital and inventory-management tool.

Financial calculations are based on:

- BZ (FOB Price)
- Quantity

The system does not use selling price (BX) as part of inventory planning decisions.

BZ is the single financial source of truth.
