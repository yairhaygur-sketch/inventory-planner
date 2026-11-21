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
- **2026-10-31 · SAP ssMin is not changed by the tool.** *(Superseded on
  2026-11-19 by SS_STAT for items with consumption — see below.)* The tool shows one
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

- **2026-11-15 · Correction: «אספקות פת.» are not customer demand either
  (OPEN_DEL, approved 15.11).** The 11.11 change added DB to customer demand.
  That was a double count: an open delivery's goods are already assigned to
  it from stock and are no longer in free stock — in the report uploaded
  on 15.11, for
  779 of 786 items with DB, non-free stock (total − free) equals DB exactly,
  and covers it in 2 more. The customer is already served from the shelf.
  Now DB is neither demand nor purchase: customer demand is «כמות בהז.פ»
  only, open purchase is «הז. רכש» only. The card keeps one info line
  (quantity; «already assigned from stock, not open demand and not purchase
  in transit»). The 11.11 benefit stays: DB is not purchase in transit, so
  «purchase in transit toward surplus» and the expedite report stay clean.
  Data-signature version 3 → 4.
  - Real report (uploaded 15.11, 10,160 items): 786 items change, all with DB. 664
    phantom «non-free stock against the order — check» items disappear (662
    back to «no action»); «today's decisions» 708 → 44; immediate 665 → 1;
    every SAP update file is unchanged.
  - Synthetic report: 15 items change, all with DB; the on-demand accessory
    returns to «on demand».

- **2026-11-15 · The 4-topics view: critical and high findings of the screen
  review (approved 15.11).** Display only — no new calculation, threshold or
  classification; every number is an engine field.
  - The «action» column shows the action of the topic: in «data quality» what
    to update in SAP (profile, ROP/SS, missing lead time or service level), in
    «consumption deviation» the advisory reorder point. Before, it showed the
    engine's first action even from another topic («open a purchase order» in
    data quality). The CSV keeps the engine's action too, as in the full tool.
  - Default sort by urgency: customer orders — no purchase first, then order
    value (in the item's currency, in currency blocks, never ₪ against $);
    zero stock — empty shelf first, then days to zero; consumption — the
    effect on the reorder point; data — broken data first.
  - Customer orders gain «order value». Zero stock shows «to order» (the
    quantity in the action itself) instead of «not supplied this month»
    (which stays in the CSV and the card).
  - One line per item; one toolbar (sub-category chips, filters folded behind
    «filter», active filters as chips); one explanation line, the full text
    behind ⓘ; a «no ETA report» banner, and the ETA columns hidden without one;
    keys C (copy) and D (handled, then the next item).
- **2026-11-15 · The 4-topics view: medium and low findings (approved 15.11).**
  Display only, as above.
  - An item that appears in another topic says so («גם בטיוב נתונים»), and
    «handled» says which other topics it leaves (it is one mark per item).
  - Columns that decide nothing leave the table (kept in the card and CSV):
    «planner» when there is one planner or one is selected, «MRP type» in
    customer orders, «lead time» in zero stock; the «status» column only when
    the list is mixed (not after choosing a sub-category).
  - Each tab shows its urgent part in red (no purchase · at zero · broken).
  - Consumption deviation: one number per cell (the cleaned sum), the report's
    sum behind a dot (and as text for screen readers); the effect on the
    reorder point («proposed ← by the trend») before the percentage.
  - One name per thing across the four topics («status», «next delivery
    (ETA)», «(SAP ← proposed)»). A zero change in value shows «—».
  - Card: when moving PD → ZM-C lowers safety stock below today's SAP value,
    a line says why (re-derived from consumption; the floor is minimum SS).

- **2026-11-18 · Sharp drop within the six-month window (SHARP_DROP, option B)
  and no «raise» inside «demand decline» (DROP_NO_RAISE) — approved by the
  planner 18.11 (PR #177).**
  - **The problem:** on the six-month path (more than 3 units in 11 months),
    the rate is the average of the last six cleaned months. When demand
    collapses inside that window, the average and σ still carry the strong
    months, so the statistical SS and the ROP rise. Test item PN-07 (real
    report, 17.11): 10 units in the last 3 months vs 346 before; suggested
    ROP 318 vs 60 in SAP, while the card said «demand decline».
  - **The rule:** when the last 3 cleaned months sum to at most 25% of the 3
    before them (RATE_DECAY), the 3 before sum to at least 6 units
    (SHARP_MIN_PRE), and there is no sale in the current month: the ROP rate
    is the last-3-months rate, σ for the statistical SS is the sample
    standard deviation of those 3 months, and supplier risk uses the same
    rate. The «ROP shorter than lead time» check uses the same rate.
  - **Unchanged:** `rate`, coverage, trend, slow stock, classification
    inputs; SAP `ssMin` and the low-demand policy floor (SS ≥ ssMin in every
    SAP row); items on the three-year path (RATE_DECAY applies there).
    The new fields exist only on affected items.
  - **DROP_NO_RAISE:** in «demand decline», when the calculation is above
    the SAP ROP, the action is «keep the SAP reorder point — do not raise
    during a demand decline», instead of «consider raising».
  - **Measured against main (real reports, outside the repo):** report B
    (about 40,000 items) — 24 items change, all ROP down, none up; items with a
    suggested ROP above SAP among them 19 → 3; in the SAP update file 10 → 2
    such items, +57 → +2 units above SAP; 8 items leave the SAP file (their
    suggestion now equals or is close to SAP); «consider raising» inside
    «demand decline» 48 → 0. PN-07: ROP 318 → 47, SS 40 (the SAP floor).
    Report A (about 10,000 items) — 15 items change, all ROP down; «raise» inside
    «demand decline» 6 → 0. Demo report: no change (business guard clean).

- **2026-11-18 · The 4-topics view: one supply action in «zero stock /
  running out», and FOB price and currency columns (approved 18.11).**
  Display only; the engine, the recommendations and the SAP file are unchanged.
  - In all three parts of «zero stock / running out»: one action per item, in
    words, with no quantity (a number in the action is adopted unchecked).
    Parameter actions (reorder point / safety stock) belong to «data quality».
    Order: no open purchase → «open a purchase order»; customer waiting and an
    ETA → «give the customer a date: dd.mm.yyyy»; purchase without an ETA →
    «ask the supplier for an ETA»; ETA too late (shelf empty, short this month
    after ETA deliveries, or runs out before the next delivery) → «expedite the
    supplier»; otherwise → «check back order with the manufacturer».
  - The «to order» column leaves the table; it stays in the CSV export.
  - «FOB price» (BZ) with its currency (BE), one column in all four topics:
    the currency sign as everywhere in the tool ($ € ₪), or a small code when
    there is no sign (e.g. SEK), and the code on hover. Two columns did not fit
    «data quality» at 1366px. Prices of 100 and above show without decimals; the exact price is in the CSV, where price
    and currency are two fields and a duplicate «currency» header appears once.

- **2026-11-18 · The 4-topics view: the planner arranges the columns
  (approved 18.11).** Display only.
  - Drag a column header to a new place, or open «columns» to show / hide a
    column, move it earlier / later, or return to the default. Columns that
    are hidden by default (for example «to order») can be shown from there.
  - Saved in the browser, per topic. «Part number» is always shown and first.
    The context rules stay: ETA columns only with an ETA report, «planner»
    only with more than one planner, «status» only under «all».
  - The CSV export does not depend on the arrangement.
  - First in the 4-topics view; the full tool can use the same mechanism later.

- **2026-11-18 · The 4-topics view: an arrivals board, deliveries grouped by
  date in the card, and «first consumption» (approved 18.11).** Display only;
  the engine, the recommendations and the SAP file are unchanged.
  - «Arrivals board» (header button): one row per arrival date in the ETA
    report, only for items that are in one of the four topics (after planner,
    search and filters); dates that passed come first («check if received»).
    A date opens its items; a part number opens the card in its first topic.
  - The card groups ETA lines by date: total units, number of lines and the
    delivery numbers.
  - «First consumption», a sub-category of «consumption deviation»: an item
    never consumed in all the history in the report (11 months and 3 years)
    with consumption in the last full month or the current month. All such
    items, no filter. Action, in words: «check whether to set up planning for
    the item». Without the 3-year columns the rule does not apply. Measured:
    report A 133 items, report B 143 (demo 108).

- **2026-11-18 · Weighted rate for lumpy and sporadic items (RATE_W, approved
  18.11).** For items whose demand pattern is lumpy or sporadic, on the
  six-month path and with all three annual columns, the rate used for the
  reorder point and the supplier risk is 70% the engine rate (six cleaned
  months), 20% last year (without its months already in the window) and 10%
  two years ago (annual column ÷ 12). σ, the SAP `ssMin` floor, the policy
  floor, `rate`, coverage, trend and classification are unchanged. Not on
  smooth items (in the backtest it hurt their coverage) and not on a sharp
  drop (SHARP_DROP sets the rate there).
  - **Backtest** (calibrate on the 8 oldest months, test on lead-time windows
    starting in the 3 latest; two real reports): lumpy — coverage 88.9% →
    88.2% with −6% units in one report, 90.2% → 91.6% with +4% in the other;
    sporadic — 94.1% → 93.0% (−4.5%) and 95.6% → 97.0% (−2%).
  - **Only with a current rate:** an item silent for the last six months keeps
    rate 0 — the weighting never invents demand from earlier years.
  - **The advisory trend numbers** are bounded by the recommendation: «if the
    decline continues» is never above it, «if the rise continues» never below
    (needed now that the reorder point can use another rate).
  - **Measured against main:** report A — 246 reorder points down, 50 up, net
    −633 units (about −$23K and −₪4K FOB); in the SAP update file 11 in / 12
    out, 41 down / 9 up. Report B — 93 down, 211 up, net +221 units (about
    +SEK 217K); in the SAP file 25 in / 15 out, 47 down / 129 up. Demo: 29 items.
  - Rejected in the same study: an empirical quantile of lead-time sums (−4%
    to −8% units, −1 to −2 points coverage); σ measured on lead-time windows
    (no difference); the three-year average alone (coverage −5 points in a
    growing report, +14% stock in the other).

- **2026-11-19 · The 4-topics view: «burning now» strip and a frozen topics row
  (approved 18.11).** Display only.
  - A strip above the topics row, on every topic: customer orders without
    coverage — both kinds (no purchase at all / partial cover) — the same list
    and order as topic 1, without items marked «handled», for the selected
    planner. It shows the count, how many have no purchase, the order value
    per currency, and the three most urgent (red border = no purchase); a click
    opens the card, «to all orders» opens topic 1. Shortages without an ETA are
    deliberately not here (an ETA may open within a day or two).
  - The topics row is larger and bolder, and stays at the top together with the
    strip while scrolling; the table scrolls inside its own area and its header
    row stays visible.

- **2026-11-19 · Neutralizing an item: «general», «seasonality», «campaign»
  (NEUTRAL, approved 18.11).** A neutralized item leaves every work queue, the
  four topics, the «burning now» strip, the SAP update file and the profile-
  change recommendation — in both tools — until the planner cancels it (no
  expiry). It uses the existing mark mechanism (the item goes to «marked»,
  like «campaign» did). Changes:
  - Two new mark types, «neutralize — general» and «seasonality»; «campaign»
    no longer expires after half a year. «Handled» (7 days), «on demand» and
    the other marks are unchanged.
  - The 4-topics view: three buttons in the item card; «neutralized (N)» in
    the header lists them with reason, date and days, and «cancel».
  - The full tool: the two new types in the card and in bulk marking; the
    card says «until you cancel», with no number of days.
  - Saved in the browser, like every mark. The demo report has no marks, so
    the engine output and exports do not change.

- **2026-11-19 · Statistical safety stock for items with consumption (SS_STAT,
  approved 19.11). Replaces «SAP ssMin is not changed by the tool» (31.10).**
  For an item with current consumption (six-month path, rate > 0) the safety
  stock is max(statistical, low-demand policy floor); the SAP `ssMin` floor is
  no longer decisive. Items without consumption in the last six months (rate 0,
  or a rate from the three-year history, including RATE_DECAY items) keep the
  SAP floor, as before. Unchanged: the policy floor, minimum reorder point 1
  (ROP_MIN_ACT and PD→ZM-C), SS ≤ ROP.
  - **SAP update file:** when the proposed SS is below `ssMin`, the «minimum
    safety stock» column carries the proposed SS (otherwise SAP would keep the
    old floor); a row also enters when only the minimum changes. In every row
    SS ≥ the minimum in that row.
  - **Measured against main:** report A — 419 items, −4,257 SS units (about
    −$104K and −₪96K FOB), 419 reorder points down (361 below SAP's current
    value); SAP file 192 → 252 rows, 98 with a lowered minimum. Report B — 344
    items, −9,063 units (about −SEK 1.96M), 344 reorder points down (153 below
    SAP); SAP file 1,365 → 1,356 rows, 112 with a lowered minimum. No SS rises.
  - **Demo:** 284 items with consumption and an `ssMin` (the effective floor
    becomes 0); in 45 of them, where the SAP floor decided before, SS and ROP
    drop. The «SS floor» screen in the full tool goes 52 → 7.

- **2026-11-19 · Supplier-delay assumption lowered to 5% for every material
  type (LTV_5, approved 19.11). Changes the default p of SS_LTVAR (12.11).**
  σ_LT = p × lead time was 25% for imports (Z001) and unknown types, 10% for
  local purchase (Z004). It is now 5% for all three, until σ_LT is measured from
  ETA history. Unchanged: the formula, the 95% service level for active items,
  the exceptions (no consumption in 11 months; decay), the floors and SS_STAT.
  - **Before deciding (main after SS_STAT, two real reports):** at 25%/10% the
    delay term added about 1,170 units (about $32K) in report A and about 3,060
    units (about SEK 770K) in report B, almost all on imports and in a few dozen
    fast, long-lead-time items. Alternatives shown for imports: 20% (about a
    third less), 15% (more than half less), 10% (about 80% less). The planner
    chose 5%, and 5% for local purchase too, so one percentage covers all.
  - **Measured against main:** report A — 149 items, −1,108 units of SS and the
    same in reorder points (about −$31K and −₪1K); SAP file 252 → 254 rows
    (4 in, 2 out). Report B — 264 items, −2,872 units (about −SEK 739K); SAP file
    1,356 → 1,351 rows (14 in, 19 out). No SS rises. Median safety cushion in
    items selling 30+ a month: about 64 → 42 days of sales (A), 30 → 16 (B).
  - **Demo:** 196 items, SS and ROP down (−842 units), none up; one item moves
    «מספר לתקן ב-SAP» → «הפרש קטן מכדי לתקן».

- **2026-11-19 · Supplier-delay assumption raised to 10% for every material
  type (LTV_10, approved 19.11, the same day as LTV_5).** After LTV_5 the
  planner felt 5% was too low, and asked for a simulation on a slow item.
  - **Simulation (the engine itself; 4 units a month, lead time 60 days,
    import, 80% service):** safety stock at 5% / 10% / 25% — steady demand
    1 / 2 / 2, moderate 2 / 3 / 3, volatile 4 / 4 / 4, lumpy 6 / 6 / 7. At 95%:
    steady 2 / 2 / 4, moderate 4 / 4 / 6. In such an item the delay percentage
    moves safety stock by one unit at most; demand variation and the service
    level decide. The planner chose 10% for all material types.
  - **Measured against main (5%):** report A — 63 items, +172 units of SS and
    the same in reorder points (about +$4K and +₪1K); SAP file stays at 254
    rows. Report B — 93 items, +495 units (about +SEK 94K); SAP file 1,351 →
    1,352 rows (2 in, 1 out). No SS goes down. Median safety cushion in items
    selling 30+ a month: about 42 → 45 days of sales (A), 16 → 19 (B).
  - **Demo:** 150 items, SS and ROP up (+236 units), none down.
  - **Test fix found on the way:** the hand-computed oracle in
    `tests/ss-ltvar.test.js` used the plain rate also for RATE_W and SHARP_DROP
    items. It passed only because rounding never crossed a whole unit; at 10%
    one sporadic item did (2.996 against 3.003). The oracle now follows the
    approved rate rule (λ / three-month rate / rateW / rate, and σ from three
    months in a sharp drop), and it also passes against main's engine.

- **2026-11-20 · Supplier-delay assumption raised to 15% for every material
  type (LTV_15, approved 20.11, after LTV_10).** The planner reviewed 10% and
  asked what 15%, 20% and 25% would add, and which items move most.
  - **What was shown:** against 10%, 15% / 20% / 25% add about 260 / 600 / 980
    units in report A and about 680 / 1,490 / 2,380 units in report B. The ten
    items that move most at 15% are all fast, steady imports (50–650 units a
    month, lead time 60–120 days, mostly at 95% service). In report B, items
    with a 60-day lead time had only 11–22 days of sales as cushion at 10%, and
    even at 15% the tool's SS stays well below what SAP holds today. On a
    two-month lead time 10% means a typical delay of 6 days, 15% of 9 days,
    25% of 15 days. The planner chose 15%.
  - **Measured against main (10%):** report A — 81 items, +261 units of SS and
    the same in reorder points (about +$8K and +₪2K); SAP file 254 → 253 rows
    (one out). Report B — 111 items, +678 units (about +SEK 159K); SAP file
    1,352 → 1,350 rows (3 in, 5 out). No SS goes down. Median safety cushion in
    items selling 30+ a month: about 45 → 50 days of sales (A), 19 → 22 (B).
  - **Demo:** 170 items, SS and ROP up (+353 units), none down.
  - **Tests:** besides the p table, `tests/statuses.test.js` holds an exact
    control value for an active sporadic item; at 15% its SS is
    ⌈66.55⌉ = 67 (was 66) and its ROP 120 (was 119). The expected values were
    updated with the computation beside them.

- **2026-11-20 · The full tool: the planner arranges the columns, and an
  optional «FOB price» column, as in the 4-topics view (COL_ORDER / FOB_COL,
  approved 20.11).** Display only; the engine, the counters, the SAP file and
  the default Excel export are unchanged. This is the first of four PRs that
  bring the 4-topics display features to the full tool (next: the arrivals
  board and ETA by date in the card; «first consumption»; one supply action
  in words for «zero stock»).
  - **Column order:** ↑ / ↓ beside every shown column in «columns», or drag a
    header in the table, in the work tables and in the movements table. Saved
    in the browser, per view, in the same list the picker already keeps.
    «Part number» stays first and cannot move; so do the selection box and the
    movements dot. A column width set by hand moves with its column; dragging
    the width handle does not move a column. «Back to default» restores the
    original order.
  - **«FOB price»:** offered in the picker of every table, in no default. The
    unit price (BZ) in the item's currency, in the same format as the 4-topics
    view; «—» when there is no price. Sortable. Screen only: it does not go to
    the «view» sheet or the movements sheet (decision 1.11: no new columns in
    the full tool's Excel).
  - **The «view» sheet** keeps following the screen: the shown columns in the
    shown order, as it already did with show / hide.
  - **Fixed on the way:** a column ticked in the picker now lands right after
    its nearest preceding column in the tool's column order («FOB price» after
    «value»). Before, every ticked column landed right after the part number,
    because «model» sits second in every default but last in that order.

- **2026-11-20 · The full tool: an arrivals board, and deliveries grouped by
  date in the item card, as in the 4-topics view (ARRIVALS_FULL, approved
  20.11).** Display only; the engine, the counters, the SAP file and the Excel
  exports are unchanged. The second of the four PRs.
  - **Arrivals board:** a «arrivals board» button in the header of «purchasing
    and deliveries» replaces the table with the board; a second click brings
    the table back (the planner's choice: a button there, not a new door in
    the rail). Leaving the screen, or picking a station on the band above it,
    also brings the table back.
  - **Which items:** every item that has lines in the ETA report, after the
    search and the filters (the planner's choice; the 4-topics view shows only
    items in its four topics). One row per arrival date: number of items and
    units. Dates that passed come first, marked «check if received». A date
    opens its lines (part number, description, quantity, delivery, purchase
    order, supplier), largest first; a part number opens its card. Without an
    ETA report the board says so and offers the upload.
  - **The card:** «deliveries by date» — one line per date with the total
    units, the number of lines and the delivery numbers; a date that passed is
    marked «check if received». The same list as the 4-topics card. Items with
    no open purchase and no ETA lines get no such section.

- **2026-11-20 · The full tool: «first consumption» (FIRST_FULL, approved
  20.11).** Display only; the engine, the classifications, the work queues,
  the SAP file and the existing counters and exports are unchanged. The third
  of the four PRs.
  - **The rule:** the same as in the 4-topics view — an item never consumed in
    the report's history (not in the 11 months and not in the 3 years) that
    now has a first consumption, in the last full month or in the current
    month. Without the 3-year columns «never» cannot be said, and the rule
    does not apply. All such items, with no further filter: the question is
    continuity — whether to set up planning for the item.
  - **Where:** a door in «planning and MRP», right after «demand decline»,
    with a counter (the planner's choice: a door in the rail, not a header
    button and not the card only). The full tool now has 27 screens.
  - **The list:** part number, model, description, consumption in the last
    full month and in the current month (screen only, decision 1.11), free
    stock, open purchase, MRP type and classification; largest consumption
    first. The «view» sheet of this screen carries the existing column types
    only.
  - **The card:** «first consumption: N units in the last full month / M in
    the current month. No earlier consumption — not in the 11 months and not
    in the 3 years in the report. Check whether to set up planning for the
    item.» Shown with the peaks and drops in the history.
  - **Measured:** report A — 133 items (68 slow stock, 24 no action, 15 OK…;
    45 with an open purchase, 47 with no free stock); demo — 108 items.

- **2026-11-21 · The full tool: one supply action in words in the shortage
  screens, as in «zero stock» of the 4-topics view (SHORT_ACT_FULL, approved
  21.11).** Display only; the engine, the recommendations, the counters and
  the SAP file are unchanged. The last of the four PRs.
  - **Where:** the «next action» column in «shortages to handle» and in
    «purchasing and deliveries». It now says one supply action in words, with
    no quantity — the same rule as the 4-topics view: no open purchase → «open
    a purchase order»; customer waiting and an ETA → «give the customer a date:
    dd.mm.yyyy»; purchase without an ETA → «ask the supplier for an ETA»; ETA
    too late → «expedite the supplier»; otherwise → «check back order with the
    manufacturer». A number in the action is adopted unchecked; the quantity is
    set against the supplier and the stock.
  - **Not changed (the planner's choice):** the item card keeps the decision
    with its quantity, and the other screens keep the engine's sentence. The
    Excel export keeps the engine's full sentence in every screen.
  - **«Quantity to order»** leaves the default columns of «by supplier» (as
    «to order» left the 4-topics table). It stays in the column picker and in
    the «all data» sheet. The «view» sheet of «by supplier» no longer carries
    it — the planner approved this one export change.

- **2026-11-21 · ZM-C, stage 2 of the profile screen (ZMC_PR2; the decisions
  were closed earlier, the column list was approved 21.11).** The engine, the
  profile states, the rail counter, the SAP file and the 4-topics view are
  unchanged.
  - **Review instead of silence:** a PD → ZM-C candidate whose simulation found
    a reason (tools/accessories group, manual «on demand» mark, inactive status,
    missing lead time, no service-level column, missing or duplicate source row,
    ROP and SS of 0, or SS of 0) is shown as «check», with the reason in «why».
    It is no longer shown as «raise», and it does not silently stay PD.
  - **Two tabs:** «raise to ZM-C» and «downgrades and waiting».
    - Raise: part number · model · description · recommendation · why · net
      consumption 12 months · months with consumption · ROP after ZM-C · SS
      after ZM-C · free stock · in transit. «Profile today / proposed» are not
      shown there: it is always PD → ZM-C.
    - Downgrades and waiting: as before (today · proposed · months ·
      recommendation · why · free · in transit), without the simulation.
  - **Export:** only the «view» sheet of this screen changes. It carries the
    open tab, with the same columns in the same order, and the simulated ROP
    and SS are exported with «(simulation)» in their headers, so that nobody
    uploads them to SAP by mistake.
  - **The reason «ROP≥2 and SS=0»** was checked before any rename: since
    SS_LTVAR a candidate gets there only with no consumption in the 11 months
    (all of it in the current month, the rate from the 3-year window). The
    internal key stays; the wording now says this.
  - **Measured:** report A — 36 raise candidates (27 «raise», 9 «check»: 4 ROP
    and SS of 0 with no consumption in six months, 2 tools/accessories, 2 too
    little consumption, 1 SS of 0) and 7 downgrade items (all «waiting»).

- **2026-11-21 · «check» in the 4-topics view too (ZMC_TOPICS, approved
  21.11).** The same status as the full tool's profile screen: a PD → ZM-C
  candidate whose simulation found a reason shows «check» with the reason —
  in the «profile change» column of «data cleanup», in its action («check
  profile: PD → ZM-C»), in the item card («check — do not change the profile
  automatically», with every reason) and in the CSV («PD → ZM-C · check»).
  Before, these showed «raise». The topic's items, its counters and the
  profile states are unchanged; report A has 9 such items.

Details, numbers and open items: `docs/PROJECT-STATUS.md`.

---

# Summary

Inventory Planner is a working-capital and inventory-management tool.

Financial calculations are based on:

- BZ (FOB Price)
- Quantity

The system does not use selling price (BX) as part of inventory planning decisions.

BZ is the single financial source of truth.
