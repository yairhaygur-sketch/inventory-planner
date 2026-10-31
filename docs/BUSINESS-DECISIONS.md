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
- Classifications, thresholds and screen counters
- Excel and SAP exports (columns, filters, values — see `docs/SAP-EXPORT.md`)

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

## Approved but not implemented

- **ROP/SS stage 2 — reduction potential.** Shown only in the item card,
  never as a second recommendation and never in any export. To be planned
  together with D3 (item card), not before (order set 2026-10-31).
- **ZM-C PR2.** Items needing review show as «לבדוק» (no automatic PD);
  downgrades and pending items in a separate tab; check what
  `ss_zero_rate3y` means before renaming it; the export change is limited to
  the «התצוגה» sheet of the MRP-profile screen, and the column list and order
  are shown to the owner before implementation.

## Proposed — NOT approved

- **Stale SAP floor threshold ($250).** Flag a floor as stale when demand
  has decayed and the extra stock the floor creates exceeds $250
  (ILS×0.27, EUR×1.08). Tested in simulation only. It measures a lower
  target stock, not immediate savings. Must not be implemented as a rule
  without explicit approval.

## Tested and rejected

- **EWMA as the rate for all items, with broad floor removal (v2).**
  Service level fell and ILS inventory rose.

Details, numbers and open items: `docs/PROJECT-STATUS.md`.
SAP file rules: `docs/SAP-EXPORT.md`.

---

# Summary

Inventory Planner is a working-capital and inventory-management tool.

Financial calculations are based on:

- BZ (FOB Price)
- Quantity

The system does not use selling price (BX) as part of inventory planning decisions.

BZ is the single financial source of truth.
