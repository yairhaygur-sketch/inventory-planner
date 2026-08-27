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

---

# Summary

Inventory Planner is a working-capital and inventory-management tool.

Financial calculations are based on:

- BZ (FOB Price)
- Quantity

The system does not use selling price (BX) as part of inventory planning decisions.

BZ is the single financial source of truth.
