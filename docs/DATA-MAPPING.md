# DATA-MAPPING.md

Version: 1.0
Status: Verified against the production ZMRP export (114 columns, 2026-08-29)
Owner: Yair Gur

# Purpose

This document defines how Inventory Planner maps ZMRP fields into business logic.

Relationship to other documents:

- BUSINESS-DECISIONS.md explains WHY a field is used.
- DATA-MAPPING.md explains WHERE a field comes from.
- DECISION-RULES.md explains HOW the field is used.

This document is the single source of truth for ZMRP field mapping.

## Financial Fields

| Column | Field | Purpose |
|---------|---------|---------|
| BE | מטבע FOB | Currency |
| BZ | מחיר FOB | Financial Source of Truth |
| BX | מחיר מכירה | Not Used |

## Approved Rules

- BZ is the only financial source of truth.
- BE is used as the currency field.
- BX is not used in inventory planning decisions.

# Inventory Fields

| Column | Field | Purpose |
|----------|----------|----------|
| CA | סה"כ מלאי | Total Stock |
| DC | מלאי פנוי | Available Stock |
| CH | הז. רכש | Open Purchase Orders |

# Verification

All mappings below were checked character-by-character against the production
ZMRP export column list. 30 of 34 code lookups match the export header
directly; 2 resolve through a documented alternate name; 2 refer to columns
the export does not contain.

| Column | Field | Status |
|----------|----------|----------|
| BI | רמת שרות | **Verified.** Present in the export. |
| U | היררכייה3 | **Verified.** Accessories are identified by this description; the export has no separate "אביזרים" column. |
| CK | בהעברה | **Verified.** Now counted toward availability alongside open purchase orders. |
| I | סוג חומר | **Verified.** Z001 = import, Z004 = local. Filter and display only — no effect on any calculation. |

## Columns the export does not contain

| Field | Consequence |
|----------|----------|
| הזמנות לקוח חסומות | Blocked customer orders are not excluded from coverage. |
| תיאור קבוצת חומרים | Resolved through `תיא.קבוצ.חומרים` (L). |
| ETA · MOQ | Not computed. |

# Planning Parameters

| Column | Field | Purpose |
|----------|----------|----------|
| BI | רמת שרות | Service Level |
| BJ | מלאי בטחון | Current Safety Stock |
| BK | נק.הז.מחדש | Current Reorder Point |
| BL | אספ.מתוכנ. | Lead Time |
| BN | מל.בט.מינ. | Minimum Safety Stock |

# Demand History

| Column | Field | Purpose |
|----------|----------|----------|
| CL | צר.השנה | Current Year Demand |
| CM | צר.שנה-1 | Previous Year Demand |
| CN | צר.שנה-2 | Two Years Ago Demand |

# Monthly Consumption Pattern

| Columns | Purpose |
|----------|----------|
| CO-CZ | Last 12 Months Consumption History |

# Customer Demand

| Column | Field | Purpose |
|----------|----------|----------|
| DA | כמות בהז.פ | Open Customer Orders |

# Document Status

## Completed

- Financial Fields
- Inventory Fields
- Planning Parameters
- Demand History
- Monthly Consumption Pattern
- Customer Demand

## Pending Validation

- Service Level (BI)
- Additional ZMRP fields not yet documented

## Next Step

Compare all documented fields against the current ZMRP export and verify field names and business meaning.
