# DATA-MAPPING.md

Version: 1.0
Status: Draft
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

# Questions

The following mappings require business validation before implementation.

| Column | Field | Status |
|----------|----------|----------|
| BI | רמת שרות | Verify against current ZMRP export |

# Planning Parameters

| Column | Field | Purpose |
|----------|----------|----------|
| BI | רמת שרות | Service Level |
| BJ | מלאי בטחון | Current Safety Stock |
| BK | נק.הז.מחדש | Current Reorder Point |
| BL | אספ.מתוכנ. | Lead Time |
| BN | מל.בט.מינ. | Minimum Safety Stock |
