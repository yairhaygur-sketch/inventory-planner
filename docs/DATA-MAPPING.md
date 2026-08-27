# DATA-MAPPING.md

Version: 1.0
Status: Draft
Owner: Yair Gur

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

