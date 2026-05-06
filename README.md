# IHS–NPC Harmonizer · v2.0
### Consumables Master Data Harmonization & Validation System
**Rwanda FDA · SOP ODDG/RES/SOP/004**

---

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

The app opens at **http://localhost:8501**

---

## What It Does

A 6-step guided workflow that maps IHS consumables to NPC codes:

| Step | Action |
|------|--------|
| 1 | Upload IHS, NPC, PHC (required) + RHIC (optional) |
| 2 | Map columns from each file |
| 3 | Configure matching thresholds & validation rules |
| 4 | Run the matching engine (progress live) |
| 5 | Review results with filters & charts |
| 6 | Export formatted Excel + CSV |

---

## Matching Hierarchy (per PRD + SOP/004)

```
IHS Product
    │
    ├─ Exact normalized match ──────────→ EXACT · NPC · HIGH
    │
    ├─ Brand-generic alias match ───────→ BRAND_MATCH · NPC · HIGH
    │   (CORTEX→CORTICAL, CAPROSYN→POLYGLYCONATE, etc.)
    │
    ├─ Fuzzy match ≥ NPC threshold ─────→ SPEC_DIFF / BRAND_MATCH · NPC
    │
    ├─ PHC fallback ≥ PHC threshold ────→ SPEC_DIFF · PHC
    │
    └─ No match ────────────────────────→ NEW_SOP · UNMATCHED
```

---

## Validation Rules

| Rule | Source |
|------|--------|
| Anatomy incompatibility (Distal Radius ≠ DHS hip plate) | AO Foundation, Wikipedia DHS |
| Cerclage wire ≠ wire tightener | PMC3157064 |
| Redon drain ≠ chest drain | PMC8408575 |
| Product family cross-check | SOP Sec 10.3/10.5 |
| Numeric size tolerance (±0.1mm) | SOP engineering tolerance |
| Unit crossover (ml vs mm) | Clinical logic |

---

## Output Columns

| Column | Description |
|--------|-------------|
| RHIC_CODE | Original IHS product code |
| ORIGINAL_DESCRIPTION | IHS product description |
| NPC_CODE | Matched or generated NPC code |
| NPC_DESCRIPTION | NPC product description |
| IHBS_CODE | IHBS code from PHC (if available) |
| MATCH_SOURCE | NPC / PHC / UNMATCHED |
| MATCH_SCORE | Fuzzy score 0–100 |
| CONFIDENCE | HIGH / MEDIUM / LOW |
| MATCH_TYPE | EXACT / BRAND_MATCH / SPEC_DIFF / NEW_SOP |
| VALIDATION_STATUS | VALID / REVIEW |
| VALIDATION_COMMENT | Human-readable explanation |
| PRODUCT_FAMILY | SCREW / PLATE / SUTURE / TROCAR / etc. |

---

## File Format Requirements

All files must be `.xlsx` or `.csv`.

**IHS file**: at minimum a column with product descriptions.

**NPC file**: columns for NPC code and product description.

**PHC file**: columns for description, NPC code, and (optionally) IHBS code.

**RHIC file** (optional): column with product descriptions for comparison.
