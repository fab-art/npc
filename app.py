"""
╔══════════════════════════════════════════════════════════════╗
║   Consumables Master Data Harmonization & Validation System  ║
║   IHS – NPC – PHC – RHIC Integration Tool                   ║
║   Rwanda FDA  |  SOP ODDG/RES/SOP/004                       ║
╚══════════════════════════════════════════════════════════════╝
"""
import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import warnings
from datetime import datetime
from rapidfuzz import fuzz, process as rfprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IHS–NPC Harmonizer",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS  — Dark clinical aesthetic with Rwanda green accent
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Syne:wght@400;600;700;800&family=Inter:wght@300;400;500;600&display=swap');

:root {
    --bg:        #0d1117;
    --surface:   #161b22;
    --surface2:  #21262d;
    --border:    #30363d;
    --accent:    #00b347;
    --accent2:   #00e676;
    --warn:      #f0a500;
    --danger:    #f85149;
    --info:      #388bfd;
    --text:      #e6edf3;
    --muted:     #8b949e;
    --mono:      'JetBrains Mono', monospace;
    --head:      'Syne', sans-serif;
    --body:      'Inter', sans-serif;
}

/* Override Streamlit defaults */
.stApp { background: var(--bg); color: var(--text); font-family: var(--body); }
section[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}
.stButton > button {
    background: var(--accent); color: #000; font-family: var(--head);
    font-weight: 700; border: none; border-radius: 6px;
    padding: 0.5rem 1.4rem; letter-spacing: 0.02em;
    transition: all .15s ease;
}
.stButton > button:hover { background: var(--accent2); transform: translateY(-1px); box-shadow: 0 4px 20px rgba(0,179,71,.35); }
.stButton > button:disabled { background: var(--surface2); color: var(--muted); }
div[data-testid="stFileUploader"] {
    border: 1.5px dashed var(--border); border-radius: 8px;
    background: var(--surface); padding: 0.5rem;
}
div[data-testid="stFileUploader"]:hover { border-color: var(--accent); }
.stSelectbox > div > div, .stMultiSelect > div > div {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    color: var(--text) !important;
}
.stProgress > div > div { background: var(--accent); }
.stDataFrame { border: 1px solid var(--border); border-radius: 8px; }
hr { border-color: var(--border); }

/* Custom components */
.page-title {
    font-family: var(--head); font-size: 2.4rem; font-weight: 800;
    background: linear-gradient(135deg, var(--accent2), var(--info));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    letter-spacing: -0.02em; margin-bottom: 0;
}
.page-sub {
    font-family: var(--body); color: var(--muted); font-size: .95rem;
    margin-top: .25rem; margin-bottom: 1.5rem;
}
.section-head {
    font-family: var(--head); font-weight: 700; font-size: 1.15rem;
    color: var(--text); letter-spacing: 0.01em;
    border-left: 3px solid var(--accent); padding-left: .75rem;
    margin: 1.25rem 0 .75rem;
}
.step-badge {
    display: inline-block; font-family: var(--mono); font-size: .7rem;
    font-weight: 700; padding: 2px 8px; border-radius: 4px;
    background: var(--surface2); color: var(--muted); margin-bottom: .25rem;
}
.step-badge.active { background: rgba(0,179,71,.15); color: var(--accent2); }
.step-badge.done   { background: rgba(0,179,71,.08); color: var(--accent); }

.kpi-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 10px; padding: 1rem 1.25rem;
}
.kpi-val { font-family: var(--head); font-size: 2rem; font-weight: 800; }
.kpi-lbl { font-size: .78rem; color: var(--muted); text-transform: uppercase; letter-spacing: .08em; }
.kpi-exact   { color: var(--accent2); }
.kpi-brand   { color: #79c0ff; }
.kpi-spec    { color: var(--warn); }
.kpi-new     { color: var(--muted); }
.kpi-crit    { color: var(--danger); }

.match-pill {
    display: inline-block; font-family: var(--mono); font-size: .7rem;
    font-weight: 700; padding: 2px 8px; border-radius: 4px;
}
.pill-EXACT      { background: rgba(0,230,118,.15); color: #00e676; }
.pill-BRAND      { background: rgba(56,139,253,.15); color: #79c0ff; }
.pill-SPEC_DIFF  { background: rgba(240,165,0,.15);  color: #f0a500; }
.pill-NEW_SOP    { background: rgba(139,148,158,.1); color: #8b949e; }
.pill-CORRECTED  { background: rgba(248,81,73,.15);  color: #f85149; }
.pill-HIGH       { background: rgba(0,230,118,.15);  color: #00e676; }
.pill-MEDIUM     { background: rgba(240,165,0,.15);  color: #f0a500; }
.pill-LOW        { background: rgba(139,148,158,.1); color: #8b949e; }
.pill-CRITICAL   { background: rgba(248,81,73,.2);   color: #f85149; }
.pill-VALID      { background: rgba(0,230,118,.12);  color: #00e676; }
.pill-REVIEW     { background: rgba(248,81,73,.12);  color: #f85149; }

.file-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 10px; padding: 1rem 1.25rem; margin-bottom: .75rem;
}
.file-card.required { border-left: 3px solid var(--accent); }
.file-card.optional  { border-left: 3px solid var(--surface2); }
.file-card.loaded    { border-color: var(--accent); background: rgba(0,179,71,.06); }
.file-label  { font-family: var(--head); font-weight: 700; font-size: .95rem; }
.file-note   { font-size: .78rem; color: var(--muted); margin-top: .15rem; }

.val-valid  { color: var(--accent2); font-weight: 600; }
.val-review { color: var(--danger);  font-weight: 600; }

.tooltip-icon { color: var(--muted); font-size: .85rem; cursor: help; }
.rule-item { padding: .4rem .6rem; border-left: 2px solid var(--border);
             margin-bottom: .3rem; font-size: .85rem; color: var(--muted); }
.rule-item.active { border-color: var(--accent); color: var(--text); }

.console-box {
    background: #010409; border: 1px solid var(--border); border-radius: 8px;
    padding: 1rem 1.25rem; font-family: var(--mono); font-size: .78rem;
    color: #58a6ff; max-height: 280px; overflow-y: auto;
    line-height: 1.8;
}
.console-ok   { color: var(--accent2); }
.console-warn { color: var(--warn); }
.console-err  { color: var(--danger); }
.console-info { color: #58a6ff; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
STEPS = ["Upload", "Map Columns", "Configure", "Process", "Review", "Export"]

def init_state():
    defaults = {
        "step": 0,
        "files": {},          # raw uploaded DataFrames
        "col_map": {},        # {file_key: {role: col_name}}
        "config": {},         # matching thresholds, options
        "results": None,      # final DataFrame
        "logs": [],           # console logs
        "run_done": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()
S = st.session_state

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p style="font-family:\'Syne\',sans-serif;font-size:1.1rem;font-weight:800;color:#00e676;letter-spacing:.05em;margin-bottom:.1rem;">🏥 IHS–NPC</p>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:.72rem;color:#8b949e;margin-top:0;margin-bottom:1.5rem;">Harmonizer · Rwanda FDA</p>', unsafe_allow_html=True)

    st.markdown("---")
    for i, name in enumerate(STEPS):
        if i < S.step:
            st.markdown(f'<div class="step-badge done">✓ STEP {i+1}</div><div style="font-family:\'Syne\',sans-serif;font-size:.85rem;color:#00b347;margin-bottom:.5rem;">{name}</div>', unsafe_allow_html=True)
        elif i == S.step:
            st.markdown(f'<div class="step-badge active">→ STEP {i+1}</div><div style="font-family:\'Syne\',sans-serif;font-size:.9rem;font-weight:700;color:#e6edf3;margin-bottom:.5rem;">{name}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="step-badge">  STEP {i+1}</div><div style="font-size:.82rem;color:#484f58;margin-bottom:.5rem;">{name}</div>', unsafe_allow_html=True)

    st.markdown("---")
    if st.button("↩ Reset / Start Over", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    st.markdown('<div style="font-size:.7rem;color:#484f58;margin-top:2rem;">SOP ODDG/RES/SOP/004<br>Rwanda FDA · v2.0</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# KNOWLEDGE BASES (internet-verified)
# ─────────────────────────────────────────────────────────────────────────────
BRAND_ALIASES = {
    "CORTEX SCREW": "CORTICAL SCREW",
    "CORTEX SCREWS": "CORTICAL SCREWS",
    "CONCELOUS": "CANCELLOUS",
    "CONCEULOUS": "CANCELLOUS",
    "CANSELLOUS": "CANCELLOUS",
    "CAPROSYN": "POLYGLYCONATE",
    "VICRYL": "POLYGLACTIN 910",
    "MONOCRYL": "POLIGLECAPRONE 25",
    "PROLENE": "POLYPROPYLENE",
    "PDS": "POLYDIOXANONE",
    "SURGICEL": "OXIDIZED REGENERATED CELLULOSE",
    "SURGICELL": "OXIDIZED REGENERATED CELLULOSE",
    "HEMOLOCK": "HEM-O-LOK",
    "HEMOLOC": "HEM-O-LOK",
    "POLYSORB": "POLYGLYCOLIC ACID LACTIDE",
    "BIOSYN": "GLYCOMER 631",
    "VELOSORB": "POLYGLYCOLIC ACID LACTIDE",
    "MAXON": "POLYGLYCOLATE",
    "SURGIPRO": "POLYPROPYLENE",
    "TICRON": "COATED POLYESTER",
    "SURGIDAC": "POLYESTER",
    "SOFSILK": "SILK",
    "SURGILON": "BRAIDED NYLON",
    "MONOSOF": "NYLON MONOFILAMENT",
    "LIGASURE": "VESSEL SEALING SYSTEM",
    "INSUFLATION": "INSUFFLATION",
    "REDON": "CLOSED WOUND SUCTION DRAIN",
    "ETHIBOND": "POLYESTER BRAIDED SUTURE",
    "MERSILENE": "POLYESTER",
    "PROLENE": "POLYPROPYLENE",
    "NOVAFIL": "POLYBUTESTER",
    "VASCUFIL": "POLYPROPYLENE VASCULAR",
}

# Source: AO Foundation, Wikipedia DHS, PMC3157064, FDA 21 CFR 888
ANATOMY_INCOMPATIBLE = {
    "DISTAL RADIUS": ["DHS", "135", "INTERTROCHANTERIC", "HIP SCREW"],
    "VOLAR LOCKING": ["DHS", "135", "HIP"],
    "VOLAR LCP":     ["DHS", "135", "HIP"],
    "STALOID":       ["DHS", "135", "HIP"],
    "FIBULA":        ["DHS", "135", "HIP"],
    "TIBIA DISTAL":  ["DHS", "135", "HIP"],
    "TIBIA PROXIMAL":["DHS", "135", "HIP"],
    "TIBIA LOCKING": ["DHS", "135", "HIP"],
    "CLAVICLE":      ["DHS", "135", "TIBIA", "FEMUR", "HIP"],
    "CLAVICULA":     ["DHS", "135", "TIBIA", "FEMUR", "HIP"],
    "PROXIMAL HUMERUS":["DHS", "135", "HIP", "TIBIA"],
    "DISTAL HUMERUS":  ["DHS", "135", "HIP", "TIBIA"],
    # Source: PMC3157064 - cerclage wire ≠ wire tightener
    "FILS DE CERCLAGE": ["TIGHTNER", "TIGHTENER"],
    "CERCLAGE WIRE":    ["TIGHTNER", "TIGHTENER"],
    # Redon drain ≠ chest drain (PMC8408575)
    "REDON": ["CHEST DRAIN", "THORACIC DRAIN", "INTERCOSTAL"],
    "DRAIN DE REDON": ["CHEST DRAIN", "THORACIC", "INTERCOSTAL"],
}

PRODUCT_FAMILIES = {
    "SCREW":   r"\bSCREW\b|\bSCREWS\b|\bCORTEX\b|\bCORTICAL\b|\bCANCELLOUS\b|\bLOCKING SCREW\b|\bINTERLOCKING\b|\bCANNULATED SCREW\b",
    "PLATE":   r"\bPLATE\b|\bLCP\b|\bDCP\b|\bDHS\b",
    "NAIL":    r"\bNAIL\b|\bNAILS\b|\bPFNA\b|\bIMN\b",
    "KWIRE":   r"\bKIRSCHNER\b|\bK-WIRE\b|\bKWIRE\b|\bGUIDE PIN\b",
    "CERCLAGE":r"\bCERCLAGE\b|\bFILS DE CERCLAGE\b",
    "SUTURE":  r"\bSUTURE\b|POLYSORB|BIOSYN|CAPROSYN|SURGIPRO|TICRON|VELOSORB|SOFSILK|SURGILON|POLYCRYL|MAXON|CHROMIC GUT|PLAIN GUT|\bNYLON\b|\bSILK\b|\bSTEEL\b|V-LOC|MONOCRYL|VICRYL|PROLENE|MONOSOF",
    "TROCAR":  r"\bTROCAR\b|\bCANNULA\b|\bVERSAPORT\b|\bVERSASTEP\b",
    "STAPLER": r"\bSTAPLER\b|\bGIA\b|EEA|CEEA|ENDO GIA|LINEAR CUTTER",
    "DRAIN":   r"\bDRAIN\b|\bREDON\b",
    "ELECTRODE":r"\bELECTRODE\b|\bBIPOLAR\b|\bMONOPOLAR\b|\bPENCIL\b",
    "CLIP":    r"\bHEMOLOCK\b|\bENDOCLIP\b|\bLAPROCLIP\b|\bSURGICLIP\b",
    "MESH":    r"\bMESH\b",
    "BLADE":   r"\bBLADE\b",
}

SOP_GROUPS_500 = {
    "SCREW":    "CANC", "PLATE":    "MISP", "NAIL":     "NAIL",
    "KWIRE":    "KWIR", "CERCLAGE": "KWIR", "MESH":     "MESH",
}
SOP_GROUPS_200 = {
    "SUTURE": "SUTS", "TROCAR": "TROC", "STAPLER": "STAP",
    "DRAIN":  "DREN", "ELECTRODE": "ELCT", "CLIP": "ECLP",
    "BLADE":  "BLAD",
}

IMPLANT_PATTERNS = r"\bSCREW\b|\bPLATE\b|\bNAIL\b|\bKIRSCHNER\b|\bK-WIRE\b|\bCERCLAGE\b|\bCANCELLOUS\b|\bCORTEX\b|\bCORTICAL\b|\bINTERLOCKING\b|\bCANNULATED\b|\bANCHOR\b|\bMESH\b|\bPFNA\b|\bIMN\b"

# ─────────────────────────────────────────────────────────────────────────────
# PREPROCESSING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def normalize(text: str) -> str:
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    s = re.sub(r"(\d+),(\d+)", r"\1.\2", s)          # European decimals 1,5→1.5
    s = re.sub(r"[×✕]", "X", s)
    s = re.sub(r"\*", "", s)
    s = re.sub(r"\.0\s*MM", "MM", s)                  # 4.0MM → 4MM
    s = re.sub(r"(\d+\.?\d*)\s*MM", r"\1MM", s)
    s = re.sub(r"(\d+\.?\d*)\s*ML", r"\1ML", s)
    s = re.sub(r"\b(?:CH|FG)\s*(\d+)", r"FR\1", s)   # CH12 → FR12
    s = re.sub(r"\bFR\s*(\d+)", r"FR\1", s)
    s = re.sub(r"\b(\d+)\s*-\s*HOLES?\b", r"\1 HOLES", s)
    s = re.sub(r"\bHOLE\b", "HOLES", s)
    s = re.sub(r"\bLH\b", "LEFT", s)
    s = re.sub(r"\bRH\b", "RIGHT", s)
    s = re.sub(r"\bS\.T\.\b|\bS/T\b|\bSELF-TAPPING\b", "SELF TAPPING", s)
    s = re.sub(r"\s{2,}", " ", s)
    s = s.strip(" ,.-")
    return s

def apply_brand_aliases(text: str) -> str:
    for brand, generic in BRAND_ALIASES.items():
        text = re.sub(r"\b" + re.escape(brand.upper()) + r"\b", generic, text)
    return text

def generic_normalize(text: str) -> str:
    return apply_brand_aliases(normalize(text))

def extract_specs(text: str) -> dict:
    u = normalize(text)
    dims = re.findall(r"(\d+\.?\d*)MM", u)
    ml   = re.findall(r"(\d+\.?\d*)ML", u)
    fr   = re.findall(r"FR(\d+)", u)
    holes= re.search(r"(\d+)\s*HOLES?", u)
    swg  = re.search(r"(\d+)\s*SWG", u)
    usp  = re.search(r"\b(\d+)-?0\b", u)
    return {
        "dims_mm": sorted([float(d) for d in dims]),
        "ml":      [float(m) for m in ml],
        "fr":      [int(f) for f in fr],
        "holes":   int(holes.group(1)) if holes else None,
        "swg":     int(swg.group(1)) if swg else None,
        "usp":     usp.group(1) + "-0" if usp else None,
    }

def get_product_family(text: str) -> str:
    u = normalize(text)
    for fam, pat in PRODUCT_FAMILIES.items():
        if re.search(pat, u):
            return fam
    return "OTHER"

def is_implant(text: str) -> bool:
    return bool(re.search(IMPLANT_PATTERNS, normalize(text)))

# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def validate_match(rhic_desc: str, npc_desc: str, match_score: float) -> tuple[str, str]:
    """Returns (status, comment). Status = VALID | REVIEW"""
    issues = []

    u_r = normalize(rhic_desc)
    u_n = normalize(npc_desc)

    # Rule 1: Anatomy incompatibility (internet-sourced: AO Foundation, PMC)
    for rhic_kw, bad_npc_kws in ANATOMY_INCOMPATIBLE.items():
        if rhic_kw.upper() in u_r:
            for bad in bad_npc_kws:
                if bad.upper() in u_n:
                    issues.append(f"Anatomy mismatch: RHIC contains '{rhic_kw}' but NPC references '{bad}' — clinically incompatible (AO Foundation)")

    # Rule 2: Product family mismatch
    r_fam = get_product_family(rhic_desc)
    n_fam = get_product_family(npc_desc)
    if r_fam != "OTHER" and n_fam != "OTHER" and r_fam != n_fam:
        issues.append(f"Product type mismatch: RHIC={r_fam}, NPC={n_fam}")

    # Rule 3: Size / numeric consistency
    r_sp = extract_specs(rhic_desc)
    n_sp = extract_specs(npc_desc)

    if r_sp["dims_mm"] and n_sp["dims_mm"]:
        r_dia = r_sp["dims_mm"][0]; n_dia = n_sp["dims_mm"][0]
        if abs(r_dia - n_dia) >= 0.15:
            issues.append(f"Size mismatch: RHIC {r_dia}mm vs NPC {n_dia}mm (tolerance ±0.1mm per SOP)")

    if r_sp["ml"] and n_sp["ml"] and abs(r_sp["ml"][0] - n_sp["ml"][0]) > 0.5:
        issues.append(f"Volume mismatch: RHIC {r_sp['ml'][0]}ml vs NPC {n_sp['ml'][0]}ml")

    if r_sp["fr"] and n_sp["fr"] and r_sp["fr"][0] != n_sp["fr"][0]:
        issues.append(f"French size mismatch: RHIC FR{r_sp['fr'][0]} vs NPC FR{n_sp['fr'][0]}")

    if r_sp["holes"] and n_sp["holes"] and r_sp["holes"] != n_sp["holes"]:
        issues.append(f"Hole count differs: RHIC {r_sp['holes']} holes vs NPC {n_sp['holes']} holes")

    # Rule 4: Unit crossover (ml vs mm)
    r_has_ml = bool(r_sp["ml"]); r_has_mm = bool(r_sp["dims_mm"])
    n_has_ml = bool(n_sp["ml"]); n_has_mm = bool(n_sp["dims_mm"])
    if (r_has_ml and n_has_mm and not n_has_ml) or (r_has_mm and n_has_ml and not n_has_mm):
        issues.append("Unit crossover: one uses ml (volume), other uses mm (dimension) — different product types")

    if issues:
        return "REVIEW", " | ".join(issues)
    return "VALID", ""

# ─────────────────────────────────────────────────────────────────────────────
# MATCHING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def match_one(rhic_desc: str, npc_index: dict, phc_index: dict,
              rhic_index: dict | None, npc_thresh: int, phc_thresh: int) -> dict:
    """Match a single IHS product through NPC → PHC → UNMATCHED hierarchy."""

    norm_r   = normalize(rhic_desc)
    generic_r= generic_normalize(rhic_desc)

    # ── Step 1: Exact normalized match against NPC ────────────────────────────
    if norm_r in npc_index:
        entry = npc_index[norm_r]
        val_status, val_comment = validate_match(rhic_desc, entry["desc"], 100)
        return _result(rhic_desc, entry["code"], entry["desc"], 100, "NPC",
                       "EXACT", "HIGH", val_status, val_comment)

    # ── Step 2: Brand-stripped exact match against NPC ────────────────────────
    if generic_r in npc_index:
        entry = npc_index[generic_r]
        val_status, val_comment = validate_match(rhic_desc, entry["desc"], 100)
        return _result(rhic_desc, entry["code"], entry["desc"], 100, "NPC",
                       "BRAND_MATCH", "HIGH", val_status,
                       val_comment or "Terminology/brand normalised (e.g. CORTEX→CORTICAL)")

    # ── Step 3: Fuzzy match against NPC ──────────────────────────────────────
    npc_choices = list(npc_index.keys())
    if npc_choices:
        best = rfprocess.extractOne(generic_r, npc_choices, scorer=fuzz.token_sort_ratio)
        if best and best[1] >= npc_thresh:
            entry = npc_index[best[0]]
            score = best[1]
            mtype = "EXACT" if score >= 95 else ("BRAND_MATCH" if score >= 85 else "SPEC_DIFF")
            conf  = "HIGH" if score >= 80 else ("MEDIUM" if score >= npc_thresh else "LOW")
            val_status, val_comment = validate_match(rhic_desc, entry["desc"], score)
            return _result(rhic_desc, entry["code"], entry["desc"], score, "NPC",
                           mtype, conf, val_status, val_comment)

    # ── Step 4: PHC fallback ──────────────────────────────────────────────────
    if phc_index:
        phc_choices = list(phc_index.keys())
        best_phc = rfprocess.extractOne(generic_r, phc_choices, scorer=fuzz.token_sort_ratio)
        if best_phc and best_phc[1] >= phc_thresh:
            entry = phc_index[best_phc[0]]
            score = best_phc[1]
            conf  = "HIGH" if score >= 80 else ("MEDIUM" if score >= phc_thresh else "LOW")
            val_status, val_comment = validate_match(rhic_desc, entry["desc"], score)
            return _result(rhic_desc,
                           entry.get("npc_code", ""), entry["desc"], score, "PHC",
                           "SPEC_DIFF", conf, val_status, val_comment,
                           ihbs_code=entry.get("ihbs_code", ""))

    # ── Step 5: No match → generate SOP code ─────────────────────────────────
    fam = get_product_family(rhic_desc)
    if is_implant(rhic_desc):
        grp  = SOP_GROUPS_500.get(fam, "MISC")
        code = f"500{grp}NEW"
        cat  = "IMPLANT"
    else:
        grp  = SOP_GROUPS_200.get(fam, "CONS")
        code = f"200{grp}NEW"
        cat  = "CONSUMABLE"
    return _result(rhic_desc, code, rhic_desc, 0, "UNMATCHED",
                   "NEW_SOP", "LOW", "REVIEW",
                   f"No NPC or PHC match ≥ threshold — new SOP code required (SOP Sec 10.3/10.5)")


def _result(desc, npc_code, npc_desc, score, source, mtype, conf,
            val_status, val_comment, ihbs_code="") -> dict:
    return {
        "ORIGINAL_DESCRIPTION": desc,
        "NPC_CODE":             npc_code,
        "NPC_DESCRIPTION":      npc_desc,
        "IHBS_CODE":            ihbs_code,
        "MATCH_SCORE":          score,
        "MATCH_SOURCE":         source,
        "MATCH_TYPE":           mtype,
        "CONFIDENCE":           conf,
        "VALIDATION_STATUS":    val_status,
        "VALIDATION_COMMENT":   val_comment,
        "PRODUCT_FAMILY":       get_product_family(desc),
    }


def build_npc_index(df: pd.DataFrame, desc_col: str, code_col: str) -> dict:
    idx = {}
    for _, row in df.iterrows():
        desc = str(row[desc_col]).strip()
        code = str(row[code_col]).strip()
        n    = normalize(desc)
        ng   = generic_normalize(desc)
        entry = {"code": code, "desc": desc}
        idx[n]  = entry
        idx[ng] = entry
    return idx


def build_phc_index(df: pd.DataFrame, desc_col: str, npc_col: str, ihbs_col: str | None) -> dict:
    idx = {}
    for _, row in df.iterrows():
        desc = str(row[desc_col]).strip()
        n    = generic_normalize(desc)
        idx[n] = {
            "desc":      desc,
            "npc_code":  str(row[npc_col]).strip() if npc_col else "",
            "ihbs_code": str(row[ihbs_col]).strip() if ihbs_col and ihbs_col in row.index else "",
        }
    return idx

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Main sheet ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Harmonized Mapping"

    cols = list(df.columns)
    ws.append(cols)

    hdr_fill = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bdr
    ws.row_dimensions[1].height = 38

    # Column widths
    cw = {"A":22,"B":65,"C":18,"D":65,"E":18,"F":14,"G":16,"H":18,"I":16,"J":16,"K":65,"L":18}
    for c, w in cw.items():
        if c <= get_column_letter(ws.max_column):
            ws.column_dimensions[c].width = w

    # Color maps
    fill_map = {
        "NPC":       PatternFill("solid", fgColor="E2EFDA"),
        "PHC":       PatternFill("solid", fgColor="DDEBF7"),
        "UNMATCHED": PatternFill("solid", fgColor="FFF2CC"),
    }
    conf_fills = {
        "HIGH":   PatternFill("solid", fgColor="C6EFCE"),
        "MEDIUM": PatternFill("solid", fgColor="FFEB9C"),
        "LOW":    PatternFill("solid", fgColor="FCE4D6"),
    }
    val_fills = {
        "VALID":  PatternFill("solid", fgColor="C6EFCE"),
        "REVIEW": PatternFill("solid", fgColor="FCE4D6"),
    }

    src_col  = cols.index("MATCH_SOURCE") + 1 if "MATCH_SOURCE" in cols else None
    conf_col = cols.index("CONFIDENCE") + 1 if "CONFIDENCE" in cols else None
    val_col  = cols.index("VALIDATION_STATUS") + 1 if "VALIDATION_STATUS" in cols else None

    for row_data in df.itertuples(index=False):
        row_vals = list(row_data)
        ws.append(row_vals)
        xl_row = ws[ws.max_row]
        src  = row_vals[src_col-1]  if src_col  else ""
        conf = row_vals[conf_col-1] if conf_col else ""
        val  = row_vals[val_col-1]  if val_col  else ""
        rf   = fill_map.get(src, PatternFill("solid", fgColor="F5F5F5"))
        for cell in xl_row:
            cell.fill = rf
            cell.font = Font(name="Calibri", size=9)
            cell.alignment = Alignment(vertical="center")
            cell.border = bdr
        if conf_col: xl_row[conf_col-1].fill = conf_fills.get(conf, rf)
        if val_col:  xl_row[val_col-1].fill  = val_fills.get(val, rf)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    total = len(df)
    mt_vc = df["MATCH_TYPE"].value_counts() if "MATCH_TYPE" in df.columns else pd.Series()
    src_vc= df["MATCH_SOURCE"].value_counts() if "MATCH_SOURCE" in df.columns else pd.Series()
    val_vc= df["VALIDATION_STATUS"].value_counts() if "VALIDATION_STATUS" in df.columns else pd.Series()

    summary = [
        ["IHS–NPC Harmonization Report","","",""],
        [f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Rwanda FDA  |  SOP ODDG/RES/SOP/004","","",""],
        ["","","",""],
        ["MATCH SOURCE","COUNT","% TOTAL",""],
        ["NPC (Primary)",int(src_vc.get("NPC",0)),f"{src_vc.get('NPC',0)/total*100:.1f}%","Direct NPC match"],
        ["PHC (Fallback)",int(src_vc.get("PHC",0)),f"{src_vc.get('PHC',0)/total*100:.1f}%","PHC fallback used"],
        ["UNMATCHED",int(src_vc.get("UNMATCHED",0)),f"{src_vc.get('UNMATCHED',0)/total*100:.1f}%","New SOP code required"],
        ["TOTAL",total,"100%",""],
        ["","","",""],
        ["MATCH TYPE","COUNT","% TOTAL","DESCRIPTION"],
        ["EXACT",int(mt_vc.get("EXACT",0)),f"{mt_vc.get('EXACT',0)/total*100:.1f}%","Identical specs"],
        ["BRAND_MATCH",int(mt_vc.get("BRAND_MATCH",0)),f"{mt_vc.get('BRAND_MATCH',0)/total*100:.1f}%","Brand/terminology difference only"],
        ["SPEC_DIFF",int(mt_vc.get("SPEC_DIFF",0)),f"{mt_vc.get('SPEC_DIFF',0)/total*100:.1f}%","Same family, spec differences noted"],
        ["NEW_SOP",int(mt_vc.get("NEW_SOP",0)),f"{mt_vc.get('NEW_SOP',0)/total*100:.1f}%","No match — SOP code generated"],
        ["","","",""],
        ["VALIDATION","COUNT","% TOTAL",""],
        ["VALID",int(val_vc.get("VALID",0)),f"{val_vc.get('VALID',0)/total*100:.1f}%","Passes all validation rules"],
        ["REVIEW",int(val_vc.get("REVIEW",0)),f"{val_vc.get('REVIEW',0)/total*100:.1f}%","Requires manual review"],
    ]
    h_fill = PatternFill("solid", fgColor="1F3864")
    s_fill = PatternFill("solid", fgColor="1A5C1A")
    for i, row_vals in enumerate(summary, 1):
        ws2.append(row_vals)
        xl_row = ws2[i]
        if i <= 2:
            for cell in xl_row:
                cell.fill = h_fill
                cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        elif any(row_vals[0] == h for h in ["MATCH SOURCE","MATCH TYPE","VALIDATION"]):
            for cell in xl_row:
                cell.fill = s_fill
                cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        else:
            for cell in xl_row:
                cell.font = Font(name="Calibri", size=10)
        for cell in xl_row:
            cell.alignment = Alignment(vertical="center")
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 12
    ws2.column_dimensions["D"].width = 45

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: pill HTML
# ─────────────────────────────────────────────────────────────────────────────
def pill(label: str, kind: str) -> str:
    cls_map = {
        "EXACT":"EXACT","BRAND_MATCH":"BRAND","SPEC_DIFF":"SPEC_DIFF",
        "NEW_SOP":"NEW_SOP","CORRECTED":"CORRECTED",
        "HIGH":"HIGH","MEDIUM":"MEDIUM","LOW":"LOW","CRITICAL":"CRITICAL",
        "VALID":"VALID","REVIEW":"REVIEW",
        "NPC":"HIGH","PHC":"MEDIUM","UNMATCHED":"LOW",
    }
    cls = cls_map.get(kind, "LOW")
    return f'<span class="match-pill pill-{cls}">{label}</span>'

# ─────────────────────────────────────────────────────────────────────────────
# STEP 0  — UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
def step_upload():
    st.markdown('<h1 class="page-title">Data Upload</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Upload your source files to begin harmonization. IHS, NPC, and PHC files are required.</p>', unsafe_allow_html=True)

    FILE_DEFS = [
        ("ihs",  "IHS File",  "Consumables to be mapped", True,  ["RHIC Code", "Product Description"]),
        ("npc",  "NPC File",  "National Product Catalogue with codes", True,  ["NPC Code", "Product Description"]),
        ("phc",  "PHC File",  "PHC catalogue (fallback mapping source)", True, ["Description", "NPC Code", "IHBS Code"]),
        ("rhic", "RHIC File", "RHIC comparison reference (optional)", False, ["Description"]),
    ]

    all_required = True
    for key, label, note, required, cols in FILE_DEFS:
        loaded = key in S.files and S.files[key] is not None
        card_cls = "loaded" if loaded else ("required" if required else "optional")
        st.markdown(f'<div class="file-card {card_cls}">', unsafe_allow_html=True)
        badge = "🔴 REQUIRED" if required else "⚪ OPTIONAL"
        st.markdown(f'<div class="file-label">{label} &nbsp;<span style="font-size:.7rem;color:#8b949e;">{badge}</span></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="file-note">{note} · Expected columns: {", ".join(cols)}</div>', unsafe_allow_html=True)

        uploaded = st.file_uploader(f"Upload {label}", type=["xlsx","xls","csv"],
                                     key=f"up_{key}", label_visibility="collapsed")
        if uploaded:
            try:
                if uploaded.name.endswith(".csv"):
                    df = pd.read_csv(uploaded)
                else:
                    xl = pd.read_excel(uploaded, sheet_name=None)
                    # Show sheet selector if multiple sheets
                    sheet_names = list(xl.keys())
                    if len(sheet_names) > 1:
                        chosen = st.selectbox(f"Sheet ({label})", sheet_names, key=f"sheet_{key}")
                        df = xl[chosen]
                    else:
                        df = list(xl.values())[0]
                df.columns = [str(c).strip() for c in df.columns]
                S.files[key] = df
                st.success(f"✓ Loaded {len(df):,} rows, {len(df.columns)} columns")
            except Exception as e:
                st.error(f"Error reading file: {e}")
        elif loaded:
            df = S.files[key]
            st.info(f"✓ Previously loaded: {len(df):,} rows, {len(df.columns)} columns")

        if required and key not in S.files:
            all_required = False
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns([3, 1])
    with col2:
        can_proceed = all(k in S.files for k in ["ihs", "npc", "phc"])
        if st.button("Next: Map Columns →", disabled=not can_proceed, use_container_width=True):
            S.step = 1
            st.rerun()
    if not can_proceed:
        st.caption("⚠ Upload IHS, NPC, and PHC files to continue")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1  — COLUMN MAPPING
# ─────────────────────────────────────────────────────────────────────────────
def step_map():
    st.markdown('<h1 class="page-title">Column Mapping</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Tell the system which columns hold product descriptions, NPC codes, and IHBS codes.</p>', unsafe_allow_html=True)

    col_map = S.col_map

    def col_select(file_key, role, label, hint="", default_kws=None):
        df = S.files.get(file_key)
        if df is None: return
        options = list(df.columns)
        # Auto-detect best default
        def score(c):
            c_up = c.upper()
            if default_kws:
                return sum(kw.upper() in c_up for kw in default_kws)
            return 0
        sorted_opts = sorted(options, key=score, reverse=True)
        saved = col_map.get(file_key, {}).get(role)
        default_idx = sorted_opts.index(saved) if saved in sorted_opts else 0
        chosen = st.selectbox(label, sorted_opts, index=default_idx,
                              key=f"cm_{file_key}_{role}", help=hint)
        col_map.setdefault(file_key, {})[role] = chosen
        st.caption(f"Preview: {', '.join(str(v) for v in df[chosen].dropna().head(3).tolist())}")

    with st.container():
        st.markdown('<div class="section-head">IHS File — Source Products</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            col_select("ihs", "desc", "Product Description column",
                       "Column containing IHS product descriptions",
                       ["description","desc","product","consumable","item","name"])
        with c2:
            df_ihs = S.files.get("ihs")
            if df_ihs is not None:
                opts = ["(none)"] + list(df_ihs.columns)
                saved_code = col_map.get("ihs", {}).get("code", "(none)")
                idx_code = opts.index(saved_code) if saved_code in opts else 0
                chosen_code = st.selectbox("RHIC Code column (optional)", opts,
                                           index=idx_code, key="cm_ihs_code")
                col_map.setdefault("ihs", {})["code"] = chosen_code if chosen_code != "(none)" else None
                if chosen_code != "(none)":
                    st.caption(f"Preview: {', '.join(str(v) for v in df_ihs[chosen_code].dropna().head(3).tolist())}")

    st.markdown("")
    with st.container():
        st.markdown('<div class="section-head">NPC File — Reference Catalogue</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            col_select("npc", "code", "NPC Code column",
                       "Column containing NPC codes",
                       ["npc","code","product code"])
        with c2:
            col_select("npc", "desc", "NPC Description column",
                       "Column containing NPC product descriptions",
                       ["description","desc","product","npc desc","rw product"])

    st.markdown("")
    with st.container():
        st.markdown('<div class="section-head">PHC File — Fallback Reference</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            col_select("phc", "desc", "Description column",
                       "PHC product description",
                       ["description","desc","product","name","item"])
        with c2:
            col_select("phc", "npc_code", "NPC Code column",
                       "NPC code in PHC file",
                       ["npc","code","npc code"])
        with c3:
            df_phc = S.files.get("phc")
            if df_phc is not None:
                opts = ["(none)"] + list(df_phc.columns)
                saved = col_map.get("phc", {}).get("ihbs_code", "(none)")
                idx = opts.index(saved) if saved in opts else 0
                chosen = st.selectbox("IHBS Code column (optional)", opts, index=idx, key="cm_phc_ihbs")
                col_map.setdefault("phc", {})["ihbs_code"] = chosen if chosen != "(none)" else None
                if chosen != "(none)":
                    st.caption(f"Preview: {', '.join(str(v) for v in df_phc[chosen].dropna().head(3).tolist())}")

    if "rhic" in S.files:
        st.markdown("")
        with st.container():
            st.markdown('<div class="section-head">RHIC File — Comparison Reference (optional)</div>', unsafe_allow_html=True)
            col_select("rhic", "desc", "Description column",
                       "RHIC product description",
                       ["description","desc","product","name"])

    S.col_map = col_map

    st.markdown("---")
    c1, c2, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("← Back", use_container_width=True):
            S.step = 0; st.rerun()
    with c3:
        ready = (
            col_map.get("ihs", {}).get("desc") and
            col_map.get("npc", {}).get("code") and
            col_map.get("npc", {}).get("desc") and
            col_map.get("phc", {}).get("desc") and
            col_map.get("phc", {}).get("npc_code")
        )
        if st.button("Next: Configure →", disabled=not ready, use_container_width=True):
            S.step = 2; st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2  — CONFIGURE
# ─────────────────────────────────────────────────────────────────────────────
def step_configure():
    st.markdown('<h1 class="page-title">Matching Configuration</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Tune thresholds, select active validation rules, and configure output options.</p>', unsafe_allow_html=True)

    c1, c2 = st.columns([3, 2])

    with c1:
        st.markdown('<div class="section-head">Match Thresholds</div>', unsafe_allow_html=True)
        npc_thresh = st.slider("NPC minimum match score", 50, 100,
                               S.config.get("npc_thresh", 65),
                               help="Fuzzy score ≥ this value triggers an NPC match (PRD: ≥65)")
        phc_thresh = st.slider("PHC fallback minimum score", 40, 95,
                               S.config.get("phc_thresh", 60),
                               help="PHC used only when NPC score < NPC threshold (PRD: ≥60)")

        st.markdown('<div class="section-head">Confidence Bands</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            high_thresh  = st.number_input("HIGH confidence floor", 60, 100, S.config.get("high_thresh", 80))
        with col_b:
            med_thresh   = st.number_input("MEDIUM confidence floor", 40, 99, S.config.get("med_thresh", 65))

        st.markdown('<div class="section-head">Brand Normalisation</div>', unsafe_allow_html=True)
        use_brand = st.toggle("Apply brand→generic aliases before matching",
                              value=S.config.get("use_brand", True),
                              help="e.g. CORTEX→CORTICAL, CAPROSYN→POLYGLYCONATE")
        if use_brand:
            with st.expander(f"View {len(BRAND_ALIASES)} active aliases"):
                rows = [{"BRAND (RHIC)": k, "GENERIC (NPC)": v} for k, v in BRAND_ALIASES.items()]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, height=200)

    with c2:
        st.markdown('<div class="section-head">Validation Rules</div>', unsafe_allow_html=True)
        rules = {
            "anatomy":  ("Anatomy incompatibility check",   "Flags DISTAL RADIUS↔DHS, cerclage wire↔tightener, etc. (AO Foundation + PMC3157064)"),
            "family":   ("Product family mismatch check",   "Prevents SCREW↔PLATE, NAIL↔DRAIN cross-matching"),
            "size":     ("Numeric size consistency check",  "Flags mm, ml, FR gauge differences beyond tolerance"),
            "holes":    ("Plate hole count check",          "Validates hole count matches for plates"),
            "units":    ("Unit crossover check",            "Detects ml vs mm unit confusion (volume vs dimension)"),
        }
        active_rules = S.config.get("active_rules", list(rules.keys()))
        new_active = []
        for key, (label, desc) in rules.items():
            checked = st.checkbox(label, value=(key in active_rules), help=desc, key=f"rule_{key}")
            if checked: new_active.append(key)

        st.markdown('<div class="section-head">Output Options</div>', unsafe_allow_html=True)
        include_rhic = "rhic" in S.files and st.toggle("Include RHIC comparison column",
                                                        value=S.config.get("include_rhic", True))
        max_rows = st.number_input("Max rows to process (0 = all)", 0, 50000,
                                   S.config.get("max_rows", 0),
                                   help="Useful for testing on a subset")

    S.config = {
        "npc_thresh": npc_thresh,
        "phc_thresh": phc_thresh,
        "high_thresh": high_thresh,
        "med_thresh":  med_thresh,
        "use_brand":   use_brand,
        "active_rules":new_active,
        "include_rhic":include_rhic if "rhic" in S.files else False,
        "max_rows":    max_rows,
    }

    st.markdown("---")
    c1, _, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("← Back", use_container_width=True): S.step = 1; st.rerun()
    with c3:
        if st.button("Next: Run Matching →", use_container_width=True): S.step = 3; st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3  — PROCESS
# ─────────────────────────────────────────────────────────────────────────────
def step_process():
    st.markdown('<h1 class="page-title">Processing</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Running the matching and validation engine across all product records.</p>', unsafe_allow_html=True)

    if S.run_done and S.results is not None:
        n = len(S.results)
        st.success(f"✓ Processing complete — {n:,} rows matched")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("← Change Config", use_container_width=True):
                S.run_done = False; S.step = 2; st.rerun()
        with c2:
            if st.button("View Results →", use_container_width=True):
                S.step = 4; st.rerun()
        return

    cfg = S.config
    cm  = S.col_map

    # Build product list
    df_ihs  = S.files["ihs"]
    df_npc  = S.files["npc"]
    df_phc  = S.files["phc"]

    desc_col   = cm["ihs"]["desc"]
    code_col   = cm["ihs"].get("code")
    npc_d_col  = cm["npc"]["desc"]
    npc_c_col  = cm["npc"]["code"]
    phc_d_col  = cm["phc"]["desc"]
    phc_n_col  = cm["phc"]["npc_code"]
    phc_i_col  = cm["phc"].get("ihbs_code")

    products = df_ihs[[c for c in [code_col, desc_col] if c]].dropna(subset=[desc_col])
    if cfg["max_rows"] > 0:
        products = products.head(cfg["max_rows"])
    total = len(products)

    logs = []
    log_box = st.empty()

    def log(msg, kind="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        cls = {"ok":"console-ok","warn":"console-warn","err":"console-err","info":"console-info"}.get(kind,"console-info")
        logs.append(f'<span class="{cls}">[{ts}] {msg}</span>')
        html = '<div class="console-box">' + "<br>".join(logs[-30:]) + "</div>"
        log_box.markdown(html, unsafe_allow_html=True)

    log(f"Loading {total:,} IHS products...")
    log(f"Threshold NPC≥{cfg['npc_thresh']}  PHC≥{cfg['phc_thresh']}")

    # Build indices
    log("Building NPC search index...")
    npc_idx = build_npc_index(df_npc, npc_d_col, npc_c_col)
    log(f"NPC index: {len(npc_idx):,} entries", "ok")

    log("Building PHC fallback index...")
    phc_idx = build_phc_index(df_phc, phc_d_col, phc_n_col, phc_i_col)
    log(f"PHC index: {len(phc_idx):,} entries", "ok")

    # RHIC index
    rhic_idx = None
    if cfg.get("include_rhic") and "rhic" in S.files:
        df_rhic = S.files["rhic"]
        rhic_d  = cm.get("rhic", {}).get("desc")
        if rhic_d:
            rhic_idx = {normalize(str(r[rhic_d])): str(r[rhic_d]) for _, r in df_rhic.iterrows()}
            log(f"RHIC index: {len(rhic_idx):,} entries", "ok")

    log("Starting matching loop...")
    progress = st.progress(0, text="Matching products...")

    results = []
    for i, row in enumerate(products.itertuples(index=False)):
        desc = str(getattr(row, desc_col.replace(" ","_"), "")).strip()
        rhic_code = str(getattr(row, code_col.replace(" ","_"), "")).strip() if code_col else f"IHS-{i+1:04d}"

        if not desc or desc.lower() == "nan":
            continue

        match = match_one(desc, npc_idx, phc_idx, rhic_idx,
                          cfg["npc_thresh"], cfg["phc_thresh"])

        # Adjust confidence based on config thresholds
        score = match["MATCH_SCORE"]
        if score >= cfg["high_thresh"]:    match["CONFIDENCE"] = "HIGH"
        elif score >= cfg["med_thresh"]:   match["CONFIDENCE"] = "MEDIUM"
        elif score > 0:                    match["CONFIDENCE"] = "LOW"

        # Add RHIC comparison
        if rhic_idx:
            norm_desc = generic_normalize(desc)
            choices = list(rhic_idx.keys())
            best = rfprocess.extractOne(norm_desc, choices, scorer=fuzz.token_sort_ratio)
            match["RHIC_MATCH"]      = rhic_idx.get(best[0], "") if best and best[1] >= 70 else ""
            match["RHIC_SCORE"]      = best[1] if best else 0
            match["RHIC_MATCH_TYPE"] = ("EXACT" if best and best[1] >= 95 else
                                         ("CATEGORY" if best and best[1] >= 70 else "NONE"))

        match["RHIC_CODE"] = rhic_code
        results.append(match)

        if i % 50 == 0 or i == total - 1:
            pct = (i + 1) / total
            progress.progress(pct, text=f"Matching {i+1:,}/{total:,} products...")
            if i % 200 == 0 and i > 0:
                matched = sum(1 for r in results if r["MATCH_SOURCE"] != "UNMATCHED")
                log(f"  {i+1:,}/{total:,} processed · {matched:,} matched ({matched/len(results)*100:.0f}%)")

    progress.progress(1.0, text="Done!")
    log(f"Matching complete — {len(results):,} rows", "ok")

    df_results = pd.DataFrame(results)
    # Reorder columns nicely
    front = ["RHIC_CODE", "ORIGINAL_DESCRIPTION", "NPC_CODE", "NPC_DESCRIPTION",
             "IHBS_CODE", "MATCH_SOURCE", "MATCH_SCORE", "CONFIDENCE",
             "MATCH_TYPE", "VALIDATION_STATUS", "VALIDATION_COMMENT", "PRODUCT_FAMILY"]
    rhic_extra = ["RHIC_MATCH", "RHIC_SCORE", "RHIC_MATCH_TYPE"]
    ordered = [c for c in front if c in df_results.columns]
    ordered += [c for c in rhic_extra if c in df_results.columns]
    df_results = df_results[ordered]

    S.results  = df_results
    S.logs     = logs
    S.run_done = True

    # Quick stats
    n_npc   = (df_results["MATCH_SOURCE"] == "NPC").sum()
    n_phc   = (df_results["MATCH_SOURCE"] == "PHC").sum()
    n_unm   = (df_results["MATCH_SOURCE"] == "UNMATCHED").sum()
    n_rev   = (df_results["VALIDATION_STATUS"] == "REVIEW").sum()

    log(f"NPC matched: {n_npc:,} · PHC fallback: {n_phc:,} · Unmatched: {n_unm:,}", "ok")
    log(f"Validation — VALID: {len(results)-n_rev:,} · REVIEW: {n_rev:,}", "warn" if n_rev > 0 else "ok")
    log("Ready for review.", "ok")

    st.success(f"✓ Complete — {len(results):,} products processed")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Reconfigure", use_container_width=True):
            S.run_done = False; S.step = 2; st.rerun()
    with col2:
        if st.button("View Results →", use_container_width=True):
            S.step = 4; st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4  — REVIEW
# ─────────────────────────────────────────────────────────────────────────────
def step_review():
    st.markdown('<h1 class="page-title">Results Review</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Inspect, filter, and validate your harmonized mapping before export.</p>', unsafe_allow_html=True)

    df = S.results
    if df is None or len(df) == 0:
        st.warning("No results yet. Go back and run processing.")
        if st.button("← Back to Processing"):
            S.step = 3; st.rerun()
        return

    total = len(df)
    n_npc = (df["MATCH_SOURCE"] == "NPC").sum() if "MATCH_SOURCE" in df.columns else 0
    n_phc = (df["MATCH_SOURCE"] == "PHC").sum() if "MATCH_SOURCE" in df.columns else 0
    n_unm = (df["MATCH_SOURCE"] == "UNMATCHED").sum() if "MATCH_SOURCE" in df.columns else 0
    n_rev = (df["VALIDATION_STATUS"] == "REVIEW").sum() if "VALIDATION_STATUS" in df.columns else 0
    n_hi  = (df["CONFIDENCE"] == "HIGH").sum() if "CONFIDENCE" in df.columns else 0
    coverage = (n_npc + n_phc) / total * 100 if total else 0

    # KPI row
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    kpis = [
        (k1, str(total), "Total Products", ""),
        (k2, f"{n_npc:,}", "NPC Matched", "kpi-exact"),
        (k3, f"{n_phc:,}", "PHC Fallback", "kpi-brand"),
        (k4, f"{n_unm:,}", "Unmatched", "kpi-new"),
        (k5, f"{n_rev:,}", "Needs Review", "kpi-spec"),
        (k6, f"{coverage:.0f}%", "Coverage", "kpi-exact" if coverage >= 80 else "kpi-spec"),
    ]
    for col, val, lbl, cls in kpis:
        with col:
            st.markdown(f'<div class="kpi-card"><div class="kpi-val {cls}">{val}</div><div class="kpi-lbl">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("")

    # ── Filters ───────────────────────────────────────────────────────────────
    st.markdown('<div class="section-head">Filters</div>', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        src_opts = ["All Sources"] + sorted(df["MATCH_SOURCE"].dropna().unique().tolist()) if "MATCH_SOURCE" in df.columns else ["All Sources"]
        f_source = st.selectbox("Match Source", src_opts)
    with f2:
        conf_opts = ["All Confidence"] + sorted(df["CONFIDENCE"].dropna().unique().tolist()) if "CONFIDENCE" in df.columns else ["All Confidence"]
        f_conf = st.selectbox("Confidence", conf_opts)
    with f3:
        val_opts = ["All Status"] + sorted(df["VALIDATION_STATUS"].dropna().unique().tolist()) if "VALIDATION_STATUS" in df.columns else ["All Status"]
        f_val = st.selectbox("Validation", val_opts)
    with f4:
        fam_opts = ["All Families"] + sorted(df["PRODUCT_FAMILY"].dropna().unique().tolist()) if "PRODUCT_FAMILY" in df.columns else ["All Families"]
        f_fam = st.selectbox("Product Family", fam_opts)

    view = df.copy()
    if f_source  != "All Sources"    and "MATCH_SOURCE"       in view.columns: view = view[view["MATCH_SOURCE"] == f_source]
    if f_conf    != "All Confidence" and "CONFIDENCE"         in view.columns: view = view[view["CONFIDENCE"] == f_conf]
    if f_val     != "All Status"     and "VALIDATION_STATUS"  in view.columns: view = view[view["VALIDATION_STATUS"] == f_val]
    if f_fam     != "All Families"   and "PRODUCT_FAMILY"     in view.columns: view = view[view["PRODUCT_FAMILY"] == f_fam]

    search = st.text_input("🔍 Search description", placeholder="e.g. locking screw 3.5mm")
    if search:
        mask = view.apply(lambda row: search.upper() in " ".join(str(v).upper() for v in row.values), axis=1)
        view = view[mask]

    st.caption(f"Showing {len(view):,} of {total:,} rows")

    # ── Table ─────────────────────────────────────────────────────────────────
    # Style the dataframe
    def style_row(row):
        styles = [""] * len(row)
        src = row.get("MATCH_SOURCE", "")
        val = row.get("VALIDATION_STATUS", "")
        conf= row.get("CONFIDENCE", "")
        bg = ""
        if src == "UNMATCHED":  bg = "background-color: rgba(248,193,79,0.1)"
        elif src == "PHC":      bg = "background-color: rgba(56,139,253,0.08)"
        elif val == "REVIEW":   bg = "background-color: rgba(248,81,73,0.08)"
        else:                   bg = "background-color: rgba(0,179,71,0.05)"
        return [bg] * len(row)

    display_cols = ["RHIC_CODE", "ORIGINAL_DESCRIPTION", "NPC_CODE", "NPC_DESCRIPTION",
                    "MATCH_SOURCE", "MATCH_SCORE", "CONFIDENCE", "MATCH_TYPE",
                    "VALIDATION_STATUS", "VALIDATION_COMMENT"]
    show_cols = [c for c in display_cols if c in view.columns]

    styled = view[show_cols].style.apply(style_row, axis=1).format(
        {"MATCH_SCORE": "{:.0f}"}
    )
    st.dataframe(styled, use_container_width=True, height=420)

    # ── Charts ────────────────────────────────────────────────────────────────
    st.markdown("")
    ch1, ch2, ch3 = st.columns(3)
    with ch1:
        st.markdown('<div class="section-head" style="font-size:.9rem;">Match Source</div>', unsafe_allow_html=True)
        if "MATCH_SOURCE" in df.columns:
            vc = df["MATCH_SOURCE"].value_counts().reset_index()
            vc.columns = ["Source", "Count"]
            st.bar_chart(vc.set_index("Source"), color="#00b347", height=180)
    with ch2:
        st.markdown('<div class="section-head" style="font-size:.9rem;">Confidence Level</div>', unsafe_allow_html=True)
        if "CONFIDENCE" in df.columns:
            vc = df["CONFIDENCE"].value_counts().reset_index()
            vc.columns = ["Confidence", "Count"]
            st.bar_chart(vc.set_index("Confidence"), color="#388bfd", height=180)
    with ch3:
        st.markdown('<div class="section-head" style="font-size:.9rem;">Product Family</div>', unsafe_allow_html=True)
        if "PRODUCT_FAMILY" in df.columns:
            vc = df["PRODUCT_FAMILY"].value_counts().head(8).reset_index()
            vc.columns = ["Family", "Count"]
            st.bar_chart(vc.set_index("Family"), color="#f0a500", height=180)

    st.markdown("---")
    c1, _, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("← Back to Process", use_container_width=True):
            S.step = 3; st.rerun()
    with c3:
        if st.button("Export Results →", use_container_width=True):
            S.step = 5; st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5  — EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def step_export():
    st.markdown('<h1 class="page-title">Export</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-sub">Download the complete harmonized dataset as a formatted Excel workbook.</p>', unsafe_allow_html=True)

    df = S.results
    if df is None:
        st.warning("No results to export.")
        return

    total = len(df)
    n_npc = (df["MATCH_SOURCE"] == "NPC").sum() if "MATCH_SOURCE" in df.columns else 0
    n_phc = (df["MATCH_SOURCE"] == "PHC").sum() if "MATCH_SOURCE" in df.columns else 0
    n_rev = (df["VALIDATION_STATUS"] == "REVIEW").sum() if "VALIDATION_STATUS" in df.columns else 0
    cov   = (n_npc + n_phc) / total * 100 if total else 0

    st.markdown('<div class="section-head">Export Summary</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-val kpi-exact">{total:,}</div><div class="kpi-lbl">Total Rows</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-val kpi-exact">{n_npc:,}</div><div class="kpi-lbl">NPC Matched</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-val kpi-spec">{n_rev:,}</div><div class="kpi-lbl">Needs Review</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-val {"kpi-exact" if cov>=80 else "kpi-spec"}">{cov:.0f}%</div><div class="kpi-lbl">Coverage</div></div>', unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="section-head">Workbook Contents</div>', unsafe_allow_html=True)
    for label in ["📋 Harmonized Mapping — all products with NPC code, IHBS code, confidence, validation",
                  "📊 Summary — match source breakdown, confidence distribution, validation stats"]:
        st.markdown(f'<div class="rule-item active">{label}</div>', unsafe_allow_html=True)

    st.markdown("")
    with st.spinner("Building Excel workbook..."):
        xlsx_bytes = build_excel(df)

    fname = f"IHS_NPC_Harmonized_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        label="⬇ Download Harmonized Excel",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("")
    # Also offer CSV
    csv_bytes = df.to_csv(index=False).encode()
    st.download_button(
        label="⬇ Download CSV (raw)",
        data=csv_bytes,
        file_name=fname.replace(".xlsx", ".csv"),
        mime="text/csv",
        use_container_width=False,
    )

    st.markdown("---")
    st.markdown(f'<div style="font-size:.8rem;color:#8b949e;text-align:center;">Generated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} · Rwanda FDA · SOP ODDG/RES/SOP/004 · IHS–NPC Harmonizer v2.0</div>', unsafe_allow_html=True)

    st.markdown("")
    if st.button("↩ Start New Session", use_container_width=False):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────────────────
STEP_FNS = [step_upload, step_map, step_configure, step_process, step_review, step_export]

current = S.step
if 0 <= current < len(STEP_FNS):
    STEP_FNS[current]()
