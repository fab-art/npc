"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  Consumables Master Data Harmonization & Validation System                  ║
║  IHS – NPC – PHC – RHIC Integration Tool   ·   v3.1 Ditto-Enhanced         ║
║  Rwanda FDA  |  SOP ODDG/RES/SOP/004                                        ║
╚══════════════════════════════════════════════════════════════════════════════╝

Online enrichment APIs:
  • openFDA Device Classification  — api.fda.gov/device/classification.json
  • openFDA 510k Clearance DB      — api.fda.gov/device/510k.json
  • DuckDuckGo Instant Answer API  — api.duckduckgo.com
  • SerpAPI / Google Search        — serpapi.com  (optional: SERPAPI_KEY)

Verify & Approve step — Ditto entity matching (local, no API key needed):
  • Uses pretrained language models (sentence-transformers) for semantic
    similarity scoring of (original_description, NPC_description) pairs.
  • Serialises records in Ditto's COL/VAL format before encoding.
  • pip install sentence-transformers   (adds ~400 MB model on first run)

API keys (optional — add to ~/.streamlit/secrets.toml on your server):
  openfda_api_key = "..."    # raises openFDA from 1000 to 120,000 req/day
  SERPAPI_KEY     = "..."    # enables Google Search enrichment via SerpAPI
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import hashlib
import time
import threading
import warnings
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

warnings.filterwarnings("ignore")

# ── Optional dependency guards ────────────────────────────────────────────────
try:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    from rapidfuzz import fuzz, process as rfprocess
    RAPIDFUZZ_OK = True
except ImportError:
    RAPIDFUZZ_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    from sentence_transformers import SentenceTransformer, util as st_util
    SBERT_OK = True
except ImportError:
    SBERT_OK = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IHS–NPC Harmonizer",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Syne:wght@400;600;700;800&family=Inter:wght@300;400;500;600&display=swap');
:root{--bg:#0d1117;--sf:#161b22;--sf2:#21262d;--bd:#30363d;--ac:#00b347;--ac2:#00e676;--warn:#f0a500;--err:#f85149;--info:#388bfd;--tx:#e6edf3;--mu:#8b949e;}
.stApp{background:var(--bg);color:var(--tx);font-family:'Inter',sans-serif;}
section[data-testid="stSidebar"]{background:var(--sf)!important;border-right:1px solid var(--bd);}
.stButton>button{background:var(--ac);color:#000;font-family:'Syne',sans-serif;font-weight:700;border:none;border-radius:6px;padding:.5rem 1.4rem;transition:all .15s;}
.stButton>button:hover{background:var(--ac2);transform:translateY(-1px);box-shadow:0 4px 16px rgba(0,179,71,.35);}
.stButton>button:disabled{background:var(--sf2);color:var(--mu);}
div[data-testid="stFileUploader"]{border:1.5px dashed var(--bd);border-radius:8px;background:var(--sf);padding:.5rem;}
div[data-testid="stFileUploader"]:hover{border-color:var(--ac);}
.stSelectbox>div>div{background:var(--sf2)!important;border:1px solid var(--bd)!important;border-radius:6px!important;}
.stProgress>div>div{background:var(--ac);}
hr{border-color:var(--bd);}
.pt{font-family:'Syne',sans-serif;font-size:2.4rem;font-weight:800;background:linear-gradient(135deg,var(--ac2),var(--info));-webkit-background-clip:text;-webkit-text-fill-color:transparent;letter-spacing:-.02em;margin-bottom:0;}
.ps{color:var(--mu);font-size:.93rem;margin-top:.2rem;margin-bottom:1.5rem;}
.sh{font-family:'Syne',sans-serif;font-weight:700;font-size:1.05rem;border-left:3px solid var(--ac);padding-left:.7rem;margin:1.1rem 0 .6rem;}
.sbg{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.68rem;font-weight:700;padding:2px 8px;border-radius:4px;background:var(--sf2);color:var(--mu);margin-bottom:.2rem;}
.sbg.a{background:rgba(0,179,71,.15);color:var(--ac2);}
.sbg.d{background:rgba(0,179,71,.08);color:var(--ac);}
.kc{background:var(--sf);border:1px solid var(--bd);border-radius:10px;padding:1rem 1.2rem;}
.kv{font-family:'Syne',sans-serif;font-size:2rem;font-weight:800;}
.kl{font-size:.75rem;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;}
.kg{color:var(--ac2);}.kb{color:#79c0ff;}.ky{color:var(--warn);}.kr{color:var(--err);}.km{color:var(--mu);}
.mp{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.68rem;font-weight:700;padding:2px 8px;border-radius:4px;}
.pE{background:rgba(0,230,118,.15);color:#00e676;}
.pB{background:rgba(56,139,253,.15);color:#79c0ff;}
.pS{background:rgba(240,165,0,.15);color:#f0a500;}
.pN{background:rgba(139,148,158,.1);color:#8b949e;}
.pV{background:rgba(0,230,118,.12);color:#00e676;}
.pR{background:rgba(248,81,73,.12);color:#f85149;}
.pO{background:rgba(255,200,0,.15);color:#ffc800;}
.fc{background:var(--sf);border:1px solid var(--bd);border-radius:10px;padding:1rem 1.2rem;margin-bottom:.7rem;}
.fc.r{border-left:3px solid var(--ac);}
.fc.o{border-left:3px solid var(--sf2);}
.fc.L{border-color:var(--ac);background:rgba(0,179,71,.06);}
.fl{font-family:'Syne',sans-serif;font-weight:700;font-size:.93rem;}
.fn{font-size:.76rem;color:var(--mu);margin-top:.12rem;}
.cb{background:#010409;border:1px solid var(--bd);border-radius:8px;padding:1rem 1.2rem;font-family:'JetBrains Mono',monospace;font-size:.74rem;color:#58a6ff;max-height:300px;overflow-y:auto;line-height:1.85;}
.co{color:var(--ac2);}.cw{color:var(--warn);}.ce{color:var(--err);}.ci{color:#58a6ff;}.cn{color:#ffc800;}
.ab{display:inline-flex;align-items:center;gap:.3rem;background:var(--sf2);border:1px solid var(--bd);border-radius:6px;padding:.2rem .55rem;font-size:.7rem;font-family:'JetBrains Mono',monospace;margin:.15rem;}
.ab.ok{border-color:#00b347;color:#00e676;}.ab.fl{border-color:#f85149;color:#f85149;}
.ec{background:rgba(255,200,0,.05);border:1px solid rgba(255,200,0,.2);border-radius:8px;padding:.7rem 1rem;margin:.4rem 0;font-size:.82rem;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
STEPS = ["Upload", "Map Columns", "Configure", "Process & Enrich", "Review", "Verify & Approve", "Export"]

def _init():
    defs = dict(step=0, files={}, col_map={}, config={},
                results=None, logs=[], run_done=False,
                api_status={}, enrich_cache={},
                verify_cache={}, verify_done=False)
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()
S = st.session_state

# ─────────────────────────────────────────────────────────────────────────────
# KNOWLEDGE BASES
# ─────────────────────────────────────────────────────────────────────────────
BRAND_ALIASES = {
    "CORTEX SCREW":"CORTICAL SCREW","CORTEX SCREWS":"CORTICAL SCREWS",
    "CONCELOUS":"CANCELLOUS","CONCEULOUS":"CANCELLOUS","CANSELLOUS":"CANCELLOUS",
    "CAPROSYN":"POLYGLYCONATE","CAPROSYN*":"POLYGLYCONATE",
    "VICRYL":"POLYGLACTIN 910","MONOCRYL":"POLIGLECAPRONE 25",
    "PROLENE":"POLYPROPYLENE","PDS":"POLYDIOXANONE","PDS II":"POLYDIOXANONE",
    "SURGICEL":"OXIDIZED REGENERATED CELLULOSE","SURGICELL":"OXIDIZED REGENERATED CELLULOSE",
    "HEMOLOCK":"HEM-O-LOK","HEMOLOC":"HEM-O-LOK",
    "POLYSORB":"POLYGLYCOLIC ACID LACTIDE","BIOSYN":"GLYCOMER 631",
    "VELOSORB":"POLYGLYCOLIC ACID LACTIDE","MAXON":"POLYGLYCOLATE",
    "SURGIPRO":"POLYPROPYLENE","TICRON":"COATED POLYESTER",
    "SURGIDAC":"POLYESTER","SOFSILK":"SILK","SURGILON":"BRAIDED NYLON",
    "MONOSOF":"NYLON MONOFILAMENT","LIGASURE":"VESSEL SEALING SYSTEM",
    "INSUFLATION":"INSUFFLATION","REDON":"CLOSED WOUND SUCTION DRAIN",
    "ETHIBOND":"POLYESTER BRAIDED SUTURE","MERSILENE":"POLYESTER",
    "NOVAFIL":"POLYBUTESTER","VASCUFIL":"POLYPROPYLENE VASCULAR",
    "AUTOSONIX":"ULTRASONIC DISSECTOR","SONICISION":"CORDLESS ULTRASONIC DISSECTOR",
    "PILLCAM":"VIDEO CAPSULE ENDOSCOPY","TRUCLEAR":"HYSTEROSCOPIC MORCELLATOR",
    "VERSAPORT":"LAPAROSCOPIC TROCAR","VERSASTEP":"LAPAROSCOPIC TROCAR SYSTEM",
    "VISIPORT":"OPTICAL ACCESS TROCAR",
}

# Internet-sourced anatomy incompatibilities
# Sources: AO Foundation, Wikipedia DHS, PMC3157064, PMC8408575
ANATOMY_INCOMPAT = {
    "DISTAL RADIUS":    ["DHS","135 DHS","INTERTROCHANTERIC","HIP SCREW"],
    "VOLAR LOCKING":    ["DHS","135 DHS","HIP"],
    "VOLAR LCP":        ["DHS","135 DHS","HIP"],
    "STALOID PLATE":    ["DHS","HIP"],
    "FIBULA":           ["DHS","135 DHS","HIP"],
    "TIBIA DISTAL":     ["DHS","135 DHS","HIP"],
    "TIBIA PROXIMAL":   ["DHS","135 DHS","HIP"],
    "TIBIA LOCKING":    ["DHS","135 DHS","HIP"],
    "CLAVICLE":         ["DHS","TIBIA","FEMUR","HIP","135 DHS"],
    "CLAVICULA":        ["DHS","TIBIA","FEMUR","HIP"],
    "PROXIMAL HUMERUS": ["DHS","HIP","TIBIA"],
    "DISTAL HUMERUS":   ["DHS","HIP","TIBIA"],
    "FILS DE CERCLAGE": ["TIGHTNER","TIGHTENER"],   # PMC3157064
    "CERCLAGE WIRE":    ["TIGHTNER","TIGHTENER"],
    "DRAIN DE REDON":   ["CHEST DRAIN","THORACIC","INTERCOSTAL"],  # PMC8408575
}

PRODUCT_FAMILIES = {
    "SCREW":     r"\bSCREW\b|\bSCREWS\b|\bCORTEX\b|\bCORTICAL\b|\bCANCELLOUS\b|\bINTERLOCKING\b|\bCANNULATED SCREW\b",
    "PLATE":     r"\bPLATE\b|\bLCP\b|\bDCP\b|\bDHS\b",
    "NAIL":      r"\bNAIL\b|\bNAILS\b|\bPFNA\b|\bIMN\b",
    "KWIRE":     r"\bKIRSCHNER\b|\bK-WIRE\b|\bKWIRE\b|\bGUIDE PIN\b",
    "CERCLAGE":  r"\bCERCLAGE\b|\bFILS DE CERCLAGE\b",
    "SUTURE":    r"\bSUTURE\b|POLYSORB|BIOSYN|CAPROSYN|SURGIPRO|TICRON|VELOSORB|SOFSILK|SURGILON|POLYCRYL|MAXON|CHROMIC GUT|PLAIN GUT|\bNYLON\b|V-LOC|MONOCRYL|VICRYL|PROLENE|MONOSOF",
    "TROCAR":    r"\bTROCAR\b|\bCANNULA\b(?!TED)|\bVERSAPORT\b|\bVERSASTEP\b",
    "STAPLER":   r"\bSTAPLER\b|\bGIA\b|\bEEA\b|\bCEEA\b|ENDO GIA|LINEAR CUTTER",
    "DRAIN":     r"\bDRAIN\b|\bREDON\b",
    "ELECTRODE": r"\bELECTRODE\b|\bBIPOLAR FORC\b|\bPENCIL\b",
    "CLIP":      r"\bHEMOLOCK\b|\bENDOCLIP\b|\bLAPROCLIP\b",
    "MESH":      r"\bMESH\b",
    "BLADE":     r"\bBLADE\b",
}

IMPLANT_PAT = r"\bSCREW\b|\bPLATE\b|\bNAIL\b|\bKIRSCHNER\b|\bK-WIRE\b|\bCERCLAGE\b|\bCANCELLOUS\b|\bCORTEX\b|\bCORTICAL\b|\bINTERLOCKING\b|\bCANNULATED\b|\bMESH\b|\bPFNA\b|\bIMN\b|\bANCHOR\b"

SOP500 = {"SCREW":"CANC","PLATE":"MISP","NAIL":"NAIL","KWIRE":"KWIR","CERCLAGE":"KWIR","MESH":"MESH"}
SOP200 = {"SUTURE":"SUTS","TROCAR":"TROC","STAPLER":"STAP","DRAIN":"DREN","ELECTRODE":"ELCT","CLIP":"ECLP","BLADE":"BLAD"}

# openFDA product code → Rwanda NPC category
FDA_TO_CAT = {
    "KWS":"IMPLANT","NDH":"IMPLANT","NDJ":"IMPLANT","NDL":"IMPLANT",
    "NDM":"IMPLANT","NDI":"IMPLANT","IYO":"IMPLANT","FRO":"IMPLANT",
    "GAG":"IMPLANT","QMP":"IMPLANT","GEI":"CONSUMABLE","GAL":"CONSUMABLE",
    "LZG":"CONSUMABLE","FSK":"CONSUMABLE","KZE":"CONSUMABLE","ITT":"CONSUMABLE",
}

# ─────────────────────────────────────────────────────────────────────────────
# HTTP LAYER — robust for hosted webserver
# ─────────────────────────────────────────────────────────────────────────────
_HTTP_SESSION = None
_req_lock = threading.Lock()
_last_req: dict = {}


def _get_session():
    global _HTTP_SESSION
    if _HTTP_SESSION is None and REQUESTS_OK:
        s = requests.Session()
        retry = Retry(total=3, backoff_factor=0.5,
                      status_forcelist=[429, 500, 502, 503, 504],
                      allowed_methods=["GET"])
        s.mount("https://", HTTPAdapter(max_retries=retry))
        s.mount("http://",  HTTPAdapter(max_retries=retry))
        s.headers.update({
            "User-Agent": "IHS-NPC-Harmonizer/3.0 (Rwanda-FDA; SOP/004)",
            "Accept": "application/json",
        })
        _HTTP_SESSION = s
    return _HTTP_SESSION


def http_get(url: str, params: dict = None, host_key: str = "default",
             timeout: int = 10, min_gap: float = 0.5):
    """Rate-limited GET → dict | None."""
    sess = _get_session()
    if sess is None:
        return None
    with _req_lock:
        now = time.time()
        gap = now - _last_req.get(host_key, 0)
        if gap < min_gap:
            time.sleep(min_gap - gap)
        _last_req[host_key] = time.time()
    try:
        r = sess.get(url, params=params, timeout=timeout)
        return r.json() if r.status_code == 200 else ({} if r.status_code == 404 else None)
    except Exception:
        return None

# ─────────────────────────────────────────────────────────────────────────────
# SEARCH ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _api_key(name: str) -> str:
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""


class SearchEngine:
    """
    Multi-source online enrichment engine.
    All results are cached in S.enrich_cache by md5(description).
    Thread-safe via per-host rate limiting in http_get().
    """

    # ── openFDA Device Classification ─────────────────────────────────────────
    @staticmethod
    def fda_classification(query: str) -> list:
        """
        https://api.fda.gov/device/classification.json
        Free, no key: 1,000 req/day  |  With key: 120,000 req/day
        Returns generic device names, FDA product codes, device class.
        """
        key = _api_key("openfda_api_key")
        # Try quoted search first, fall back to unquoted
        for q in [f'device_name:"{query}"', f"device_name:{'+'.join(query.split()[:4])}"]:
            params = {"search": q, "limit": 5}
            if key:
                params["api_key"] = key
            data = http_get("https://api.fda.gov/device/classification.json",
                            params, "fda_cls", min_gap=0.6)
            if data and data.get("results"):
                return [
                    {
                        "source":       "openFDA-Classification",
                        "device_name":  r.get("device_name", ""),
                        "product_code": r.get("product_code", ""),
                        "device_class": r.get("device_class", ""),
                        "definition":   r.get("definition", "")[:400],
                        "regulation":   r.get("regulation_number", ""),
                    }
                    for r in data["results"]
                ]
        return []

    # ── openFDA 510k Clearance ────────────────────────────────────────────────
    @staticmethod
    def fda_510k(query: str) -> list:
        """
        https://api.fda.gov/device/510k.json
        Cleared device generic names, applicant, decision date.
        """
        key = _api_key("openfda_api_key")
        params = {"search": f'device_name:"{query}"', "limit": 3,
                  "sort": "date_received:desc"}
        if key:
            params["api_key"] = key
        data = http_get("https://api.fda.gov/device/510k.json", params, "fda_510k", min_gap=0.6)
        if not data or not data.get("results"):
            return []
        return [
            {
                "source":        "openFDA-510k",
                "device_name":   r.get("device_name", ""),
                "generic_name":  r.get("generic_name", ""),
                "product_code":  r.get("product_code", ""),
                "applicant":     r.get("applicant", ""),
                "decision_date": r.get("decision_date_as_string", ""),
            }
            for r in data["results"]
        ]

    # ── DuckDuckGo Instant Answer ─────────────────────────────────────────────
    @staticmethod
    def duckduckgo(query: str) -> dict:
        """
        https://api.duckduckgo.com/?format=json
        Free, no key required. Returns abstract + related topics.
        Best for context / generic device descriptions.
        """
        data = http_get(
            "https://api.duckduckgo.com/",
            {"q": f"{query} medical device classification",
             "format": "json", "no_html": "1", "skip_disambig": "1"},
            "ddg", min_gap=1.0
        )
        if not data:
            return {}
        return {
            "source":   "DuckDuckGo",
            "abstract": data.get("AbstractText", ""),
            "topics":   [t.get("Text","") for t in data.get("RelatedTopics",[])
                         if isinstance(t, dict) and t.get("Text")][:3],
            "url":      data.get("AbstractURL", ""),
        }

    # ── SerpAPI (Google Search) ───────────────────────────────────────────────
    @staticmethod
    def serpapi(query: str) -> list:
        """
        https://serpapi.com/search.json
        Requires SERPAPI_KEY in st.secrets.
        Returns top-3 organic Google results.
        """
        key = _api_key("SERPAPI_KEY")
        if not key:
            return []
        data = http_get(
            "https://serpapi.com/search.json",
            {"q": f"{query} medical device generic name NPC classification",
             "api_key": key, "num": 5, "engine": "google"},
            "serpapi", min_gap=1.2
        )
        if not data:
            return []
        return [
            {"source": "SerpAPI-Google",
             "title":   r.get("title",""),
             "snippet": r.get("snippet",""),
             "link":    r.get("link","")}
            for r in data.get("organic_results", [])[:3]
        ]

    # ── Aggregate ─────────────────────────────────────────────────────────────
    @classmethod
    def enrich(cls, description: str, sources: list) -> dict:
        """
        Run all enabled sources, merge into one record.
        Results cached by md5(description) in S.enrich_cache.
        """
        cache_key = hashlib.md5(description.lower().encode()).hexdigest()
        if cache_key in S.enrich_cache:
            return S.enrich_cache[cache_key]

        # Build a clean query: strip specs, keep meaningful words
        clean = re.sub(r"\d+\.?\d*\s*(?:MM|ML|CM|FR|SWG)", "", description.upper())
        words = [w for w in clean.split() if len(w) > 2]
        query = " ".join(words[:7])

        result = {
            "ONLINE_GENERIC_NAME": "",
            "FDA_PRODUCT_CODE":    "",
            "FDA_DEVICE_CLASS":    "",
            "FDA_REGULATION":      "",
            "ONLINE_DEFINITION":   "",
            "ONLINE_SOURCE":       "",
            "ONLINE_CONFIDENCE":   "LOW",
            "SUGGESTED_CATEGORY":  "",
        }
        used = []

        if "openfda_class" in sources:
            hits = cls.fda_classification(query)
            if hits:
                top = hits[0]
                result.update({
                    "ONLINE_GENERIC_NAME": top["device_name"],
                    "FDA_PRODUCT_CODE":    top["product_code"],
                    "FDA_DEVICE_CLASS":    f"Class {top['device_class']}",
                    "FDA_REGULATION":      top["regulation"],
                    "ONLINE_DEFINITION":   top["definition"],
                    "SUGGESTED_CATEGORY":  FDA_TO_CAT.get(top["product_code"], ""),
                    "ONLINE_CONFIDENCE":   "HIGH",
                })
                used.append("openFDA-Class")

        if "openfda_510k" in sources and not result["ONLINE_GENERIC_NAME"]:
            hits = cls.fda_510k(query)
            if hits:
                top = hits[0]
                result["ONLINE_GENERIC_NAME"] = top.get("generic_name") or top.get("device_name","")
                result["FDA_PRODUCT_CODE"]    = result["FDA_PRODUCT_CODE"] or top.get("product_code","")
                result["ONLINE_CONFIDENCE"]   = "MEDIUM"
                used.append("openFDA-510k")

        if "duckduckgo" in sources and not result["ONLINE_DEFINITION"]:
            ddg = cls.duckduckgo(query)
            if ddg.get("abstract"):
                result["ONLINE_DEFINITION"] = ddg["abstract"][:350]
                result["ONLINE_CONFIDENCE"] = result["ONLINE_CONFIDENCE"] or "LOW"
                used.append("DuckDuckGo")

        if "serpapi" in sources:
            serp = cls.serpapi(query)
            if serp and not result["ONLINE_GENERIC_NAME"]:
                result["ONLINE_GENERIC_NAME"] = serp[0]["title"][:80]
                used.append("SerpAPI")

        result["ONLINE_SOURCE"] = " + ".join(used) if used else "OFFLINE"
        S.enrich_cache[cache_key] = result
        return result

    # ── Ping all APIs ─────────────────────────────────────────────────────────
    @staticmethod
    def ping_all() -> dict:
        status = {}
        d = http_get("https://api.fda.gov/device/classification.json",
                     {"search": "device_name:screw", "limit": 1}, "ping_fda", timeout=6)
        status["openFDA Classification"] = bool(d and "results" in d)

        d2 = http_get("https://api.fda.gov/device/510k.json",
                      {"search": "device_name:plate", "limit": 1}, "ping_510k", timeout=6)
        status["openFDA 510k"] = bool(d2 and "results" in d2)

        d3 = http_get("https://api.duckduckgo.com/",
                      {"q": "bone screw", "format": "json"}, "ping_ddg", timeout=6)
        status["DuckDuckGo"] = d3 is not None

        status["SerpAPI (Google)"] = bool(_api_key("SERPAPI_KEY"))
        status["Ditto / SBERT"]    = SBERT_OK
        return status

# ─────────────────────────────────────────────────────────────────────────────
# DITTO ENTITY MATCHER
# ─────────────────────────────────────────────────────────────────────────────
_DITTO_MODEL_NAME = "all-MiniLM-L6-v2"   # fast, 80 MB; swap for
                                           # "pritamdeka/S-PubMedBert-MS-MARCO"
                                           # for biomedical domain boost

def _ditto_serialize(desc: str) -> str:
    """
    Serialise a single product description in Ditto's COL/VAL format.
    e.g. "LOCKING SCREW 3.5MM" → "COL description VAL LOCKING SCREW 3.5MM"
    Multi-attribute records can extend this:
        "COL description VAL ... COL family VAL SCREW ..."
    """
    return f"COL description VAL {desc.strip()}"


class DittoMatcher:
    """
    Ditto-style entity matching for product description pairs.

    Architecture
    ────────────
    1. Serialise each record with _ditto_serialize() into the standard
       COL/VAL token format used by the original Ditto paper.
    2. Encode both serialised strings with a SentenceTransformer model,
       producing fixed-length sentence embeddings.
    3. Compute cosine similarity → match decision + confidence score.

    Thresholds (tunable via cls.T_MATCH / cls.T_POSSIBLE):
      ≥ T_MATCH    → MATCH       (high semantic overlap)
      ≥ T_POSSIBLE → POSSIBLE    (partial / ambiguous)
      <  T_POSSIBLE → NO_MATCH

    Thread-safety: model is loaded once and shared; encode() is read-only.
    """
    _model    = None
    _lock     = threading.Lock()
    T_MATCH   = 0.82   # cosine similarity → definite match
    T_POSSIBLE= 0.58   # cosine similarity → possible match

    @classmethod
    def _load_model(cls):
        if cls._model is None and SBERT_OK:
            with cls._lock:
                if cls._model is None:          # double-checked locking
                    cls._model = SentenceTransformer(_DITTO_MODEL_NAME)
        return cls._model

    @classmethod
    def available(cls) -> bool:
        return SBERT_OK

    @classmethod
    def match_pair(cls, desc_a: str, desc_b: str) -> dict:
        """
        Score a single (original, NPC) pair.
        Returns:
            { "match": MATCH|POSSIBLE|NO_MATCH|UNAVAILABLE,
              "score": float,  "note": str }
        """
        model = cls._load_model()
        if model is None:
            return {"match": "UNAVAILABLE", "score": 0.0,
                    "note": "sentence-transformers not installed — run: pip install sentence-transformers"}
        ser_a = _ditto_serialize(desc_a)
        ser_b = _ditto_serialize(desc_b)
        emb_a = model.encode(ser_a, convert_to_tensor=True, show_progress_bar=False)
        emb_b = model.encode(ser_b, convert_to_tensor=True, show_progress_bar=False)
        score = float(st_util.cos_sim(emb_a, emb_b)[0][0])
        return cls._decision(score)

    @classmethod
    def batch_match(cls, pairs: list[tuple]) -> list[dict]:
        """
        Efficiently encode a batch of (desc_a, desc_b) pairs in one pass.
        Returns a list of result dicts in the same order as pairs.
        """
        model = cls._load_model()
        if model is None:
            return [{"match": "UNAVAILABLE", "score": 0.0,
                     "note": "sentence-transformers not installed"}
                    for _ in pairs]
        sers_a = [_ditto_serialize(a) for a, _ in pairs]
        sers_b = [_ditto_serialize(b) for _, b in pairs]
        # Encode both lists in a single batched call (GPU-friendly if available)
        all_sers  = sers_a + sers_b
        all_embs  = model.encode(all_sers, convert_to_tensor=True,
                                  show_progress_bar=False, batch_size=64)
        embs_a = all_embs[:len(pairs)]
        embs_b = all_embs[len(pairs):]
        results = []
        for i in range(len(pairs)):
            score = float(st_util.cos_sim(embs_a[i], embs_b[i])[0][0])
            results.append(cls._decision(score))
        return results

    @classmethod
    def _decision(cls, score: float) -> dict:
        pct = f"{score:.1%}"
        if score >= cls.T_MATCH:
            return {"match": "MATCH",     "score": score,
                    "note": f"Strong semantic match ({pct})"}
        elif score >= cls.T_POSSIBLE:
            return {"match": "POSSIBLE",  "score": score,
                    "note": f"Possible match ({pct}) — review recommended"}
        else:
            return {"match": "NO_MATCH",  "score": score,
                    "note": f"Low semantic similarity ({pct}) — likely mismatch"}


# ─────────────────────────────────────────────────────────────────────────────
# TEXT PREPROCESSING
# ─────────────────────────────────────────────────────────────────────────────
def normalize(text: str) -> str:
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    s = re.sub(r"(\d+),(\d+)", r"\1.\2", s)
    s = re.sub(r"[×✕]", "X", s)
    s = re.sub(r"\*+", "", s)
    s = re.sub(r"\.0\s*MM", "MM", s)
    s = re.sub(r"(\d+\.?\d*)\s*MM", r"\1MM", s)
    s = re.sub(r"(\d+\.?\d*)\s*ML", r"\1ML", s)
    s = re.sub(r"\b(?:CH|FG)\s*(\d+)", r"FR\1", s)
    s = re.sub(r"\b(\d+)\s*-\s*HOLES?\b", r"\1 HOLES", s)
    s = re.sub(r"\bHOLE\b", "HOLES", s)
    s = re.sub(r"\bLH\b", "LEFT", s)
    s = re.sub(r"\bRH\b", "RIGHT", s)
    s = re.sub(r"\bS\.T\.\b|\bS/T\b|\bSELF-TAPPING\b", "SELF TAPPING", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip(" ,.-")

def generic_norm(text: str) -> str:
    s = normalize(text)
    for brand, gen in BRAND_ALIASES.items():
        s = re.sub(r"\b" + re.escape(brand.upper()) + r"\b", gen, s)
    return s

def extract_specs(text: str) -> dict:
    u = normalize(text)
    return {
        "dims": sorted([float(v) for v in re.findall(r"(\d+\.?\d*)MM", u)]),
        "ml":   [float(v) for v in re.findall(r"(\d+\.?\d*)ML", u)],
        "fr":   [int(v)   for v in re.findall(r"FR(\d+)", u)],
        "holes": int(m.group(1)) if (m := re.search(r"(\d+)\s*HOLES?", u)) else None,
    }

def get_family(text: str) -> str:
    u = normalize(text)
    for fam, pat in PRODUCT_FAMILIES.items():
        if re.search(pat, u):
            return fam
    return "OTHER"

def is_implant(text: str) -> bool:
    return bool(re.search(IMPLANT_PAT, normalize(text)))

# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def validate(rhic_d: str, npc_d: str) -> tuple:
    issues = []
    ur, un = normalize(rhic_d), normalize(npc_d)

    for kw, bad_list in ANATOMY_INCOMPAT.items():
        if kw.upper() in ur:
            for bad in bad_list:
                if bad.upper() in un:
                    issues.append(f"Anatomy incompatibility: '{kw}' cannot match '{bad}' (AO Foundation / PMC)")

    rf, nf = get_family(rhic_d), get_family(npc_d)
    if rf != "OTHER" and nf != "OTHER" and rf != nf:
        issues.append(f"Product family mismatch: RHIC={rf} vs NPC={nf}")

    rs, ns = extract_specs(rhic_d), extract_specs(npc_d)
    if rs["dims"] and ns["dims"] and abs(rs["dims"][0] - ns["dims"][0]) >= 0.15:
        issues.append(f"Size: RHIC {rs['dims'][0]}mm vs NPC {ns['dims'][0]}mm (tol ±0.1mm)")
    if rs["ml"] and ns["ml"] and abs(rs["ml"][0] - ns["ml"][0]) > 0.5:
        issues.append(f"Volume: {rs['ml'][0]}ml vs {ns['ml'][0]}ml")
    if rs["fr"] and ns["fr"] and rs["fr"][0] != ns["fr"][0]:
        issues.append(f"French: FR{rs['fr'][0]} vs FR{ns['fr'][0]}")
    if rs["holes"] and ns["holes"] and rs["holes"] != ns["holes"]:
        issues.append(f"Holes: {rs['holes']} vs {ns['holes']}")
    if bool(rs["ml"]) and bool(ns["dims"]) and not bool(rs["dims"]):
        issues.append("Unit crossover: RHIC=ml (volume), NPC=mm (dimension)")

    return ("REVIEW", " | ".join(issues)) if issues else ("VALID", "")

# ─────────────────────────────────────────────────────────────────────────────
# MATCHING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _res(desc, code, npc_d, score, src, mtype, conf, vs, vc, ihbs="", **kw):
    r = dict(ORIGINAL_DESCRIPTION=desc, NPC_CODE=code, NPC_DESCRIPTION=npc_d,
             IHBS_CODE=ihbs, MATCH_SCORE=score, MATCH_SOURCE=src,
             MATCH_TYPE=mtype, CONFIDENCE=conf,
             VALIDATION_STATUS=vs, VALIDATION_COMMENT=vc,
             PRODUCT_FAMILY=get_family(desc))
    r.update(kw)
    return r

def match_one(desc, npc_idx, phc_idx, npc_t, phc_t):
    nr = normalize(desc)
    gr = generic_norm(desc)

    if nr in npc_idx:
        e = npc_idx[nr]; vs, vc = validate(desc, e["desc"])
        return _res(desc, e["code"], e["desc"], 100, "NPC", "EXACT", "HIGH", vs, vc)

    if gr in npc_idx:
        e = npc_idx[gr]; vs, vc = validate(desc, e["desc"])
        return _res(desc, e["code"], e["desc"], 100, "NPC", "BRAND_MATCH", "HIGH", vs,
                    vc or "Brand/terminology normalised (CORTEX→CORTICAL etc.)")

    if RAPIDFUZZ_OK:
        choices = list(npc_idx.keys())
        best = rfprocess.extractOne(gr, choices, scorer=fuzz.token_sort_ratio)
        if best and best[1] >= npc_t:
            e = npc_idx[best[0]]; sc = best[1]
            mt = "EXACT" if sc >= 95 else ("BRAND_MATCH" if sc >= 85 else "SPEC_DIFF")
            cf = "HIGH" if sc >= 80 else ("MEDIUM" if sc >= npc_t else "LOW")
            vs, vc = validate(desc, e["desc"])
            return _res(desc, e["code"], e["desc"], sc, "NPC", mt, cf, vs, vc)

        if phc_idx:
            best2 = rfprocess.extractOne(gr, list(phc_idx.keys()), scorer=fuzz.token_sort_ratio)
            if best2 and best2[1] >= phc_t:
                e = phc_idx[best2[0]]; sc = best2[1]
                cf = "HIGH" if sc >= 80 else ("MEDIUM" if sc >= phc_t else "LOW")
                vs, vc = validate(desc, e["desc"])
                return _res(desc, e.get("npc_code",""), e["desc"], sc, "PHC",
                            "SPEC_DIFF", cf, vs, vc, ihbs=e.get("ihbs",""))

    fam = get_family(desc)
    code = f"500{SOP500.get(fam,'MISC')}NEW" if is_implant(desc) else f"200{SOP200.get(fam,'CONS')}NEW"
    return _res(desc, code, desc, 0, "UNMATCHED", "NEW_SOP", "LOW", "REVIEW",
                "No NPC/PHC match — new SOP code required (Sec 10.3/10.5)")

def build_npc_idx(df, d_col, c_col):
    idx = {}
    for _, r in df.iterrows():
        d, c = str(r[d_col]).strip(), str(r[c_col]).strip()
        idx[normalize(d)] = {"code": c, "desc": d}
        idx[generic_norm(d)] = {"code": c, "desc": d}
    return idx

def build_phc_idx(df, d_col, n_col, i_col):
    idx = {}
    for _, r in df.iterrows():
        d = str(r[d_col]).strip()
        idx[generic_norm(d)] = {
            "desc":     d,
            "npc_code": str(r[n_col]).strip() if n_col else "",
            "ihbs":     str(r[i_col]).strip() if i_col and i_col in r.index else "",
        }
    return idx

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    if not OPENPYXL_OK:
        return df.to_csv(index=False).encode()

    wb = Workbook()
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws   = wb.active
    ws.title = "Harmonized Mapping"
    cols = list(df.columns)
    ws.append(cols)

    hf = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.fill = hf
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bdr
    ws.row_dimensions[1].height = 38

    ci = {c: cols.index(c)+1 for c in ["MATCH_SOURCE","CONFIDENCE","VALIDATION_STATUS","ONLINE_CONFIDENCE"] if c in cols}
    src_f  = {"NPC": PatternFill("solid", fgColor="E2EFDA"), "PHC": PatternFill("solid", fgColor="DDEBF7"), "UNMATCHED": PatternFill("solid", fgColor="FFF2CC")}
    conf_f = {"HIGH": PatternFill("solid", fgColor="C6EFCE"), "MEDIUM": PatternFill("solid", fgColor="FFEB9C"), "LOW": PatternFill("solid", fgColor="FCE4D6")}
    val_f  = {"VALID": PatternFill("solid", fgColor="C6EFCE"), "REVIEW": PatternFill("solid", fgColor="FCE4D6")}

    for rv in df.itertuples(index=False):
        row = list(rv)
        ws.append(row)
        xl = ws[ws.max_row]
        src  = row[ci["MATCH_SOURCE"]-1]      if "MATCH_SOURCE"      in ci else ""
        conf = row[ci["CONFIDENCE"]-1]         if "CONFIDENCE"         in ci else ""
        vst  = row[ci["VALIDATION_STATUS"]-1]  if "VALIDATION_STATUS"  in ci else ""
        ocf  = row[ci["ONLINE_CONFIDENCE"]-1]  if "ONLINE_CONFIDENCE"  in ci else ""
        rf   = src_f.get(src, PatternFill("solid", fgColor="F5F5F5"))
        for cell in xl:
            cell.fill = rf; cell.font = Font(name="Calibri", size=9)
            cell.alignment = Alignment(vertical="center"); cell.border = bdr
        if "CONFIDENCE"        in ci: xl[ci["CONFIDENCE"]-1].fill        = conf_f.get(conf, rf)
        if "VALIDATION_STATUS" in ci: xl[ci["VALIDATION_STATUS"]-1].fill  = val_f.get(vst, rf)
        if "ONLINE_CONFIDENCE" in ci: xl[ci["ONLINE_CONFIDENCE"]-1].fill  = conf_f.get(ocf, rf)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(cols, 1):
        w = 65 if any(x in col for x in ["DESCRIPTION","DEFINITION","COMMENT","SNIPPET"]) else \
            18 if any(x in col for x in ["CODE","SOURCE","TYPE","VERIFIED"]) else \
            14 if any(x in col for x in ["SCORE","CLASS","CONFIDENCE"]) else 22
        ws.column_dimensions[get_column_letter(i)].width = w

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    total = len(df)
    sc  = df["MATCH_SOURCE"].value_counts()       if "MATCH_SOURCE"       in df.columns else pd.Series()
    cc  = df["CONFIDENCE"].value_counts()          if "CONFIDENCE"          in df.columns else pd.Series()
    vc  = df["VALIDATION_STATUS"].value_counts()   if "VALIDATION_STATUS"   in df.columns else pd.Series()
    enr = df["FDA_PRODUCT_CODE"].notna().sum()     if "FDA_PRODUCT_CODE"    in df.columns else 0
    sv  = df["DITTO_MATCH"].value_counts()          if "DITTO_MATCH"         in df.columns else pd.Series()
    rows = [
        ["IHS–NPC Harmonization Report  ·  v3.0 Online-Enhanced","","",""],
        [f"Generated: {datetime.now():%Y-%m-%d %H:%M}  |  Rwanda FDA  |  SOP ODDG/RES/SOP/004","","",""],
        ["","","",""],
        ["MATCH SOURCE","COUNT","% TOTAL","NOTES"],
        ["NPC (Primary)",  int(sc.get("NPC",0)),       f"{sc.get('NPC',0)/max(total,1)*100:.1f}%", "Direct NPC match"],
        ["PHC (Fallback)", int(sc.get("PHC",0)),       f"{sc.get('PHC',0)/max(total,1)*100:.1f}%", "PHC fallback"],
        ["UNMATCHED",      int(sc.get("UNMATCHED",0)), f"{sc.get('UNMATCHED',0)/max(total,1)*100:.1f}%", "New SOP code required"],
        ["TOTAL", total, "100%",""],
        ["","","",""],
        ["CONFIDENCE","COUNT","",""],
        ["HIGH", int(cc.get("HIGH",0)),"",""],["MEDIUM", int(cc.get("MEDIUM",0)),"",""],["LOW", int(cc.get("LOW",0)),"",""],
        ["","","",""],
        ["VALIDATION","COUNT","",""],
        ["VALID", int(vc.get("VALID",0)),"",""],["REVIEW", int(vc.get("REVIEW",0)),"","Requires review"],
        ["","","",""],
        ["ONLINE ENRICHMENT","","",""],
        ["FDA-enriched products", enr,"", "openFDA product codes found"],
        ["","","",""],
        ["DITTO ENTITY MATCHING","","",""],
        ["Match",      int(sv.get("MATCH",0)),     "","Strong semantic similarity"],
        ["Possible",   int(sv.get("POSSIBLE",0)),  "","Review recommended"],
        ["No Match",   int(sv.get("NO_MATCH",0)),  "","Likely mismatch — check manually"],
        ["Not checked",int(sv.get("NOT_CHECKED",0)),"",""],
    ]
    hf2 = PatternFill("solid", fgColor="1A5C1A")
    for i, rv in enumerate(rows, 1):
        ws2.append(rv)
        xl = ws2[i]
        if i<=2:
            for c in xl: c.fill=hf; c.font=Font(name="Calibri",bold=True,color="FFFFFF",size=11)
        elif rv[0] in ("MATCH SOURCE","CONFIDENCE","VALIDATION","ONLINE ENRICHMENT"):
            for c in xl: c.fill=hf2; c.font=Font(name="Calibri",bold=True,color="FFFFFF",size=10)
        else:
            for c in xl: c.font=Font(name="Calibri",size=10)
        for c in xl: c.alignment=Alignment(vertical="center")
    for col, w in zip("ABCD", [32,12,12,45]):
        ws2.column_dimensions[col].width = w

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p style="font-family:\'Syne\',sans-serif;font-size:1.1rem;font-weight:800;color:#00e676;letter-spacing:.05em;margin-bottom:.05rem;">🏥 IHS–NPC</p>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:.7rem;color:#8b949e;margin:0 0 .4rem;">Harmonizer · v3.0 Online · Rwanda FDA</p>', unsafe_allow_html=True)
    badge_cls = "ok" if REQUESTS_OK else "fl"
    badge_txt = "🟢 ONLINE" if REQUESTS_OK else "🔴 OFFLINE"
    st.markdown(f'<span class="ab {badge_cls}">{badge_txt}</span>', unsafe_allow_html=True)
    st.markdown("---")
    for i, name in enumerate(STEPS):
        if i < S.step:
            st.markdown(f'<div class="sbg d">✓ {i+1}</div><div style="font-family:\'Syne\',sans-serif;font-size:.83rem;color:#00b347;margin-bottom:.4rem;">{name}</div>', unsafe_allow_html=True)
        elif i == S.step:
            st.markdown(f'<div class="sbg a">→ {i+1}</div><div style="font-family:\'Syne\',sans-serif;font-size:.88rem;font-weight:700;color:#e6edf3;margin-bottom:.4rem;">{name}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="sbg">&nbsp;{i+1}</div><div style="font-size:.8rem;color:#484f58;margin-bottom:.4rem;">{name}</div>', unsafe_allow_html=True)
    st.markdown("---")
    with st.expander("🌐 API Status", expanded=False):
        if st.button("Check APIs", key="ping_btn", use_container_width=True):
            with st.spinner("Pinging..."):
                S.api_status = SearchEngine.ping_all()
        for name, ok in S.api_status.items():
            st.markdown(f'<div class="ab {"ok" if ok else "fl"}">{"✓" if ok else "✗"} {name}</div>', unsafe_allow_html=True)
    st.markdown("---")
    if st.button("↩ Reset", use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()
    st.markdown('<div style="font-size:.65rem;color:#484f58;margin-top:1.5rem;">SOP ODDG/RES/SOP/004<br>Rwanda FDA · v3.0</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# STEPS
# ─────────────────────────────────────────────────────────────────────────────
def step_upload():
    st.markdown('<h1 class="pt">Data Upload</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Upload source files. IHS, NPC, and PHC are required. RHIC is optional.</p>', unsafe_allow_html=True)

    for key, label, note, req, cols in [
        ("ihs",  "IHS File",  "Consumables to be mapped",           True,  "RHIC Code, Description"),
        ("npc",  "NPC File",  "National Product Catalogue",         True,  "NPC Code, Description"),
        ("phc",  "PHC File",  "PHC catalogue — fallback mapping",   True,  "Description, NPC Code, IHBS Code"),
        ("rhic", "RHIC File", "RHIC reference comparison",          False, "Description"),
    ]:
        loaded = key in S.files
        cls = "L" if loaded else ("r" if req else "o")
        badge = "🔴 REQUIRED" if req else "⚪ OPTIONAL"
        st.markdown(f'<div class="fc {cls}"><div class="fl">{label} &nbsp;<span style="font-size:.7rem;color:#8b949e;">{badge}</span></div><div class="fn">{note} · {cols}</div></div>', unsafe_allow_html=True)
        up = st.file_uploader(f"Upload {label}", type=["xlsx","xls","csv"], key=f"up_{key}", label_visibility="collapsed")
        if up:
            try:
                if up.name.endswith(".csv"):
                    df = pd.read_csv(up)
                else:
                    xl = pd.read_excel(up, sheet_name=None)
                    sheets = list(xl.keys())
                    chosen = st.selectbox(f"Sheet ({label})", sheets, key=f"sh_{key}") if len(sheets)>1 else sheets[0]
                    df = xl[chosen]
                df.columns = [str(c).strip() for c in df.columns]
                S.files[key] = df
                st.success(f"✓ {len(df):,} rows · {len(df.columns)} columns")
            except Exception as e:
                st.error(f"Error: {e}")
        elif loaded:
            st.info(f"✓ Previously loaded: {len(S.files[key]):,} rows")

    st.markdown("---")
    ready = all(k in S.files for k in ["ihs","npc","phc"])
    c1, c2 = st.columns([3,1])
    with c2:
        if st.button("Map Columns →", disabled=not ready, use_container_width=True):
            S.step = 1; st.rerun()
    if not ready:
        st.caption("⚠ Upload IHS, NPC, and PHC to continue")


def _sel(fkey, role, label, kws, hint=""):
    df = S.files.get(fkey)
    if df is None: return
    cols = list(df.columns)
    ranked = sorted(cols, key=lambda c: sum(k.upper() in c.upper() for k in kws), reverse=True)
    saved = S.col_map.get(fkey, {}).get(role)
    idx = ranked.index(saved) if saved in ranked else 0
    v = st.selectbox(label, ranked, index=idx, key=f"cm_{fkey}_{role}", help=hint)
    S.col_map.setdefault(fkey, {})[role] = v
    st.caption(f"→ {', '.join(str(x) for x in df[v].dropna().head(3).tolist())}")

def _sel_opt(fkey, role, label, kws):
    df = S.files.get(fkey)
    if df is None: return
    cols = ["(none)"] + sorted(list(df.columns), key=lambda c: sum(k.upper() in c.upper() for k in kws), reverse=True)
    saved = S.col_map.get(fkey, {}).get(role, "(none)")
    idx = cols.index(saved) if saved in cols else 0
    v = st.selectbox(label, cols, index=idx, key=f"cmo_{fkey}_{role}")
    S.col_map.setdefault(fkey, {})[role] = v if v != "(none)" else None
    if v != "(none)":
        df2 = S.files[fkey]
        st.caption(f"→ {', '.join(str(x) for x in df2[v].dropna().head(3).tolist())}")

def step_map():
    st.markdown('<h1 class="pt">Column Mapping</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Select which column in each file holds descriptions, codes, and identifiers.</p>', unsafe_allow_html=True)

    st.markdown('<div class="sh">IHS File</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: _sel("ihs", "desc", "Product Description", ["description","desc","product","consumable","item"])
    with c2: _sel_opt("ihs", "code", "RHIC Code (optional)", ["rhic","code","id"])

    st.markdown('<div class="sh">NPC File</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: _sel("npc", "code", "NPC Code", ["npc","code"])
    with c2: _sel("npc", "desc", "NPC Description", ["description","desc","product","rw"])

    st.markdown('<div class="sh">PHC File</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: _sel("phc", "desc",     "Description", ["description","desc","product"])
    with c2: _sel("phc", "npc_code", "NPC Code",    ["npc","code"])
    with c3: _sel_opt("phc", "ihbs", "IHBS Code",   ["ihbs","code"])

    if "rhic" in S.files:
        st.markdown('<div class="sh">RHIC File</div>', unsafe_allow_html=True)
        _sel("rhic", "desc", "Description", ["description","desc","product"])

    st.markdown("---")
    c1, _, c3 = st.columns([1,3,1])
    with c1:
        if st.button("← Back", use_container_width=True): S.step=0; st.rerun()
    with c3:
        ready = all([S.col_map.get("ihs",{}).get("desc"), S.col_map.get("npc",{}).get("code"),
                     S.col_map.get("npc",{}).get("desc"), S.col_map.get("phc",{}).get("desc"),
                     S.col_map.get("phc",{}).get("npc_code")])
        if st.button("Configure →", disabled=not ready, use_container_width=True):
            S.step=2; st.rerun()


def step_configure():
    st.markdown('<h1 class="pt">Configuration</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Set thresholds, enable online enrichment sources, and configure validation rules.</p>', unsafe_allow_html=True)

    c1, c2 = st.columns([3,2])
    with c1:
        st.markdown('<div class="sh">Match Thresholds</div>', unsafe_allow_html=True)
        npc_t = st.slider("NPC minimum score", 50, 100, S.config.get("npc_t", 65), help="PRD: ≥65")
        phc_t = st.slider("PHC fallback minimum score", 40, 95, S.config.get("phc_t", 60), help="PRD: ≥60")
        col_a, col_b = st.columns(2)
        with col_a: hi_t = st.number_input("HIGH confidence floor", 60, 100, S.config.get("hi_t", 80))
        with col_b: md_t = st.number_input("MEDIUM confidence floor", 40, 99, S.config.get("md_t", 65))

        st.markdown('<div class="sh">Brand Normalisation</div>', unsafe_allow_html=True)
        use_brand = st.toggle("Apply brand→generic aliases", value=S.config.get("use_brand", True))
        if use_brand:
            with st.expander(f"View {len(BRAND_ALIASES)} aliases"):
                st.dataframe(pd.DataFrame([{"Brand":k,"Generic":v} for k,v in BRAND_ALIASES.items()]),
                             use_container_width=True, height=180)
        max_rows = st.number_input("Max rows to process (0 = all)", 0, 50000, S.config.get("max_rows", 0))

    with c2:
        st.markdown('<div class="sh">🌐 Online Enrichment</div>', unsafe_allow_html=True)
        if not REQUESTS_OK:
            st.warning("`requests` not installed — online enrichment unavailable")
        use_online = st.toggle("Enable online enrichment", value=S.config.get("use_online", True) and REQUESTS_OK, disabled=not REQUESTS_OK)

        if use_online:
            st.markdown("**Sources to query:**")
            src_fda_c = st.checkbox("openFDA Classification",  value=S.config.get("src_fda_c", True),  help="api.fda.gov — free, no key, 1k req/day; 120k with key")
            src_fda_5 = st.checkbox("openFDA 510k DB",         value=S.config.get("src_fda_5", True),  help="Cleared device generic names")
            src_ddg   = st.checkbox("DuckDuckGo Instant API",  value=S.config.get("src_ddg", True),   help="api.duckduckgo.com — free, no key, context lookup")
            src_serp  = st.checkbox("SerpAPI (Google)",        value=S.config.get("src_serp", False),  help="Requires SERPAPI_KEY in st.secrets")

            enrich_scope = st.radio("Enrich which products?",
                                    ["UNMATCHED only","All rows"],
                                    index=["UNMATCHED only","All rows"].index(S.config.get("enrich_scope","UNMATCHED only")))
            max_enrich = st.number_input("Max enrichment calls (0 = unlimited)", 0, 10000, S.config.get("max_enrich", 500))

            st.markdown("""<div class="ec">
<b>🔑 API Keys (optional)</b><br>
Add to <code>~/.streamlit/secrets.toml</code>:<br>
<code>openfda_api_key = "YOUR_KEY"</code><br>
<code>SERPAPI_KEY = "YOUR_SERPAPI_KEY"</code><br>
<small style="color:#8b949e">Without keys: openFDA = 1,000 req/day · SerpAPI = disabled<br>
Verify &amp; Approve step uses <b>Ditto</b> (local, no key needed).</small>
</div>""", unsafe_allow_html=True)
        else:
            src_fda_c = src_fda_5 = src_ddg = src_serp = False
            enrich_scope = "UNMATCHED only"; max_enrich = 0

        st.markdown('<div class="sh">Validation Rules</div>', unsafe_allow_html=True)
        active_rules = []
        for k, lbl in [("anatomy","Anatomy incompatibility (AO Foundation)"),
                        ("family", "Product family mismatch"),
                        ("size",   "Numeric size tolerance (±0.1mm)"),
                        ("holes",  "Plate hole count check"),
                        ("units",  "Unit crossover ml vs mm")]:
            if st.checkbox(lbl, value=k in S.config.get("active_rules", ["anatomy","family","size","holes","units"]), key=f"r_{k}"):
                active_rules.append(k)

    S.config = dict(npc_t=npc_t, phc_t=phc_t, hi_t=hi_t, md_t=md_t,
                    use_brand=use_brand, max_rows=max_rows,
                    use_online=use_online, src_fda_c=src_fda_c,
                    src_fda_5=src_fda_5, src_ddg=src_ddg, src_serp=src_serp,
                    enrich_scope=enrich_scope, max_enrich=max_enrich,
                    active_rules=active_rules)
    st.markdown("---")
    c1, _, c3 = st.columns([1,3,1])
    with c1:
        if st.button("← Back", use_container_width=True): S.step=1; st.rerun()
    with c3:
        if st.button("Run →", use_container_width=True): S.step=3; st.rerun()


def step_process():
    st.markdown('<h1 class="pt">Processing & Online Enrichment</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Two-phase pipeline: fuzzy matching then online API enrichment for unresolved products.</p>', unsafe_allow_html=True)

    if S.run_done and S.results is not None:
        st.success(f"✓ Done — {len(S.results):,} rows")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("← Reconfigure", use_container_width=True): S.run_done=False; S.step=2; st.rerun()
        with c2:
            if st.button("View Results →", use_container_width=True): S.step=4; st.rerun()
        return

    cfg = S.config; cm = S.col_map
    df_ihs = S.files["ihs"]; df_npc = S.files["npc"]; df_phc = S.files["phc"]
    desc_col = cm["ihs"]["desc"]; code_col = cm["ihs"].get("code")
    npc_c = cm["npc"]["code"]; npc_d = cm["npc"]["desc"]
    phc_d = cm["phc"]["desc"]; phc_n = cm["phc"]["npc_code"]; phc_i = cm["phc"].get("ihbs")

    products = df_ihs.dropna(subset=[desc_col])
    if cfg["max_rows"] > 0: products = products.head(cfg["max_rows"])
    total = len(products)

    logs = []
    log_box = st.empty()
    def log(msg, kind="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        cls = {"ok":"co","warn":"cw","err":"ce","info":"ci","net":"cn"}.get(kind,"ci")
        logs.append(f'<span class="{cls}">[{ts}] {msg}</span>')
        log_box.markdown('<div class="cb">' + "<br>".join(logs[-35:]) + "</div>", unsafe_allow_html=True)

    log(f"Products: {total:,}  |  NPC threshold: {cfg['npc_t']}  |  PHC threshold: {cfg['phc_t']}")
    log("Building NPC index...")
    npc_idx = build_npc_idx(df_npc, npc_d, npc_c)
    log(f"NPC index ready: {len(npc_idx):,} entries", "ok")
    log("Building PHC index...")
    phc_idx = build_phc_idx(df_phc, phc_d, phc_n, phc_i)
    log(f"PHC index ready: {len(phc_idx):,} entries", "ok")

    # ── PHASE 1: MATCHING ─────────────────────────────────────────────────────
    st.markdown('<div class="sh">Phase 1 — Fuzzy Matching</div>', unsafe_allow_html=True)
    prog1 = st.progress(0, "Matching...")
    results = []
    for i, row in enumerate(products.itertuples(index=False)):
        raw = str(getattr(row, desc_col.replace(" ","_"), "")).strip()
        rhic = str(getattr(row, code_col.replace(" ","_"), "") if code_col else f"IHS-{i+1:04d}").strip()
        if not raw or raw.lower() == "nan": continue

        m = match_one(raw, npc_idx, phc_idx, cfg["npc_t"], cfg["phc_t"])
        sc = m["MATCH_SCORE"]
        m["CONFIDENCE"] = "HIGH" if sc >= cfg["hi_t"] else ("MEDIUM" if sc >= cfg["md_t"] else "LOW")
        m["RHIC_CODE"] = rhic
        results.append(m)
        pct = (i+1)/total
        prog1.progress(pct, f"Matching {i+1:,}/{total:,}")
        if i % 200 == 0 and i > 0:
            nm = sum(1 for r in results if r["MATCH_SOURCE"]!="UNMATCHED")
            log(f"  {i+1:,}/{total:,} · matched {nm:,} ({nm/len(results)*100:.0f}%)")

    prog1.progress(1.0, "Phase 1 complete ✓")
    n_npc = sum(1 for r in results if r["MATCH_SOURCE"]=="NPC")
    n_phc = sum(1 for r in results if r["MATCH_SOURCE"]=="PHC")
    n_unm = sum(1 for r in results if r["MATCH_SOURCE"]=="UNMATCHED")
    log(f"Matching done — NPC: {n_npc:,} · PHC: {n_phc:,} · Unmatched: {n_unm:,}", "ok")

    # ── PHASE 2: ONLINE ENRICHMENT ────────────────────────────────────────────
    if cfg.get("use_online") and REQUESTS_OK:
        st.markdown('<div class="sh">Phase 2 — Online Enrichment</div>', unsafe_allow_html=True)
        online_sources = []
        if cfg.get("src_fda_c"): online_sources.append("openfda_class")
        if cfg.get("src_fda_5"): online_sources.append("openfda_510k")
        if cfg.get("src_ddg"):   online_sources.append("duckduckgo")
        if cfg.get("src_serp"):  online_sources.append("serpapi")

        if online_sources:
            scope = cfg.get("enrich_scope", "UNMATCHED only")
            candidates = [r for r in results if scope == "All rows" or r["MATCH_SOURCE"]=="UNMATCHED"]
            max_e = cfg.get("max_enrich", 500)
            if max_e > 0: candidates = candidates[:max_e]

            unique_descs = list({r["ORIGINAL_DESCRIPTION"] for r in candidates})
            log(f"Enriching {len(unique_descs):,} unique products via: {', '.join(online_sources)}", "net")
            prog2 = st.progress(0, "Enriching online...")
            enr_map: dict = {}

            def do_enrich(desc):
                return desc, SearchEngine.enrich(desc, online_sources)

            with ThreadPoolExecutor(max_workers=3) as ex:
                futures = {ex.submit(do_enrich, d): d for d in unique_descs}
                done = 0
                for fut in as_completed(futures):
                    try:
                        desc_r, enr = fut.result()
                        enr_map[desc_r] = enr
                    except Exception:
                        pass
                    done += 1
                    prog2.progress(done/max(len(unique_descs),1), f"Enriching {done:,}/{len(unique_descs):,}")
                    if done % 20 == 0:
                        found = sum(1 for v in enr_map.values() if v.get("FDA_PRODUCT_CODE"))
                        log(f"  {done:,}/{len(unique_descs):,} done · {found:,} FDA codes found", "net")

            prog2.progress(1.0, "Phase 2 complete ✓")
            for r in results:
                if r["ORIGINAL_DESCRIPTION"] in enr_map:
                    r.update(enr_map[r["ORIGINAL_DESCRIPTION"]])
            n_fda = sum(1 for r in results if r.get("FDA_PRODUCT_CODE"))
            log(f"Enrichment done — {n_fda:,} FDA product codes found online", "ok")
        else:
            log("Online enabled but no sources selected", "warn")
    else:
        log("Offline mode — skipping online enrichment", "warn")

    # ── FINALISE ──────────────────────────────────────────────────────────────
    df_out = pd.DataFrame(results)
    front = ["RHIC_CODE","ORIGINAL_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION","IHBS_CODE",
             "MATCH_SOURCE","MATCH_SCORE","CONFIDENCE","MATCH_TYPE",
             "VALIDATION_STATUS","VALIDATION_COMMENT","PRODUCT_FAMILY"]
    online_cols = ["ONLINE_GENERIC_NAME","FDA_PRODUCT_CODE","FDA_DEVICE_CLASS",
                   "FDA_REGULATION","ONLINE_DEFINITION","ONLINE_SOURCE",
                   "ONLINE_CONFIDENCE","SUGGESTED_CATEGORY",
                   "DITTO_MATCH","DITTO_SCORE","DITTO_NOTE"]
    ordered = [c for c in front if c in df_out.columns] + [c for c in online_cols if c in df_out.columns]
    S.results = df_out[ordered]; S.run_done = True
    n_rev = (S.results["VALIDATION_STATUS"]=="REVIEW").sum() if "VALIDATION_STATUS" in S.results.columns else 0
    log(f"VALID: {len(results)-n_rev:,} · REVIEW: {n_rev:,}", "warn" if n_rev else "ok")
    log("Ready for review ✓", "ok")
    st.success(f"✓ {len(results):,} products processed")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("← Reconfigure", use_container_width=True): S.run_done=False; S.step=2; st.rerun()
    with c2:
        if st.button("View Results →", use_container_width=True): S.step=4; st.rerun()


def step_review():
    st.markdown('<h1 class="pt">Results Review</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Inspect, filter, and verify the harmonized mapping before export.</p>', unsafe_allow_html=True)

    df = S.results
    if df is None: st.warning("No results — run processing first."); return

    total = len(df)
    n_npc = (df["MATCH_SOURCE"]=="NPC").sum()     if "MATCH_SOURCE" in df.columns else 0
    n_phc = (df["MATCH_SOURCE"]=="PHC").sum()     if "MATCH_SOURCE" in df.columns else 0
    n_unm = (df["MATCH_SOURCE"]=="UNMATCHED").sum() if "MATCH_SOURCE" in df.columns else 0
    n_rev = (df["VALIDATION_STATUS"]=="REVIEW").sum() if "VALIDATION_STATUS" in df.columns else 0
    n_fda = df["FDA_PRODUCT_CODE"].notna().sum()  if "FDA_PRODUCT_CODE" in df.columns else 0
    cov   = (n_npc+n_phc)/total*100 if total else 0

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    for col, val, lbl, cls in [
        (k1, f"{total:,}",  "Total",       "km"),
        (k2, f"{n_npc:,}", "NPC Matched",  "kg"),
        (k3, f"{n_phc:,}", "PHC Fallback", "kb"),
        (k4, f"{n_unm:,}", "Unmatched",    "km"),
        (k5, f"{n_rev:,}", "Needs Review", "ky"),
        (k6, f"{cov:.0f}%","Coverage",     "kg" if cov>=80 else "ky"),
    ]:
        with col:
            st.markdown(f'<div class="kc"><div class="kv {cls}">{val}</div><div class="kl">{lbl}</div></div>', unsafe_allow_html=True)

    if n_fda > 0:
        st.markdown(f'<div style="margin:.7rem 0;font-size:.82rem;color:#ffc800;">🌐 {n_fda:,} products enriched with FDA codes from online APIs</div>', unsafe_allow_html=True)

    st.markdown('<div class="sh">Filters</div>', unsafe_allow_html=True)
    f1,f2,f3,f4,f5 = st.columns(5)
    with f1: fs = st.selectbox("Source",     ["All"]+sorted(df["MATCH_SOURCE"].dropna().unique().tolist()) if "MATCH_SOURCE" in df.columns else ["All"])
    with f2: fc = st.selectbox("Confidence", ["All"]+["HIGH","MEDIUM","LOW"])
    with f3: fv = st.selectbox("Validation", ["All","VALID","REVIEW"])
    with f4: ff = st.selectbox("Family",     ["All"]+sorted(df["PRODUCT_FAMILY"].dropna().unique().tolist()) if "PRODUCT_FAMILY" in df.columns else ["All"])
    with f5: ffda = st.selectbox("FDA Code", ["All","Has FDA Code","No FDA Code"]) if "FDA_PRODUCT_CODE" in df.columns else "All"

    view = df.copy()
    if fs  != "All" and "MATCH_SOURCE"      in view.columns: view = view[view["MATCH_SOURCE"]==fs]
    if fc  != "All" and "CONFIDENCE"        in view.columns: view = view[view["CONFIDENCE"]==fc]
    if fv  != "All" and "VALIDATION_STATUS" in view.columns: view = view[view["VALIDATION_STATUS"]==fv]
    if ff  != "All" and "PRODUCT_FAMILY"    in view.columns: view = view[view["PRODUCT_FAMILY"]==ff]
    if ffda == "Has FDA Code" and "FDA_PRODUCT_CODE" in view.columns:
        view = view[view["FDA_PRODUCT_CODE"].notna() & (view["FDA_PRODUCT_CODE"]!="")]
    if ffda == "No FDA Code" and "FDA_PRODUCT_CODE" in view.columns:
        view = view[view["FDA_PRODUCT_CODE"].isna() | (view["FDA_PRODUCT_CODE"]=="")]

    srch = st.text_input("🔍 Search", placeholder="e.g. locking screw 3.5mm")
    if srch:
        view = view[view.apply(lambda r: srch.upper() in " ".join(str(v).upper() for v in r.values), axis=1)]

    st.caption(f"Showing {len(view):,} / {total:,} rows")

    disp = ["RHIC_CODE","ORIGINAL_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION",
            "MATCH_SOURCE","MATCH_SCORE","CONFIDENCE","MATCH_TYPE",
            "VALIDATION_STATUS","FDA_PRODUCT_CODE","ONLINE_GENERIC_NAME","VALIDATION_COMMENT"]
    show = [c for c in disp if c in view.columns]

    def row_style(row):
        val = row.get("VALIDATION_STATUS","")
        src = row.get("MATCH_SOURCE","")
        if val=="REVIEW":    return ["background:rgba(248,81,73,.08)"]*len(row)
        if src=="UNMATCHED": return ["background:rgba(240,165,0,.07)"]*len(row)
        if src=="PHC":       return ["background:rgba(56,139,253,.05)"]*len(row)
        return ["background:rgba(0,179,71,.04)"]*len(row)

    st.dataframe(view[show].style.apply(row_style, axis=1).format({"MATCH_SCORE":"{:.0f}"}),
                 use_container_width=True, height=440)

    st.markdown("")
    ch1,ch2,ch3,ch4 = st.columns(4)
    for col, cn, color, lbl in [
        (ch1,"MATCH_SOURCE","#00b347","Match Source"),
        (ch2,"CONFIDENCE","#388bfd","Confidence"),
        (ch3,"PRODUCT_FAMILY","#f0a500","Product Family"),
        (ch4,"VALIDATION_STATUS","#e67676","Validation"),
    ]:
        with col:
            st.markdown(f'<div class="sh" style="font-size:.82rem;">{lbl}</div>', unsafe_allow_html=True)
            if cn in df.columns:
                vc = df[cn].value_counts().head(8).reset_index()
                vc.columns=[cn,"Count"]
                st.bar_chart(vc.set_index(cn), color=color, height=160)

    st.markdown("---")
    c1,_,c3 = st.columns([1,3,1])
    with c1:
        if st.button("← Back", use_container_width=True): S.step=3; st.rerun()
    with c3:
        if st.button("Verify & Approve →", use_container_width=True): S.step=5; st.rerun()


def step_verify():
    st.markdown('<h1 class="pt">Verify & Approve</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Ditto entity matching — pretrained language model scores semantic similarity between each original description and its matched NPC entry. Runs locally, no API key needed.</p>', unsafe_allow_html=True)

    df = S.results
    if df is None:
        st.warning("No results — run processing first.")
        return

    if not DittoMatcher.available():
        st.error("""**sentence-transformers not installed.**
Run `pip install sentence-transformers` then restart the app.
The Ditto matcher requires this package to load the pretrained language model.""")
        st.markdown("---")
        c1, _, c3 = st.columns([1,3,1])
        with c1:
            if st.button("← Back to Review", use_container_width=True): S.step=4; st.rerun()
        with c3:
            if st.button("Skip — Export anyway →", use_container_width=True): S.step=6; st.rerun()
        return

    # ── Identify flagged rows ─────────────────────────────────────────────────
    needs_check = (
        (df["VALIDATION_STATUS"] == "REVIEW") |
        (df["CONFIDENCE"]         == "LOW")   |
        (df["MATCH_SOURCE"]       == "UNMATCHED")
    ) if all(c in df.columns for c in ["VALIDATION_STATUS","CONFIDENCE","MATCH_SOURCE"]) \
      else pd.Series([True]*len(df))

    n_flag  = int(needs_check.sum())
    n_total = len(df)
    n_pass  = n_total - n_flag

    k1,k2,k3,k4 = st.columns(4)
    for col, val, lbl, cls in [
        (k1, f"{n_total:,}", "Total Rows",         "km"),
        (k2, f"{n_flag:,}",  "Flagged for Ditto",  "ky"),
        (k3, f"{n_pass:,}",  "High-Conf Pass",     "kg"),
        (k4, _DITTO_MODEL_NAME, "Model",            "kb"),
    ]:
        with col:
            st.markdown(f'<div class="kc"><div class="kv {cls}" style="font-size:{"1.1rem" if len(str(val))>10 else "2rem"}">{val}</div><div class="kl">{lbl}</div></div>',
                        unsafe_allow_html=True)

    st.markdown("---")

    # ── Run Ditto ──────────────────────────────────────────────────────────────
    if not S.verify_done:
        c_cfg, c_act = st.columns([3,2])
        with c_cfg:
            max_v = st.number_input(
                "Max flagged rows to score (0 = all)",
                0, 50000, min(1000, n_flag), key="max_ditto",
                help="Ditto processes ~200–500 pairs/sec on CPU.")
            t_match    = st.slider("MATCH threshold",    0.50, 0.99, DittoMatcher.T_MATCH,    0.01,
                                    help="Cosine similarity ≥ this → MATCH")
            t_possible = st.slider("POSSIBLE threshold", 0.30, 0.80, DittoMatcher.T_POSSIBLE, 0.01,
                                    help="Cosine similarity ≥ this → POSSIBLE (else NO_MATCH)")
            DittoMatcher.T_MATCH    = t_match
            DittoMatcher.T_POSSIBLE = t_possible
        with c_act:
            st.markdown("<br><br>", unsafe_allow_html=True)
            if st.button("Skip — export without Ditto scoring", use_container_width=True):
                S.verify_done = True
                st.rerun()

        if st.button("▶ Run Ditto Entity Matching", use_container_width=True, type="primary"):
            flagged = df[needs_check]
            batch   = flagged.head(max_v) if max_v > 0 else flagged

            # Deduplicate pairs — don't score the same pair twice
            seen, unique_pairs = set(), []
            for _, row in batch.iterrows():
                a = str(row.get("ORIGINAL_DESCRIPTION","")).strip()
                b = str(row.get("NPC_DESCRIPTION","")).strip()
                key = hashlib.md5((a+"|"+b).lower().encode()).hexdigest()
                if key not in seen and a and b:
                    seen.add(key)
                    unique_pairs.append((a, b, key))

            log_box  = st.empty()
            prog     = st.progress(0, "Loading Ditto model…")
            logs: list[str] = []
            def _log(msg: str, kind: str = "ci"):
                logs.append(f'<span class="{kind}">{msg}</span>')
                log_box.markdown('<div class="cb">' + "<br>".join(logs[-25:]) + "</div>",
                                 unsafe_allow_html=True)

            _log(f"Ditto model: <span class='cn'>{_DITTO_MODEL_NAME}</span>")
            _log(f"Unique pairs to score: <span class='co'>{len(unique_pairs):,}</span>")
            _log("Serialising records in COL/VAL format…")

            # Batch encode
            pairs_only = [(a, b) for a, b, _ in unique_pairs]
            prog.progress(0.05, "Encoding with sentence-transformer…")
            try:
                scores = DittoMatcher.batch_match(pairs_only)
            except Exception as e:
                st.error(f"Ditto error: {e}")
                return

            cache: dict = {}
            for (a, b, key_hash), result in zip(unique_pairs, scores):
                cache[key_hash] = result
                S.verify_cache[key_hash] = result

            prog.progress(0.9, "Stamping results…")
            match_col, score_col, note_col = [], [], []
            for _, row in S.results.iterrows():
                a = str(row.get("ORIGINAL_DESCRIPTION","")).strip()
                b = str(row.get("NPC_DESCRIPTION","")).strip()
                k = hashlib.md5((a+"|"+b).lower().encode()).hexdigest()
                hit = cache.get(k, S.verify_cache.get(k))
                if hit:
                    match_col.append(hit["match"])
                    score_col.append(round(hit["score"], 4))
                    note_col.append(hit["note"])
                else:
                    match_col.append("NOT_CHECKED")
                    score_col.append(None)
                    note_col.append("")

            S.results["DITTO_MATCH"] = match_col
            S.results["DITTO_SCORE"] = score_col
            S.results["DITTO_NOTE"]  = note_col
            S.verify_done = True

            n_match   = match_col.count("MATCH")
            n_poss    = match_col.count("POSSIBLE")
            n_nomatch = match_col.count("NO_MATCH")
            for m in scores[:8]:
                icon = {"MATCH":"✅","POSSIBLE":"🟡","NO_MATCH":"🔴"}.get(m["match"],"⚪")
                _log(f'{icon} <span class="co">{m["match"]}</span> · score={m["score"]:.3f} · {m["note"]}')
            prog.progress(1.0, "Done ✓")
            st.success(f"✓ Ditto done — MATCH: {n_match:,} · POSSIBLE: {n_poss:,} · NO_MATCH: {n_nomatch:,}")
            st.rerun()

    # ── Results table ─────────────────────────────────────────────────────────
    if S.verify_done:
        df_v = S.results
        if "DITTO_MATCH" in df_v.columns:
            mc = df_v["DITTO_MATCH"].value_counts()
            badges = [
                ('<span style="color:#00e676">✅ MATCH</span>',     int(mc.get("MATCH",0))),
                ('<span style="color:#f0a500">🟡 POSSIBLE</span>',  int(mc.get("POSSIBLE",0))),
                ('<span style="color:#f85149">🔴 NO_MATCH</span>',  int(mc.get("NO_MATCH",0))),
                ('<span style="color:#8b949e">⚪ NOT_CHECKED</span>',int(mc.get("NOT_CHECKED",0))),
            ]
            st.markdown("**Ditto results:** " +
                        " &nbsp;·&nbsp; ".join(f"{b} ({n:,})" for b, n in badges if n > 0),
                        unsafe_allow_html=True)

            # Score distribution bar chart
            scored = df_v[df_v["DITTO_SCORE"].notna()]
            if len(scored):
                bins = pd.cut(scored["DITTO_SCORE"], bins=[0,.4,.58,.75,.82,.90,.95,1.0],
                              labels=["<0.40","0.40–0.58","0.58–0.75","0.75–0.82",
                                      "0.82–0.90","0.90–0.95",">0.95"])
                dist = bins.value_counts().sort_index().reset_index()
                dist.columns = ["Similarity Band","Count"]
                st.markdown('<div class="sh">Score distribution</div>', unsafe_allow_html=True)
                st.bar_chart(dist.set_index("Similarity Band"), color="#388bfd", height=140)

        st.markdown('<div class="sh">Flagged rows</div>', unsafe_allow_html=True)
        show_cols = ["RHIC_CODE","ORIGINAL_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION",
                     "MATCH_SOURCE","CONFIDENCE","VALIDATION_STATUS"]
        if "DITTO_MATCH"  in df_v.columns: show_cols.append("DITTO_MATCH")
        if "DITTO_SCORE"  in df_v.columns: show_cols.append("DITTO_SCORE")
        if "DITTO_NOTE"   in df_v.columns: show_cols.append("DITTO_NOTE")
        show_cols = [c for c in show_cols if c in df_v.columns]

        filt_opts = ["Flagged only","NO_MATCH only","All rows"]
        filt = st.radio("Show", filt_opts, horizontal=True, key="ditto_filter")
        view = df_v.copy()
        if filt == "Flagged only":
            view = view[needs_check]
        elif filt == "NO_MATCH only" and "DITTO_MATCH" in view.columns:
            view = view[view["DITTO_MATCH"].isin(["NO_MATCH","NOT_CHECKED"])]

        def _vstyle(row):
            v = row.get("DITTO_MATCH","")
            if v == "MATCH":     return ["background:rgba(0,230,118,.07)"]*len(row)
            if v == "POSSIBLE":  return ["background:rgba(240,165,0,.07)"]*len(row)
            if v == "NO_MATCH":  return ["background:rgba(248,81,73,.09)"]*len(row)
            return ["background:rgba(139,148,158,.04)"]*len(row)

        fmt = {"DITTO_SCORE": "{:.3f}", "MATCH_SCORE": "{:.0f}"}
        st.dataframe(
            view[show_cols].style.apply(_vstyle, axis=1).format(
                {k: v for k, v in fmt.items() if k in show_cols}, na_rep="—"),
            use_container_width=True, height=420,
        )

    st.markdown("---")
    c1, _, c3 = st.columns([1,3,1])
    with c1:
        if st.button("← Back to Review", use_container_width=True): S.step=4; st.rerun()
    with c3:
        if st.button("Export →", use_container_width=True): S.step=6; st.rerun()


def step_export():
    st.markdown('<h1 class="pt">Export</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Download the complete harmonized and online-enriched dataset as Excel or CSV.</p>', unsafe_allow_html=True)

    df = S.results
    if df is None: st.warning("No results."); return

    total = len(df)
    n_npc = (df["MATCH_SOURCE"]=="NPC").sum()    if "MATCH_SOURCE" in df.columns else 0
    n_phc = (df["MATCH_SOURCE"]=="PHC").sum()    if "MATCH_SOURCE" in df.columns else 0
    n_rev = (df["VALIDATION_STATUS"]=="REVIEW").sum() if "VALIDATION_STATUS" in df.columns else 0
    n_fda = df["FDA_PRODUCT_CODE"].notna().sum() if "FDA_PRODUCT_CODE" in df.columns else 0
    cov   = (n_npc+n_phc)/total*100 if total else 0

    k1,k2,k3,k4 = st.columns(4)
    for col, val, lbl, cls in [
        (k1,f"{total:,}","Total Rows","km"),
        (k2,f"{cov:.0f}%","Coverage","kg" if cov>=80 else "ky"),
        (k3,f"{n_rev:,}","Needs Review","ky"),
        (k4,f"{n_fda:,}","FDA Enriched","kb"),
    ]:
        with col:
            st.markdown(f'<div class="kc"><div class="kv {cls}">{val}</div><div class="kl">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("")
    for line in [
        "📋 Sheet 1: Harmonized Mapping — all products with NPC code, IHBS code, confidence, validation, FDA enrichment",
        "📊 Sheet 2: Summary — source breakdown, confidence, validation, enrichment statistics",
    ]:
        st.markdown(f'<div style="padding:.4rem .7rem;border-left:2px solid var(--ac);margin:.3rem 0;font-size:.83rem;">{line}</div>', unsafe_allow_html=True)

    st.markdown("")
    with st.spinner("Building Excel workbook..."):
        xlsx = build_excel(df)

    fname = f"IHS_NPC_Harmonized_{datetime.now():%Y%m%d_%H%M}.xlsx"
    st.download_button("⬇ Download Harmonized Excel (.xlsx)",
                       data=xlsx, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.markdown("")
    st.download_button("⬇ Download CSV (raw data)",
                       data=df.to_csv(index=False).encode(),
                       file_name=fname.replace(".xlsx",".csv"),
                       mime="text/csv")

    st.markdown("---")
    st.markdown(f'<div style="font-size:.73rem;color:#484f58;text-align:center;">Generated {datetime.now():%Y-%m-%d %H:%M:%S} · Rwanda FDA · SOP ODDG/RES/SOP/004 · IHS–NPC Harmonizer v3.0 Online</div>', unsafe_allow_html=True)
    st.markdown("")
    if st.button("↩ New Session"):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────────────────
[step_upload, step_map, step_configure, step_process, step_review, step_verify, step_export][S.step]()
