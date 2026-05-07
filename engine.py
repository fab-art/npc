"""
MedMatch AI — Core Matching Engine
Hybrid: HF Inference API (semantic) + TF-IDF (fallback) + RapidFuzz (fuzzy) + Rule Validation
"""
import re, time, hashlib, warnings
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity as sk_cos
from rapidfuzz import fuzz, process as rfp
warnings.filterwarnings("ignore")

try:
    import requests as _req
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

# ─── Brand / synonym aliases ──────────────────────────────────────────────────
BRAND_ALIASES = {
    "PANADOL":"PARACETAMOL","TYLENOL":"PARACETAMOL","CALPOL":"PARACETAMOL",
    "DISPRIN":"ASPIRIN","CARDIPRIN":"ASPIRIN",
    "BRUFEN":"IBUPROFEN","NUROFEN":"IBUPROFEN","ADVIL":"IBUPROFEN",
    "AUGMENTIN":"AMOXICILLIN CLAVULANATE","CO-AMOXICLAV":"AMOXICILLIN CLAVULANATE",
    "SEPTRIN":"COTRIMOXAZOLE","BACTRIM":"COTRIMOXAZOLE",
    "FLAGYL":"METRONIDAZOLE","METROGYL":"METRONIDAZOLE",
    "ZITHROMAX":"AZITHROMYCIN","ZITROCIN":"AZITHROMYCIN",
    "CIPROBAY":"CIPROFLOXACIN","CIPROXIN":"CIPROFLOXACIN",
    "GENTOCIN":"GENTAMICIN","GARAMYCIN":"GENTAMICIN",
    "DIFLUCAN":"FLUCONAZOLE","CANESTEN":"CLOTRIMAZOLE",
    "ZOVIRAX":"ACYCLOVIR","FAMVIR":"FAMCICLOVIR",
    "LASIX":"FUROSEMIDE","FRUSEMIDE":"FUROSEMIDE",
    "TENORMIN":"ATENOLOL","LOPRESSOR":"METOPROLOL",
    "NORVASC":"AMLODIPINE","CARDIZEM":"DILTIAZEM",
    "LANOXIN":"DIGOXIN","ALDACTONE":"SPIRONOLACTONE",
    "GLUCOPHAGE":"METFORMIN","AMARYL":"GLIMEPIRIDE",
    "LANTUS":"INSULIN GLARGINE","ACTRAPID":"INSULIN REGULAR",
    "VENTOLIN":"SALBUTAMOL","BRICANYL":"TERBUTALINE",
    "BECOTIDE":"BECLOMETHASONE","PULMICORT":"BUDESONIDE",
    "ZANTAC":"RANITIDINE","LOSEC":"OMEPRAZOLE",
    "NEXIUM":"ESOMEPRAZOLE","PANTOLOC":"PANTOPRAZOLE",
    "IMODIUM":"LOPERAMIDE","MOTILIUM":"DOMPERIDONE",
    "MAXOLON":"METOCLOPRAMIDE","STEMETIL":"PROCHLORPERAZINE",
    "HALDOL":"HALOPERIDOL","LARGACTIL":"CHLORPROMAZINE",
    "VALIUM":"DIAZEPAM","ATIVAN":"LORAZEPAM",
    "EPILIM":"SODIUM VALPROATE","TEGRETOL":"CARBAMAZEPINE",
    "AMOXIL":"AMOXICILLIN","PENICILLIN V":"PHENOXYMETHYLPENICILLIN",
    "CECLOR":"CEFACLOR","FORTUM":"CEFTAZIDIME","ROCEPHIN":"CEFTRIAXONE",
    "ZINACEF":"CEFUROXIME","MERONEM":"MEROPENEM","TIENAM":"IMIPENEM CILASTATIN",
    "VIBRAMYCIN":"DOXYCYCLINE","SUMYCIN":"TETRACYCLINE",
    "CHLOROMYCETIN":"CHLORAMPHENICOL",
    # Surgical / device
    "CORTEX SCREW":"CORTICAL SCREW","CORTEX SCREWS":"CORTICAL SCREWS",
    "CONCELOUS":"CANCELLOUS","CONCEULOUS":"CANCELLOUS",
    "CAPROSYN":"POLYGLYCONATE SUTURE","VICRYL":"POLYGLACTIN 910 SUTURE",
    "MONOCRYL":"POLIGLECAPRONE 25 SUTURE","PROLENE":"POLYPROPYLENE SUTURE",
    "PDS":"POLYDIOXANONE SUTURE","SURGICEL":"OXIDIZED REGENERATED CELLULOSE",
    "POLYSORB":"POLYGLYCOLIC ACID LACTIDE SUTURE","BIOSYN":"GLYCOMER 631 SUTURE",
    "VELOSORB":"POLYGLYCOLIC ACID LACTIDE SUTURE","MAXON":"POLYGLYCOLATE SUTURE",
    "SURGIPRO":"POLYPROPYLENE SUTURE","TICRON":"COATED POLYESTER SUTURE",
    "HEMOLOCK":"HEM-O-LOK CLIP","LIGASURE":"VESSEL SEALING SYSTEM",
    "REDON":"CLOSED WOUND SUCTION DRAIN",
    # Consumables
    "ABBOCATH":"IV CANNULA","VENFLON":"IV CANNULA",
    "NELATON":"URINARY CATHETER","FOLEY":"URINARY CATHETER FOLEY",
    "RYLES TUBE":"NASOGASTRIC TUBE","NGT":"NASOGASTRIC TUBE",
    "PENROSE":"SOFT DRAIN PENROSE",
}

UNIT_STD = [
    (r"(\d+)\s*MG\b", r"\1MG"), (r"(\d+\.?\d*)\s*ML\b", r"\1ML"),
    (r"(\d+\.?\d*)\s*MM\b", r"\1MM"), (r"(\d+\.?\d*)\s*MCG\b", r"\1MCG"),
    (r"(\d+\.?\d*)\s*G\b(?!/)", r"\1G"), (r"(\d+\.?\d*)\s*IU\b", r"\1IU"),
    (r"(\d+\.?\d*)\s*%", r"\1PCT"), (r"(\d+),(\d+)", r"\1.\2"),
    (r"\b(?:CH|FG)\s*(\d+)", r"FR\1"), (r"\bFL\.?\s*OZ\b", "ML"),
    (r"\bTAB(S|LET)?\.?\b","TABLET"), (r"\bCAP(S|SULE)?\.?\b","CAPSULE"),
    (r"\bINJ\.?\b","INJECTION"), (r"\bSOL\.?\b","SOLUTION"),
    (r"\bSUSP\.?\b","SUSPENSION"), (r"\bINF\.?\b","INFUSION"),
    (r"\bOPH\.?\b","OPHTHALMIC"), (r"\bSYR\.?\b","SYRUP"),
    (r"\bSUPP\.?\b","SUPPOSITORY"), (r"\bOINT\.?\b","OINTMENT"),
    (r"\bCREAM\b","CREAM"), (r"\bGEL\b","GEL"),
]

PRODUCT_FAMILIES = {
    "TABLET":      r"\bTABLET\b|\bTABLETS\b|\bTABS?\b",
    "CAPSULE":     r"\bCAPSULE\b|\bCAPSULES\b|\bCAPS?\b",
    "INJECTION":   r"\bINJECTION\b|\bINJECTABLE\b|\bINJ\b|\bAMPOULE\b|\bVIAL\b",
    "SOLUTION":    r"\bSOLUTION\b|\bSOL\b|\bDROP\b|\bEYE DROP\b",
    "SUSPENSION":  r"\bSUSPENSION\b|\bSUSP\b|\bSYRUP\b|\bELIXIR\b",
    "INFUSION":    r"\bINFUSION\b|\bIV FLUID\b|\bIV BAG\b|\bDRIP\b",
    "CREAM_OINT":  r"\bCREAM\b|\bOINTMENT\b|\bGEL\b|\bLOTION\b|\bPASTE\b",
    "SUPPOSITORY": r"\bSUPPOSITORY\b|\bSUPP\b|\bPESSARY\b",
    "INHALER":     r"\bINHALER\b|\bINHALATION\b|\bNEBULIZ\b|\bMDI\b",
    "SCREW":       r"\bSCREW\b|\bCORTICAL\b|\bCANCELLOUS\b|\bLOCKING SCREW\b",
    "PLATE":       r"\bPLATE\b|\bLCP\b|\bDCP\b|\bDHS\b",
    "NAIL":        r"\bNAIL\b|\bPFNA\b|\bIMN\b",
    "SUTURE":      r"\bSUTURE\b|\bPOLYGLACTIN\b|\bPOLYGLYCOLIC\b|\bCATGUT\b",
    "CATHETER":    r"\bCATHETER\b|\bCANNULA\b(?!TED)|\bFOLEY\b|\bNELATON\b",
    "DRAIN":       r"\bDRAIN\b|\bREDON\b",
    "GLOVE":       r"\bGLOVE\b|\bGLOVES\b",
    "SYRINGE":     r"\bSYRINGE\b|\bSYRINGES\b",
    "BANDAGE":     r"\bBANDAGE\b|\bGAUZE\b|\bDRESSING\b",
    "MASK":        r"\bMASK\b|\bN95\b|\bSURGICAL MASK\b",
}

DOSE_FORMS_ORDER = ["TABLET","CAPSULE","INJECTION","SOLUTION","SUSPENSION","INFUSION","CREAM_OINT","SUPPOSITORY","INHALER"]
DEVICE_FORMS = ["SCREW","PLATE","NAIL","SUTURE","CATHETER","DRAIN","GLOVE","SYRINGE","BANDAGE","MASK"]

# ─── Normalization ─────────────────────────────────────────────────────────────
def normalize(text: str) -> str:
    if not text or pd.isna(text): return ""
    s = str(text).upper().strip()
    for pat, rep in UNIT_STD: s = re.sub(pat, rep, s)
    s = re.sub(r"[^\w\s%]", " ", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def brand_strip(text: str) -> str:
    s = normalize(text)
    for b, g in BRAND_ALIASES.items():
        s = re.sub(r"\b" + re.escape(b.upper()) + r"\b", g.upper(), s)
    return s

def get_family(text: str) -> str:
    u = normalize(text)
    for fam, pat in PRODUCT_FAMILIES.items():
        if re.search(pat, u): return fam
    return "OTHER"

def extract_strength(text: str) -> list:
    """Extract numeric specs: mg, ml, mm, %, IU etc."""
    return re.findall(r"\d+\.?\d*(?:MG|ML|MM|MCG|G|IU|PCT|FR)", normalize(text))

# ─── Embedding Engine ──────────────────────────────────────────────────────────
class EmbeddingEngine:
    """
    Primary:  HF Inference API  (all-MiniLM-L6-v2)
    Fallback: TF-IDF vectorizer (medical bigrams, local, no deps)
    """
    HF_EMBED_URL  = "https://api-inference.huggingface.co/pipeline/feature-extraction/sentence-transformers/all-MiniLM-L6-v2"
    HF_CROSS_URL  = "https://api-inference.huggingface.co/models/cross-encoder/ms-marco-MiniLM-L-6-v2"

    def __init__(self, hf_token: str = ""):
        self.token   = hf_token
        self.headers = {"Authorization": f"Bearer {hf_token}"} if hf_token else {}
        self.mode    = "tfidf"   # updated after first successful HF call
        self._tfidf: TfidfVectorizer | None = None
        self._tfidf_matrix = None
        self._catalog_texts: list[str] = []

    # ── HF API calls ────────────────────────────────────────────────────────
    def _hf_embed(self, texts: list[str]) -> np.ndarray | None:
        if not REQUESTS_OK or not self.token: return None
        try:
            r = _req.post(self.HF_EMBED_URL,
                          headers=self.headers,
                          json={"inputs": texts, "options": {"wait_for_model": True}},
                          timeout=15)
            if r.status_code == 200:
                arr = np.array(r.json(), dtype=np.float32)
                if arr.ndim == 3: arr = arr[:, 0, :]   # pooled output
                norms = np.linalg.norm(arr, axis=1, keepdims=True)
                return arr / np.maximum(norms, 1e-9)
        except Exception:
            pass
        return None

    def _hf_crossenc(self, query: str, candidates: list[str]) -> list[float] | None:
        if not REQUESTS_OK or not self.token: return None
        try:
            pairs = [[query, c] for c in candidates]
            r = _req.post(self.HF_CROSS_URL,
                          headers=self.headers,
                          json={"inputs": pairs, "options": {"wait_for_model": True}},
                          timeout=20)
            if r.status_code == 200:
                scores = r.json()
                if isinstance(scores, list):
                    return [float(s) if not isinstance(s, dict) else s.get("score", 0.0)
                            for s in scores]
        except Exception:
            pass
        return None

    # ── TF-IDF fallback ──────────────────────────────────────────────────────
    def _build_tfidf(self, texts: list[str]):
        self._catalog_texts = texts
        self._tfidf = TfidfVectorizer(
            ngram_range=(1, 3),
            min_df=1,
            analyzer="word",
            sublinear_tf=True,
        )
        self._tfidf_matrix = self._tfidf.fit_transform(texts)

    def _tfidf_query(self, text: str) -> np.ndarray:
        q = self._tfidf.transform([text])
        return sk_cos(q, self._tfidf_matrix).flatten()

    # ── Public API ───────────────────────────────────────────────────────────
    def index_catalog(self, texts: list[str]) -> np.ndarray | None:
        """Build index. Returns embedding matrix (HF) or None (TF-IDF)."""
        self._build_tfidf(texts)   # always build TF-IDF as safety net
        normed = [brand_strip(t) for t in texts]
        embs = self._hf_embed(normed)
        if embs is not None:
            self.mode = "hf"
            return embs
        self.mode = "tfidf"
        return None   # signal: use TF-IDF

    def query(self, text: str, catalog_embs: np.ndarray | None, top_k: int = 20) -> tuple[np.ndarray, np.ndarray]:
        """Return (top_k indices, scores) sorted descending."""
        normed = brand_strip(text)
        if self.mode == "hf" and catalog_embs is not None:
            q_emb = self._hf_embed([normed])
            if q_emb is not None:
                sims = (q_emb @ catalog_embs.T).flatten()
                idx  = np.argsort(sims)[::-1][:top_k]
                return idx, sims[idx]
        # TF-IDF fallback
        sims = self._tfidf_query(normed)
        idx  = np.argsort(sims)[::-1][:top_k]
        return idx, sims[idx]

    def rerank(self, query: str, candidates: list[str]) -> list[float]:
        """Cross-encoder re-ranking. Falls back to RapidFuzz token_set_ratio."""
        scores = self._hf_crossenc(query, candidates)
        if scores is not None:
            # Normalize to [0,1]
            arr = np.array(scores, dtype=float)
            mn, mx = arr.min(), arr.max()
            if mx > mn: arr = (arr - mn) / (mx - mn)
            return arr.tolist()
        # RapidFuzz fallback
        q = brand_strip(query)
        return [fuzz.token_set_ratio(q, brand_strip(c)) / 100.0 for c in candidates]

    @property
    def engine_name(self) -> str:
        return f"all-MiniLM-L6-v2 (HF API)" if self.mode == "hf" else "TF-IDF + Medical Bigrams (local)"


# ─── Validation Engine ─────────────────────────────────────────────────────────
ANATOMY_BAD = {
    "DISTAL RADIUS":    ["DHS","INTERTROCHANTERIC","HIP SCREW"],
    "VOLAR LOCKING":    ["DHS","HIP"],
    "CLAVICLE":         ["DHS","TIBIA","FEMUR","HIP"],
    "FILS DE CERCLAGE": ["TIGHTNER","TIGHTENER"],
    "CERCLAGE WIRE":    ["TIGHTNER","TIGHTENER"],
    "DRAIN DE REDON":   ["CHEST DRAIN","THORACIC","INTERCOSTAL"],
}

def validate(query: str, candidate: str) -> tuple[str, str, float]:
    """Returns (status, comment, penalty). Status = VALID | CAUTION | REJECT."""
    q, c = normalize(query), normalize(candidate)
    issues, penalty = [], 0.0

    # Family check
    qf, cf = get_family(query), get_family(candidate)
    if qf != "OTHER" and cf != "OTHER" and qf != cf:
        issues.append(f"Product type: query={qf}, candidate={cf}")
        penalty += 0.25

    # Anatomy incompatibility
    for kw, bad_list in ANATOMY_BAD.items():
        if kw.upper() in q:
            for bad in bad_list:
                if bad.upper() in c:
                    issues.append(f"Anatomy conflict: '{kw}' ↔ '{bad}'")
                    penalty += 0.4

    # Strength consistency
    q_str = set(extract_strength(query))
    c_str = set(extract_strength(candidate))
    if q_str and c_str:
        overlap = q_str & c_str
        if not overlap:
            issues.append(f"Strength mismatch: {q_str} vs {c_str}")
            penalty += 0.15
        elif q_str != c_str:
            issues.append(f"Strength differs: query={q_str}, candidate={c_str}")
            penalty += 0.08

    # Dose form for pharmaceuticals
    q_fam_ord = next((f for f in DOSE_FORMS_ORDER if re.search(PRODUCT_FAMILIES[f], q)), None)
    c_fam_ord = next((f for f in DOSE_FORMS_ORDER if re.search(PRODUCT_FAMILIES[f], c)), None)
    if q_fam_ord and c_fam_ord and q_fam_ord != c_fam_ord:
        issues.append(f"Dose form: {q_fam_ord} vs {c_fam_ord}")
        penalty += 0.20

    if penalty >= 0.4:  status = "REJECT"
    elif penalty > 0:   status = "CAUTION"
    else:               status = "VALID"

    return status, " | ".join(issues) if issues else "", min(penalty, 1.0)


# ─── Match-type classifier ─────────────────────────────────────────────────────
def classify_match(query: str, candidate: str, score: float) -> tuple[str, float, str]:
    """Returns (MATCH_TYPE, final_confidence, reason)."""
    q, c = brand_strip(query), brand_strip(candidate)
    qn, cn = normalize(query), normalize(candidate)

    # Exact after normalization
    if qn == cn: return "EXACT", 1.0, "Exact match after normalization"

    # Brand difference only
    if q == c and qn != cn: return "BRAND_MATCH", 0.97, "Same product — brand/terminology difference"

    # High semantic similarity + brand stripped match
    q_words = set(q.split()); c_words = set(c.split())
    if q_words == c_words: return "BRAND_MATCH", 0.95, "Equivalent after brand normalization"

    # Score-based classification
    if score >= 0.90: return "EXACT",       score, "Very high semantic similarity"
    if score >= 0.78: return "BRAND_MATCH", score, "High similarity — likely same product"
    if score >= 0.60: return "SPEC_DIFF",   score, "Same family — verify specifications"
    if score >= 0.40: return "SPEC_DIFF",   score, "Possible match — review required"
    return "NEW_SOP", score, "Low similarity — new code may be required"


# ─── Master pipeline ───────────────────────────────────────────────────────────
class MedMatchPipeline:
    def __init__(self, engine: EmbeddingEngine):
        self.engine      = engine
        self.catalog_df  = None        # full catalog DataFrame
        self.catalog_emb = None        # np.ndarray or None (TF-IDF mode)
        self.texts       = []          # normalized catalog texts for display

    def load_catalog(self, df: pd.DataFrame, desc_col: str, code_col: str,
                     progress_cb=None) -> str:
        """Index catalog. Returns status string."""
        df = df[[desc_col, code_col]].dropna().copy()
        df.columns = ["desc", "code"]
        df["desc_norm"] = df["desc"].apply(brand_strip)
        self.catalog_df  = df
        self.texts       = df["desc_norm"].tolist()
        if progress_cb: progress_cb(0.3, "Building embeddings...")
        self.catalog_emb = self.engine.index_catalog(self.texts)
        if progress_cb: progress_cb(1.0, "Done")
        return self.engine.engine_name

    def search(self, query: str, top_n: int = 5) -> list[dict]:
        """Run full pipeline. Returns top_n results as list of dicts."""
        if self.catalog_df is None: return []
        t0 = time.perf_counter()

        # Step 1: Retrieve top 20
        idx, raw_scores = self.engine.query(query, self.catalog_emb, top_k=20)
        candidates = [self.catalog_df.iloc[i] for i in idx]

        # Step 2: Re-rank
        cand_texts = [c["desc_norm"] for c in candidates]
        rerank_scores = self.engine.rerank(brand_strip(query), cand_texts)

        # Step 3: Hybrid score (70% semantic + 30% rerank)
        hybrid = [0.7 * float(raw_scores[i]) + 0.3 * float(rerank_scores[i])
                  for i in range(len(candidates))]

        # Step 4: Sort by hybrid
        order  = sorted(range(len(candidates)), key=lambda i: hybrid[i], reverse=True)
        results = []
        for rank, i in enumerate(order[:top_n]):
            cand   = candidates[i]
            score  = hybrid[i]
            mtype, conf, reason = classify_match(query, cand["desc"], score)
            val_status, val_comment, penalty = validate(query, cand["desc"])
            final_conf = max(0.0, conf - penalty)

            results.append({
                "rank":          rank + 1,
                "npc_code":      cand["code"],
                "candidate":     cand["desc"],
                "generic_name":  brand_strip(cand["desc"]),
                "match_type":    mtype,
                "raw_score":     round(float(raw_scores[i]), 4),
                "rerank_score":  round(float(rerank_scores[i]), 4),
                "hybrid_score":  round(score, 4),
                "confidence":    round(final_conf * 100, 1),
                "val_status":    val_status,
                "val_comment":   val_comment,
                "reason":        reason,
            })

        elapsed = round((time.perf_counter() - t0) * 1000, 1)
        for r in results: r["latency_ms"] = elapsed
        return results

    def batch_search(self, queries: list[str], progress_cb=None) -> pd.DataFrame:
        rows = []
        for i, q in enumerate(queries):
            hits = self.search(q, top_n=1)
            if hits:
                h = hits[0]; h["query"] = q
                rows.append(h)
            else:
                rows.append({"query": q, "npc_code":"","candidate":"","confidence":0,
                             "match_type":"NEW_SOP","val_status":"REVIEW","val_comment":"No match"})
            if progress_cb: progress_cb((i+1)/len(queries), f"{i+1}/{len(queries)}")
        return pd.DataFrame(rows)
