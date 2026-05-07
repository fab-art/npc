"""
MedMatch AI — Medical Standardization Engine v2.0
Rwanda FDA · SOP ODDG/RES/SOP/004
Pipeline: Input → Normalize → Embed → Retrieve → Re-rank → Validate → Enrich → Output
"""
import streamlit as st, pandas as pd, numpy as np
import re, io, time, hashlib, threading, warnings
from datetime import datetime
from collections import deque
warnings.filterwarnings("ignore")

try:
    from sentence_transformers import SentenceTransformer, CrossEncoder
    ST_OK = True
except ImportError:
    ST_OK = False

try:
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    REQ_OK = True
except ImportError:
    REQ_OK = False

try:
    from rapidfuzz import fuzz, process as rfproc
    RF_OK = True
except ImportError:
    RF_OK = False

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity as sklearn_cos
    SKL_OK = True
except ImportError:
    SKL_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPX_OK = True
except ImportError:
    OPX_OK = False

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="MedMatch AI", page_icon="🧬", layout="wide",
                   initial_sidebar_state="expanded")

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Syne:wght@400;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');
:root{--bg:#060d18;--sf:#0d1b2a;--sf2:#112236;--sf3:#162d44;--bd:#1e3a5f;--bd2:#2a4f7a;
  --ac:#00d4ff;--ac2:#7eeeff;--ac3:#00ffbb;--warn:#ffb300;--err:#ff4757;--ok:#00e676;
  --pur:#a78bfa;--pink:#f472b6;--tx:#e8f4fd;--mu:#6b8fa8;--mu2:#4a6b85;
  --mono:'JetBrains Mono',monospace;--head:'Syne',sans-serif;--body:'Inter',sans-serif;}
.stApp{background:var(--bg);color:var(--tx);font-family:var(--body);}
section[data-testid="stSidebar"]{background:var(--sf)!important;border-right:1px solid var(--bd);}
.stButton>button{background:linear-gradient(135deg,#0077aa,#00d4ff);color:#fff;
  font-family:var(--head);font-weight:700;border:none;border-radius:8px;padding:.55rem 1.6rem;
  letter-spacing:.03em;transition:all .2s;}
.stButton>button:hover{transform:translateY(-2px);box-shadow:0 6px 24px rgba(0,212,255,.35);}
.stButton>button:disabled{background:var(--sf3);color:var(--mu);box-shadow:none;transform:none;}
div[data-testid="stFileUploader"]{border:1.5px dashed var(--bd2);border-radius:10px;background:var(--sf);}
.stTextInput>div>div>input,.stTextArea>div>div>textarea{background:var(--sf2)!important;border:1px solid var(--bd)!important;border-radius:8px!important;color:var(--tx)!important;}
.stSelectbox>div>div{background:var(--sf2)!important;border:1px solid var(--bd)!important;border-radius:8px!important;}
.stProgress>div>div{background:linear-gradient(90deg,var(--ac),var(--ac3));}
hr{border-color:var(--bd);}
.stTabs [data-baseweb="tab-list"]{background:var(--sf);border-bottom:1px solid var(--bd);}
.stTabs [aria-selected="true"]{color:var(--ac)!important;border-bottom:2px solid var(--ac)!important;}
.logo{font-family:var(--head);font-size:1.5rem;font-weight:800;
  background:linear-gradient(135deg,var(--ac),var(--ac3));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.logo-sub{font-size:.65rem;color:var(--mu);letter-spacing:.12em;text-transform:uppercase;}
.pt{font-family:var(--head);font-size:2.1rem;font-weight:800;
  background:linear-gradient(135deg,var(--ac2),var(--ac3));-webkit-background-clip:text;
  -webkit-text-fill-color:transparent;letter-spacing:-.02em;}
.ps{color:var(--mu);font-size:.9rem;margin-top:.25rem;margin-bottom:1.5rem;}
.sh{font-family:var(--head);font-weight:700;font-size:1rem;border-left:3px solid var(--ac);
  padding-left:.7rem;margin:1.1rem 0 .6rem;color:var(--tx);}
.sc{background:var(--sf);border:1px solid var(--bd);border-radius:12px;padding:1.1rem 1.3rem;
  position:relative;overflow:hidden;}
.sc::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--ac),var(--ac3));}
.sv{font-family:var(--head);font-size:2rem;font-weight:800;}
.sl{font-size:.72rem;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;}
.sg{color:var(--ok);}.sb{color:var(--ac);}.sy{color:var(--warn);}.sp{color:var(--pur);}
.rc{background:var(--sf);border:1px solid var(--bd);border-radius:12px;padding:1.1rem 1.3rem;
  margin:.5rem 0;transition:border-color .2s;}
.rc:hover{border-color:var(--bd2);}
.rc.r1{border-color:rgba(0,212,255,.45);background:rgba(0,212,255,.04);}
.rc.r2{border-color:rgba(0,255,187,.25);}
.rn{font-family:var(--mono);font-size:.77rem;color:var(--ac);font-weight:700;letter-spacing:.06em;}
.rd{font-family:var(--head);font-size:.97rem;font-weight:700;color:var(--tx);margin:.2rem 0;}
.cbw{background:var(--sf3);border-radius:100px;height:6px;margin:.4rem 0;}
.cbar{height:6px;border-radius:100px;}
.ch{background:linear-gradient(90deg,#00e676,#00ffbb);}
.cm{background:linear-gradient(90deg,#ffb300,#ffd54f);}
.cl{background:linear-gradient(90deg,#ff4757,#ff6b81);}
.pill{display:inline-block;font-family:var(--mono);font-size:.65rem;font-weight:700;
  padding:2px 8px;border-radius:5px;margin:.1rem;}
.pE{background:rgba(0,230,118,.15);color:#00e676;border:1px solid rgba(0,230,118,.25);}
.pB{background:rgba(0,212,255,.15);color:var(--ac);border:1px solid rgba(0,212,255,.25);}
.pS{background:rgba(255,179,0,.15);color:#ffb300;border:1px solid rgba(255,179,0,.25);}
.pN{background:rgba(107,143,168,.1);color:var(--mu);border:1px solid rgba(107,143,168,.2);}
.pV{background:rgba(0,230,118,.12);color:#00e676;}
.pR{background:rgba(255,71,87,.12);color:#ff4757;}
.pAI{background:rgba(167,139,250,.15);color:var(--pur);border:1px solid rgba(167,139,250,.25);}
.mb{display:inline-flex;align-items:center;gap:.3rem;padding:.2rem .6rem;border-radius:6px;
  font-family:var(--mono);font-size:.67rem;font-weight:700;}
.mbok{background:rgba(0,230,118,.1);color:#00e676;border:1px solid rgba(0,230,118,.25);}
.mbwn{background:rgba(255,179,0,.1);color:#ffb300;border:1px solid rgba(255,179,0,.25);}
.mbof{background:rgba(107,143,168,.1);color:var(--mu);border:1px solid rgba(107,143,168,.2);}
.pipe{display:inline-flex;flex-direction:column;align-items:center;gap:.15rem;padding:.45rem .7rem;
  border:1px solid var(--bd);border-radius:8px;font-size:.65rem;font-family:var(--mono);
  color:var(--mu);background:var(--sf);min-width:75px;text-align:center;}
.pipe.on{border-color:var(--ac);color:var(--ac);background:rgba(0,212,255,.06);}
.pipe.dn{border-color:var(--ok);color:var(--ok);background:rgba(0,230,118,.05);}
.cons{background:#020810;border:1px solid var(--bd);border-radius:10px;padding:1rem 1.2rem;
  font-family:var(--mono);font-size:.72rem;max-height:260px;overflow-y:auto;line-height:1.9;}
.co{color:var(--ok);}.cw{color:var(--warn);}.ce{color:var(--err);}
.cai{color:var(--pur);}.cn{color:var(--ac);}.ci{color:#58a6ff;}
.fda{background:rgba(0,212,255,.04);border:1px solid rgba(0,212,255,.2);
  border-radius:8px;padding:.8rem 1rem;margin:.4rem 0;font-size:.82rem;}
.ff{color:var(--mu);font-size:.7rem;text-transform:uppercase;letter-spacing:.07em;}
.fv{font-family:var(--mono);font-size:.8rem;color:var(--ac2);margin-bottom:.3rem;}
.hi{background:var(--sf2);border:1px solid var(--bd);border-radius:7px;padding:.45rem .7rem;
  margin:.25rem 0;font-size:.78rem;color:var(--mu);}
</style>""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
def _init():
    for k,v in dict(page="search",catalog_df=None,embeddings=None,embed_meta=None,
                    tfidf_mat=None,tfidf_vec=None,history=deque(maxlen=50),
                    batch_results=None,model_name=None,enrich_cache={},
                    stats=dict(queries=0,exact=0,brand=0,spec=0,new=0,avg_ms=0)).items():
        if k not in st.session_state: st.session_state[k]=v
_init()
S=st.session_state

# ── KNOWLEDGE BASES ───────────────────────────────────────────────────────────
BRAND = {
    "CORTEX SCREW":"CORTICAL BONE SCREW","CONCELOUS":"CANCELLOUS","CANSELLOUS":"CANCELLOUS",
    "CAPROSYN":"POLYGLYTONE ABSORBABLE MONOFILAMENT SUTURE",
    "VICRYL":"POLYGLACTIN 910 ABSORBABLE BRAIDED SUTURE",
    "MONOCRYL":"POLIGLECAPRONE ABSORBABLE MONOFILAMENT SUTURE",
    "PROLENE":"POLYPROPYLENE NON-ABSORBABLE SUTURE","PDS":"POLYDIOXANONE SUTURE",
    "POLYSORB":"POLYGLYCOLIC ACID LACTIDE ABSORBABLE BRAIDED SUTURE",
    "BIOSYN":"GLYCOMER ABSORBABLE MONOFILAMENT SUTURE",
    "VELOSORB":"POLYGLYCOLIC ACID LACTIDE ABSORBABLE BRAIDED SUTURE",
    "MAXON":"POLYGLYCOLATE ABSORBABLE MONOFILAMENT SUTURE",
    "SURGIPRO":"POLYPROPYLENE NON-ABSORBABLE MONOFILAMENT SUTURE",
    "TICRON":"COATED POLYESTER NON-ABSORBABLE BRAIDED SUTURE",
    "SURGIDAC":"POLYESTER NON-ABSORBABLE BRAIDED SUTURE",
    "SOFSILK":"SILK NON-ABSORBABLE BRAIDED SUTURE",
    "SURGILON":"BRAIDED NYLON NON-ABSORBABLE SUTURE",
    "MONOSOF":"NYLON MONOFILAMENT NON-ABSORBABLE SUTURE",
    "SURGICEL":"OXIDIZED REGENERATED CELLULOSE HAEMOSTATIC",
    "HEMOLOCK":"HEM-O-LOK POLYMER LIGATION CLIP",
    "LIGASURE":"VESSEL SEALING ELECTROTHERMAL SYSTEM",
    "REDON":"CLOSED WOUND SUCTION DRAIN",
    "VERSAPORT":"LAPAROSCOPIC RADIALLY EXPANDING TROCAR",
    "SONICISION":"CORDLESS ULTRASONIC DISSECTOR",
    "PANADOL":"PARACETAMOL ANALGESIC ANTIPYRETIC",
    "AUGMENTIN":"AMOXICILLIN CLAVULANIC ACID ANTIBIOTIC",
    "FLAGYL":"METRONIDAZOLE ANTIBIOTIC","BRUFEN":"IBUPROFEN NSAID",
    "VENTOLIN":"SALBUTAMOL BRONCHODILATOR","LASIX":"FUROSEMIDE DIURETIC",
    "ZITHROMAX":"AZITHROMYCIN ANTIBIOTIC","CIPROBAY":"CIPROFLOXACIN ANTIBIOTIC",
    "SEPTRIN":"TRIMETHOPRIM SULFAMETHOXAZOLE ANTIBIOTIC",
}

FAMILIES={
    "SCREW":    r"\bSCREW\b|\bCORTEX\b|\bCORTICAL\b|\bCANCELLOUS\b|\bINTERLOCKING\b",
    "PLATE":    r"\bPLATE\b|\bLCP\b|\bDCP\b|\bDHS\b",
    "NAIL":     r"\bNAIL\b|\bPFNA\b|\bIMN\b",
    "KWIRE":    r"\bKIRSCHNER\b|\bK-WIRE\b",
    "SUTURE":   r"\bSUTURE\b|POLYSORB|BIOSYN|CAPROSYN|SURGIPRO|TICRON|VELOSORB|MONOCRYL|VICRYL|PROLENE|POLYGLACTIN|POLYDIOXANONE|CHROMIC GUT|PLAIN GUT",
    "TROCAR":   r"\bTROCAR\b|\bCANNULA\b(?!TED)|\bVERSAPORT\b",
    "STAPLER":  r"\bSTAPLER\b|\bGIA\b|\bEEA\b|\bCEEA\b",
    "DRAIN":    r"\bDRAIN\b|\bREDON\b",
    "DRUG":     r"\bTABLET\b|\bCAPSULE\b|\bSYRUP\b|\bINJECTION\b|\bMG\b.*\bTAB\b|\bVIAL\b|\bAMPOULE\b",
    "CATHETER": r"\bCATHETER\b|\bFOLEY\b",
    "SYRINGE":  r"\bSYRINGE\b","NEEDLE":r"\bNEEDLE\b",
    "GLOVE":    r"\bGLOVE\b","BANDAGE":r"\bBANDAGE\b|\bGAUZE\b|\bDRESSING\b",
    "MESH":     r"\bMESH\b",
}

ANAT_INCOMPAT={
    "DISTAL RADIUS":["DHS","HIP SCREW"],"CLAVICLE":["DHS","TIBIA","HIP"],
    "PROXIMAL HUMERUS":["DHS","HIP"],"FILS DE CERCLAGE":["TIGHTNER","TIGHTENER"],
}

# ── MODULE 1: NORMALIZATION ────────────────────────────────────────────────────
class NE:
    @staticmethod
    def clean(t):
        if pd.isna(t): return ""
        s=str(t).upper().strip()
        s=re.sub(r"(\d+),(\d+)",r"\1.\2",s)
        s=re.sub(r"[×✕*]"," ",s)
        s=re.sub(r"\.0\s*MM","MM",s)
        s=re.sub(r"(\d+\.?\d*)\s*MM",r"\1MM",s)
        s=re.sub(r"(\d+\.?\d*)\s*ML",r"\1ML",s)
        s=re.sub(r"(\d+\.?\d*)\s*MG",r"\1MG",s)
        s=re.sub(r"\b(?:CH|FG)\s*(\d+)",r"FR\1",s)
        s=re.sub(r"\bS\.T\.\b|\bSELF-TAPPING\b","SELF TAPPING",s)
        s=re.sub(r"\bLH\b","LEFT",s); s=re.sub(r"\bRH\b","RIGHT",s)
        return re.sub(r"\s{2,}"," ",s).strip(" ,.-")

    @staticmethod
    def generic(t):
        s=NE.clean(t)
        for b,g in BRAND.items():
            s=re.sub(r"\b"+re.escape(b.upper())+r"\b",g,s)
        return s

    @staticmethod
    def for_embed(t): return NE.generic(t).lower()

    @staticmethod
    def specs(t):
        u=NE.clean(t)
        return {"dims":sorted([float(v) for v in re.findall(r"(\d+\.?\d*)MM",u)]),
                "ml":[float(v) for v in re.findall(r"(\d+\.?\d*)ML",u)],
                "mg":[float(v) for v in re.findall(r"(\d+\.?\d*)MG",u)],
                "fr":[int(v) for v in re.findall(r"FR(\d+)",u)],
                "holes":int(m.group(1)) if (m:=re.search(r"(\d+)\s*HOLES?",u)) else None}

    @staticmethod
    def family(t):
        u=NE.clean(t)
        for fam,pat in FAMILIES.items():
            if re.search(pat,u): return fam
        return "OTHER"

# ── MODULE 2: EMBEDDING ENGINE ─────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def load_models():
    if not ST_OK: return None,None
    try:
        bi=SentenceTransformer("all-MiniLM-L6-v2")
        try: ce=CrossEncoder("cross-encoder/ms-marco-MiniLM-L-6-v2")
        except Exception: ce=None
        return bi,ce
    except Exception: return None,None

class EmbedEngine:
    HF_URL="https://api-inference.huggingface.co/models/sentence-transformers/all-MiniLM-L6-v2"
    def __init__(self):
        self.bi,self.ce=load_models()
        if self.bi: S.model_name="all-MiniLM-L6-v2"
        elif SKL_OK: S.model_name="TF-IDF"
        else: S.model_name="RapidFuzz"

    def embed(self,texts):
        if self.bi:
            return self.bi.encode(texts,normalize_embeddings=True,show_progress_bar=False,batch_size=64)
        if REQ_OK:
            tok=""
            try: tok=st.secrets.get("HF_TOKEN","")
            except: pass
            if tok:
                try:
                    r=requests.post(self.HF_URL,headers={"Authorization":f"Bearer {tok}"},
                                    json={"inputs":texts,"options":{"wait_for_model":True}},timeout=20)
                    if r.status_code==200:
                        arr=np.array(r.json())
                        nrm=np.linalg.norm(arr,axis=1,keepdims=True)
                        return arr/np.maximum(nrm,1e-9)
                except: pass
        return None

    def rerank(self,query,candidates):
        if self.ce is None or len(candidates)<=1: return candidates
        try:
            pairs=[[query,c[1]["desc"]] for c in candidates]
            scores=self.ce.predict(pairs)
            ranked=sorted([(float(scores[i]),c[1]) for i,c in enumerate(candidates)],key=lambda x:-x[0])
            mx=max(s for s,_ in ranked) or 1; mn=min(s for s,_ in ranked); rng=mx-mn or 1
            return [((s-mn)/rng,m) for s,m in ranked]
        except: return candidates

EE=EmbedEngine()

# ── CATALOG INDEXER ────────────────────────────────────────────────────────────
def build_index(df,d_col,c_col,cat_col=None):
    records=[]; texts=[]
    for _,row in df.iterrows():
        desc=str(row[d_col]).strip(); code=str(row[c_col]).strip()
        if not desc or desc.lower()=="nan": continue
        cat=str(row[cat_col]).strip() if cat_col and cat_col in row.index else NE.family(desc)
        g=NE.for_embed(desc)
        records.append({"code":code,"desc":desc,"generic":g,"category":cat,"family":NE.family(desc)})
        texts.append(g)
    if not records: return False
    S.embed_meta=records
    embs=EE.embed(texts)
    if embs is not None: S.embeddings=embs.astype(np.float32)
    if SKL_OK:
        vec=TfidfVectorizer(ngram_range=(1,2),min_df=1)
        S.tfidf_mat=vec.fit_transform(texts); S.tfidf_vec=vec
    return True

# ── RETRIEVAL ──────────────────────────────────────────────────────────────────
def retrieve(query_generic,top_k=20):
    if not S.embed_meta: return []
    if S.embeddings is not None:
        qv=EE.embed([query_generic])
        if qv is not None:
            sims=np.dot(S.embeddings,qv[0]).astype(float)
            idx=np.argsort(sims)[::-1][:top_k]
            return [(float(sims[i]),S.embed_meta[i]) for i in idx]
    if S.tfidf_vec is not None:
        qv=S.tfidf_vec.transform([query_generic])
        sims=sklearn_cos(qv,S.tfidf_mat).flatten()
        idx=np.argsort(sims)[::-1][:top_k]
        return [(float(sims[i]),S.embed_meta[i]) for i in idx]
    if RF_OK:
        choices=[m["generic"] for m in S.embed_meta]
        hits=rfproc.extract(query_generic,choices,scorer=fuzz.token_sort_ratio,limit=top_k)
        return [(s/100.0,S.embed_meta[choices.index(t)]) for t,s,_ in hits]
    return []

# ── MODULE 5: VALIDATION ───────────────────────────────────────────────────────
def validate(inp,npc):
    issues=[]
    ui,ni=NE.clean(inp),NE.clean(npc)
    for kw,bads in ANAT_INCOMPAT.items():
        if kw.upper() in ui:
            for b in bads:
                if b.upper() in ni: issues.append(f"Anatomy: '{kw}' ≠ '{b}'")
    fi,fn=NE.family(inp),NE.family(npc)
    if fi!="OTHER" and fn!="OTHER" and fi!=fn: issues.append(f"Family: {fi} vs {fn}")
    si,sn=NE.specs(inp),NE.specs(npc)
    if si["dims"] and sn["dims"] and abs(si["dims"][0]-sn["dims"][0])>=0.15:
        issues.append(f"Size: {si['dims'][0]}mm vs {sn['dims'][0]}mm")
    if si["ml"] and sn["ml"] and abs(si["ml"][0]-sn["ml"][0])>0.5:
        issues.append(f"Volume: {si['ml'][0]}ml vs {sn['ml'][0]}ml")
    if si["mg"] and sn["mg"] and si["mg"][0]!=sn["mg"][0]:
        issues.append(f"CRITICAL dose: {si['mg'][0]}mg vs {sn['mg'][0]}mg")
    if si["fr"] and sn["fr"] and si["fr"][0]!=sn["fr"][0]:
        issues.append(f"FR size: {si['fr'][0]} vs {sn['fr'][0]}")
    return ("REVIEW",issues) if issues else ("VALID",[])

def match_type(inp,npc,score):
    ig,ng=NE.generic(inp),NE.generic(npc)
    ic,nc=NE.clean(inp),NE.clean(npc)
    if ic==nc or ig==ng: return "EXACT"
    if score>=0.97: return "EXACT"
    if score>=0.88: return "BRAND_MATCH"
    if NE.family(inp)==NE.family(npc) and NE.family(inp)!="OTHER" and score>=0.65: return "SPEC_DIFF"
    return "NEW_SOP"

# ── MODULE 6: ENRICHMENT ───────────────────────────────────────────────────────
_hts=None; _rlock=threading.Lock(); _lreq={}

def _get(url,params=None,host="x",gap=0.7,tout=10):
    global _hts
    if not REQ_OK: return None
    if _hts is None:
        s=requests.Session()
        s.mount("https://",HTTPAdapter(max_retries=Retry(total=2,backoff_factor=0.4,status_forcelist=[429,500,502,503])))
        s.headers["User-Agent"]="MedMatchAI/2.0 Rwanda-FDA"
        _hts=s
    with _rlock:
        w=gap-(time.time()-_lreq.get(host,0))
        if w>0: time.sleep(w)
        _lreq[host]=time.time()
    try:
        r=_hts.get(url,params=params,timeout=tout)
        return r.json() if r.status_code==200 else ({} if r.status_code==404 else None)
    except: return None

def _fda_key():
    try: return st.secrets.get("OPENFDA_KEY","")
    except: return ""

def enrich(description):
    ck=hashlib.md5(description.lower().encode()).hexdigest()
    if ck in S.enrich_cache: return S.enrich_cache[ck]
    clean=re.sub(r"\d+\.?\d*\s*(?:mm|ml|mg|mcg|iu|fr|g)\b","",description,flags=re.I).strip()
    words=[w for w in clean.upper().split() if len(w)>2]; q=" ".join(words[:6])
    key=_fda_key(); fam=NE.family(description)
    result={}
    if fam=="DRUG":
        params={"search":f'brand_name:"{q}"',"limit":1}
        if key: params["api_key"]=key
        d=_get("https://api.fda.gov/drug/label.json",params,"fda_drug")
        if d and d.get("results"):
            top=d["results"][0]; ofd=top.get("openfda",{})
            result={"source":"openFDA Drug Label",
                    "generic_name":ofd.get("generic_name",[""])[0],
                    "product_type":ofd.get("product_type",[""])[0],
                    "manufacturer":ofd.get("manufacturer_name",[""])[0]}
    else:
        for q2 in [f'device_name:"{q}"',f"device_name:{chr(43).join(q.split()[:3])}"]:
            params={"search":q2,"limit":3}
            if key: params["api_key"]=key
            d=_get("https://api.fda.gov/device/classification.json",params,"fda_dev")
            if d and d.get("results"):
                top=d["results"][0]
                result={"source":"openFDA Classification",
                        "device_name":top.get("device_name",""),
                        "product_code":top.get("product_code",""),
                        "device_class":f"Class {top.get('device_class','')}",
                        "definition":top.get("definition","")[:250],
                        "regulation":top.get("regulation_number","")}
                break
    S.enrich_cache[ck]=result; return result

# ── MEDMATCH PIPELINE ──────────────────────────────────────────────────────────
def medmatch(description,top_n=5,do_enrich=False):
    if not S.embed_meta: return {"error":"No catalog loaded."}
    t0=time.perf_counter()
    generic=NE.for_embed(description)
    specs=NE.specs(description); family=NE.family(description)
    candidates=retrieve(generic,top_k=20)
    if not candidates: return {"error":"No candidates.","results":[]}
    reranked=EE.rerank(generic,candidates)
    top=reranked[:top_n]
    results_out=[]
    for rank,(score,meta) in enumerate(top,1):
        vs,vi=validate(description,meta["desc"])
        mt=match_type(description,meta["desc"],score)
        fuzz_s=(fuzz.token_sort_ratio(NE.clean(description),NE.clean(meta["desc"]))/100.0) if RF_OK else score
        hybrid=round((0.7*score+0.3*fuzz_s)*100,1)
        results_out.append({"rank":rank,"npc_code":meta["code"],"npc_desc":meta["desc"],
                            "category":meta["category"],"family":meta["family"],
                            "score":hybrid,"match_type":mt,"val_status":vs,"val_issues":vi})
    ed={} 
    if do_enrich and REQ_OK: ed=enrich(description)
    ms=round((time.perf_counter()-t0)*1000,1)
    S.stats["queries"]+=1; mt0=results_out[0]["match_type"] if results_out else "NEW_SOP"
    S.stats[{"EXACT":"exact","BRAND_MATCH":"brand","SPEC_DIFF":"spec"}.get(mt0,"new")]+=1
    n=S.stats["queries"]
    S.stats["avg_ms"]=round((S.stats["avg_ms"]*(n-1)+ms)/n,1)
    return {"input":description,"normalized":generic,"family":family,"specs":specs,
            "results":results_out,"elapsed_ms":ms,"engine":S.model_name or "?","enrich":ed}

# ── EXPORT ─────────────────────────────────────────────────────────────────────
def to_excel(rows):
    if not OPX_OK: return pd.DataFrame(rows).to_csv(index=False).encode()
    wb=Workbook(); ws=wb.active; ws.title="MedMatch Results"
    thin=Side(style="thin",color="1e3a5f"); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    cols=["INPUT_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION","CATEGORY","MATCH_SCORE",
          "MATCH_TYPE","VALIDATION_STATUS","VALIDATION_COMMENT","FDA_PRODUCT_CODE","FDA_DEFINITION"]
    ws.append(cols)
    hf=PatternFill("solid",fgColor="060d18")
    for c in ws[1]:
        c.fill=hf; c.font=Font(name="Calibri",bold=True,color="00D4FF",size=10)
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=bdr
    ws.row_dimensions[1].height=34
    mtf={"EXACT":PatternFill("solid",fgColor="063a1f"),"BRAND_MATCH":PatternFill("solid",fgColor="062040"),
         "SPEC_DIFF":PatternFill("solid",fgColor="3d2e00"),"NEW_SOP":PatternFill("solid",fgColor="1a1a1a")}
    for r in rows:
        rv=[r.get(c,"") for c in cols]; ws.append(rv); xl=ws[ws.max_row]
        rf=mtf.get(r.get("MATCH_TYPE",""),PatternFill("solid",fgColor="111111"))
        for cell in xl:
            cell.fill=rf; cell.font=Font(name="Calibri",size=9,color="E8F4FD")
            cell.alignment=Alignment(vertical="center"); cell.border=bdr
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for i,c in enumerate(cols,1):
        ws.column_dimensions[get_column_letter(i)].width=60 if "DESCRIPTION" in c or "DEFINITION" in c else 20
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

# ── UI HELPERS ─────────────────────────────────────────────────────────────────
def cc(s): return "#00e676" if s>=85 else ("#ffb300" if s>=60 else "#ff4757")
def bc(s): return "ch" if s>=85 else ("cm" if s>=60 else "cl")
def mpill(mt): c={"EXACT":"E","BRAND_MATCH":"B","SPEC_DIFF":"S","NEW_SOP":"N"}.get(mt,"N"); return f'<span class="pill p{c}">{mt}</span>'
def vpill(vs): c="V" if vs=="VALID" else "R"; return f'<span class="pill p{c}">{vs}</span>'
def epill(n):
    if any(x in n for x in ["MiniLM","BERT","PubMed"]): return f'<span class="pill pAI">🧠 {n}</span>'
    return f'<span class="pill pAI" style="color:var(--pink)">📊 {n}</span>'

def render_card(r,top=False):
    cls="rc r1" if top else ("rc r2" if r["rank"]==2 else "rc")
    s=r["score"]; col=cc(s); bar=bc(s)
    iss="".join(f'<span style="color:#ff4757;font-size:.7rem">⚠ {i}</span><br>' for i in r["val_issues"])
    st.markdown(f"""<div class="{cls}">
<div style="display:flex;justify-content:space-between;align-items:flex-start">
  <div><div class="rn">#{r['rank']} &nbsp; {r['npc_code']}</div>
  <div class="rd">{r['npc_desc']}</div>
  <div style="font-size:.72rem;color:var(--mu)">📁 {r['category']}  ·  🏷 {r['family']}</div></div>
  <div style="text-align:right"><div style="font-family:'JetBrains Mono',monospace;font-size:1.4rem;font-weight:800;color:{col}">{s:.0f}%</div>
  <div style="font-size:.65rem;color:var(--mu)">confidence</div></div>
</div>
<div class="cbw"><div class="{bar} cbar" style="width:{min(s,100):.0f}%"></div></div>
<div>{mpill(r['match_type'])}&nbsp;{vpill(r['val_status'])}</div>
{'<div style="margin-top:.3rem">'+ iss +'</div>' if iss else ""}
</div>""", unsafe_allow_html=True)

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="logo">🧬 MedMatch AI</div><div class="logo-sub">Medical Standardization Engine</div>', unsafe_allow_html=True)
    st.markdown("---")
    # Engine badges
    badges=[("🧠 Sentence Transformer","mbok" if ST_OK else "mbof"),
            ("⚡ Cross-Encoder","mbok" if (ST_OK and load_models()[1]) else "mbof"),
            ("📊 TF-IDF Fallback","mbok" if SKL_OK else "mbof"),
            ("🔁 RapidFuzz","mbok" if RF_OK else "mbof"),
            ("🌐 openFDA","mbok" if REQ_OK else "mbof")]
    for lbl,cls in badges:
        st.markdown(f'<div class="mb {cls}">{lbl}</div>', unsafe_allow_html=True)
    n_cat=len(S.embed_meta) if S.embed_meta else 0
    st.markdown("")
    st.markdown(f'<div class="mb {"mbok" if n_cat else "mbwn"}">📚 {"Catalog: "+str(n_cat)+" items" if n_cat else "No catalog"}</div>', unsafe_allow_html=True)
    st.markdown("---")
    # Nav
    for icon,key,lbl in [("🔍","search","Single Search"),("⚡","batch","Batch Process"),("📚","catalog","Catalog Manager"),("📊","dashboard","Dashboard")]:
        if st.button(f"{icon}  {lbl}",key=f"nav_{key}",use_container_width=True):
            S.page=key; st.rerun()
    st.markdown("---")
    if S.history:
        st.markdown('<div style="font-size:.7rem;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;margin-bottom:.3rem">Recent</div>', unsafe_allow_html=True)
        for item in list(S.history)[-5:][::-1]:
            st.markdown(f'<div class="hi">{item[:40]+"…" if len(item)>40 else item}</div>', unsafe_allow_html=True)
        if st.button("Clear",key="clr"): S.history.clear(); st.rerun()
    st.markdown("---")
    st.markdown('<div style="font-size:.62rem;color:var(--mu2)">MedMatch AI v2.0<br>Rwanda FDA · SOP/004</div>', unsafe_allow_html=True)

# ── PAGE: DASHBOARD ────────────────────────────────────────────────────────────
if S.page=="dashboard":
    st.markdown('<h1 class="pt">Dashboard</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">System health, engine status, and query statistics.</p>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5)
    for col,val,lbl,cls in [(c1,S.stats["queries"],"Total Queries","sb"),(c2,S.stats["exact"],"Exact","sg"),(c3,S.stats["brand"],"Brand Match","sb"),(c4,S.stats["spec"],"Spec Diff","sy"),(c5,S.stats["avg_ms"],"Avg ms","sp")]:
        with col: st.markdown(f'<div class="sc"><div class="sv {cls}">{val}</div><div class="sl">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("")
    # Pipeline
    st.markdown('<div class="sh">Processing Pipeline</div>', unsafe_allow_html=True)
    pipe=[("📥","INPUT",""),("🔤","NORMALIZE","Brand aliases"),("🧠","EMBED","384-dim vector"),("🔍","RETRIEVE","Top-20"),("⚡","RERANK","Cross-Encoder"),("✅","VALIDATE","Rules"),("🌐","ENRICH","openFDA"),("📤","OUTPUT","")]
    parts=[]
    for i,(ico,nm,desc) in enumerate(pipe):
        ok=nm in ("NORMALIZE","RETRIEVE","VALIDATE","OUTPUT")
        cls="pipe dn" if ok else "pipe"
        parts.append(f'<div class="{cls}"><span style="font-size:1.1rem">{ico}</span><b>{nm}</b><span style="font-size:.58rem;color:var(--mu)">{desc}</span></div>')
        if i<len(pipe)-1: parts.append('<span style="color:var(--bd2);font-size:1rem;align-self:center">→</span>')
    st.markdown(f'<div style="display:flex;flex-wrap:wrap;gap:.2rem;align-items:center;margin:.5rem 0">{"".join(parts)}</div>', unsafe_allow_html=True)
    # Capability table
    st.markdown('<div class="sh">Engine Capabilities</div>', unsafe_allow_html=True)
    cap={"Component":["Sentence Transformer (all-MiniLM-L6-v2)","Cross-Encoder (ms-marco-MiniLM)","HuggingFace Inference API","TF-IDF Fallback","RapidFuzz Fuzzy","openFDA Enrichment"],
         "Status":["✅" if ST_OK else "❌ pip install sentence-transformers","✅" if (ST_OK and load_models()[1]) else "⚠ not loaded","🔑 Needs HF_TOKEN in secrets","✅" if SKL_OK else "❌","✅" if RF_OK else "❌","✅" if REQ_OK else "❌"],
         "Purpose":["Primary semantic matching","Deep re-ranking top-20","Remote embed without local GPU","Fast keyword fallback","Edit-distance fallback","FDA device/drug metadata"]}
    st.dataframe(pd.DataFrame(cap),use_container_width=True,hide_index=True)
    if S.embed_meta:
        st.markdown('<div class="sh">Catalog Distribution</div>', unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            fvc=pd.Series([m["family"] for m in S.embed_meta]).value_counts().head(10)
            st.bar_chart(fvc,color="#00d4ff",height=200)
        with c2:
            cvc=pd.Series([m["category"] for m in S.embed_meta]).value_counts().head(10)
            st.bar_chart(cvc,color="#00e676",height=200)

# ── PAGE: CATALOG MANAGER ─────────────────────────────────────────────────────
elif S.page=="catalog":
    st.markdown('<h1 class="pt">Catalog Manager</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Upload your NPC catalog. MedMatch AI builds a semantic search index automatically.</p>', unsafe_allow_html=True)
    c1,c2=st.columns([3,2])
    with c1:
        st.markdown('<div class="sh">Upload NPC Catalog</div>', unsafe_allow_html=True)
        up=st.file_uploader("Upload (.xlsx, .csv)",type=["xlsx","xls","csv"],key="cup")
        if up:
            try:
                df_c=pd.read_csv(up) if up.name.endswith(".csv") else list(pd.read_excel(up,sheet_name=None).values())[0]
                df_c.columns=[str(c).strip() for c in df_c.columns]
                st.success(f"✓ {len(df_c):,} rows")
                st.dataframe(df_c.head(4),use_container_width=True)
                def best(kws): return sorted(df_c.columns,key=lambda c:sum(k.upper() in c.upper() for k in kws),reverse=True)
                cc_col=st.selectbox("NPC Code column",best(["npc","code"]))
                dc_col=st.selectbox("Description column",best(["description","desc","product","name"]))
                cat_opts=["(auto)"]+list(df_c.columns)
                ct_col=st.selectbox("Category column (optional)",cat_opts)
                ct_col=ct_col if ct_col!="(auto)" else None
                if st.button("🔨 Build Index",use_container_width=True):
                    with st.spinner("Building semantic index…"):
                        ok=build_index(df_c,dc_col,cc_col,ct_col)
                    if ok: st.success(f"✓ {len(S.embed_meta):,} products indexed · Engine: {S.model_name}")
                    else: st.error("Failed — check column selections.")
            except Exception as e: st.error(f"Error: {e}")
    with c2:
        st.markdown('<div class="sh">Index Status</div>', unsafe_allow_html=True)
        if S.embed_meta:
            for val,lbl,cls in [(len(S.embed_meta),"Products indexed","sb"),
                                 (S.embeddings.shape[1] if S.embeddings is not None else 0,"Embedding dim","sg"),
                                 (len(S.tfidf_vec.vocabulary_) if S.tfidf_vec else 0,"TF-IDF vocab","sp")]:
                st.markdown(f'<div class="sc" style="margin-bottom:.5rem"><div class="sv {cls}">{val:,}</div><div class="sl">{lbl}</div></div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(S.embed_meta[:6])[["code","desc","family"]],use_container_width=True,hide_index=True)
            if st.button("🗑 Clear Index",use_container_width=True):
                S.embed_meta=None; S.embeddings=None; S.tfidf_mat=None; S.tfidf_vec=None; st.rerun()
        else:
            st.info("No catalog loaded yet.\n\n**Expected columns:**\n- NPC Code\n- Product Description")

# ── PAGE: SINGLE SEARCH ────────────────────────────────────────────────────────
elif S.page=="search":
    st.markdown('<h1 class="pt">Product Search</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Enter any medical product — branded, abbreviated, or generic. MedMatch AI finds the closest NPC match.</p>', unsafe_allow_html=True)
    if not S.embed_meta:
        st.warning("⚠ No catalog loaded — go to Catalog Manager first.")
        if st.button("→ Catalog Manager"): S.page="catalog"; st.rerun()
        st.stop()
    c1,c2,c3,c4=st.columns([3,1,1,1])
    with c1: q=st.text_input("","",placeholder='e.g. "CAPROSYN 2-0 SH 70CM" or "3.5 cortex screw 26mm"',key="sq",label_visibility="collapsed")
    with c2: top_n=st.selectbox("Results",[3,5,10],index=1)
    with c3: do_e=st.toggle("openFDA",False)
    with c4: run=st.button("🔍 Search",use_container_width=True,disabled=not q.strip())
    # Quick examples
    examples=["CAPROSYN 2-0 SH 70CM","3.5MM CORTEX SCREW 26MM","POLYSORB 0 CTX 90CM",
               "PROXIMAL HUMERUS PLATE 7H","LIGASURE VESSEL SEALER","EXPERT TIBIA NAIL 10MM"]
    ec=st.columns(len(examples))
    for col,ex in zip(ec,examples):
        with col:
            if st.button(ex[:20]+"…" if len(ex)>20 else ex,key=f"ex_{ex}",use_container_width=True):
                st.session_state.sq=ex; run=True; q=ex
    st.markdown("---")
    if run and q.strip():
        S.history.append(q.strip())
        with st.spinner("Searching…"):
            res=medmatch(q.strip(),top_n=top_n,do_enrich=do_e)
        if "error" in res: st.error(res["error"]); st.stop()
        # Metric row
        c1,c2,c3,c4=st.columns(4)
        tr=res["results"][0] if res["results"] else {}
        with c1: st.markdown(f'<div class="sc"><div class="sv sb">{tr.get("score",0):.0f}%</div><div class="sl">Top Score</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="sc"><div class="sv sg">{len(res["results"])}</div><div class="sl">Results</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="sc"><div class="sv sp">{res["elapsed_ms"]}ms</div><div class="sl">Query Time</div></div>', unsafe_allow_html=True)
        with c4: st.markdown(f'<div class="sc"><div class="sv sy">{res["family"]}</div><div class="sl">Product Family</div></div>', unsafe_allow_html=True)
        st.markdown("")
        st.markdown(f'<div style="background:var(--sf);border:1px solid var(--bd);border-radius:8px;padding:.5rem 1rem;font-family:var(--mono);font-size:.76rem;color:var(--mu);margin:.4rem 0">Normalized: <span style="color:var(--ac)">{res["normalized"]}</span>&nbsp;&nbsp;{epill(res["engine"])}</div>', unsafe_allow_html=True)
        # Spec chips
        sp=res["specs"]; chips=[]
        chips+=[f"📏 {d}mm" for d in sp["dims"]]+[f"💧 {m}ml" for m in sp["ml"]]+[f"💊 {g}mg" for g in sp["mg"]]
        if sp["holes"]: chips.append(f"🔩 {sp['holes']}H")
        if sp["fr"]: chips+=[f"🩺 FR{f}" for f in sp["fr"]]
        if chips:
            st.markdown(" ".join(f'<span style="background:var(--sf3);border:1px solid var(--bd);border-radius:4px;padding:2px 8px;font-size:.7rem;font-family:var(--mono)">{c}</span>' for c in chips), unsafe_allow_html=True)
        st.markdown("")
        c_l,c_r=st.columns([3,1])
        with c_l:
            st.markdown('<div class="sh">Ranked Matches</div>', unsafe_allow_html=True)
            for r in res["results"]: render_card(r,r["rank"]==1)
        with c_r:
            st.markdown('<div class="sh">Top Match</div>', unsafe_allow_html=True)
            if res["results"]:
                best=res["results"][0]
                st.markdown(f'<div class="fda"><div class="ff">NPC Code</div><div class="fv">{best["npc_code"]}</div><div class="ff">Match Type</div><div>{mpill(best["match_type"])}</div><div class="ff" style="margin-top:.4rem">Validation</div><div>{vpill(best["val_status"])}</div><div class="ff" style="margin-top:.4rem">Family</div><div class="fv">{best["family"]}</div></div>', unsafe_allow_html=True)
            if res["enrich"]:
                ed=res["enrich"]; st.markdown('<div class="sh" style="margin-top:.8rem">FDA Enrichment</div>', unsafe_allow_html=True)
                flds=[("Generic Name",ed.get("device_name") or ed.get("generic_name","")),("FDA Code",ed.get("product_code","")),("Class",ed.get("device_class","")),("Regulation",ed.get("regulation","")),("Manufacturer",ed.get("manufacturer",""))]
                html='<div class="fda">'+''.join(f'<div class="ff">{l}</div><div class="fv">{v}</div>' for l,v in flds if v)+'</div>'
                st.markdown(html,unsafe_allow_html=True)
        # Export
        st.markdown("")
        rows=[{"INPUT_DESCRIPTION":q,"NPC_CODE":r["npc_code"],"NPC_DESCRIPTION":r["npc_desc"],"CATEGORY":r["category"],"MATCH_SCORE":r["score"],"MATCH_TYPE":r["match_type"],"VALIDATION_STATUS":r["val_status"],"VALIDATION_COMMENT":" | ".join(r["val_issues"]),"FDA_PRODUCT_CODE":res["enrich"].get("product_code",""),"FDA_DEFINITION":res["enrich"].get("definition","")[:200]} for r in res["results"]]
        st.download_button("⬇ Export Results (.xlsx)",data=to_excel(rows),file_name=f"medmatch_{datetime.now():%Y%m%d_%H%M%S}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── PAGE: BATCH ────────────────────────────────────────────────────────────────
elif S.page=="batch":
    st.markdown('<h1 class="pt">Batch Processing</h1>', unsafe_allow_html=True)
    st.markdown('<p class="ps">Upload an Excel file with multiple products. MedMatch AI maps all to NPC codes in one run.</p>', unsafe_allow_html=True)
    if not S.embed_meta:
        st.warning("⚠ No catalog loaded."); 
        if st.button("→ Catalog Manager"): S.page="catalog"; st.rerun()
        st.stop()
    c1,c2=st.columns([3,2])
    with c1:
        st.markdown('<div class="sh">Upload Products</div>', unsafe_allow_html=True)
        up=st.file_uploader("Upload (.xlsx, .csv)",type=["xlsx","xls","csv"],key="bup")
        if up:
            try:
                df_b=pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
                df_b.columns=[str(c).strip() for c in df_b.columns]
                st.success(f"✓ {len(df_b):,} rows"); st.dataframe(df_b.head(3),use_container_width=True)
                ranked=sorted(df_b.columns,key=lambda c:sum(k.upper() in c.upper() for k in ["description","desc","product","item","name"]),reverse=True)
                dc=st.selectbox("Description column",ranked)
                cc2=st.selectbox("Code column (opt)",["(none)"]+list(df_b.columns))
                mr=st.number_input("Max rows (0=all)",0,10000,0)
                de=st.toggle("openFDA enrichment",False)
            except Exception as e: st.error(f"Error: {e}"); st.stop()
            if st.button("⚡ Run Batch Matching",use_container_width=True):
                data=df_b.dropna(subset=[dc])
                if mr>0: data=data.head(mr)
                total=len(data)
                logs=[]; lb=st.empty()
                def log(m,k="info"):
                    ts=datetime.now().strftime("%H:%M:%S")
                    c={"ok":"co","warn":"cw","err":"ce","ai":"cai","net":"cn","info":"ci"}.get(k,"ci")
                    logs.append(f'<span class="{c}">[{ts}] {m}</span>')
                    lb.markdown('<div class="cons">'+'<br>'.join(logs[-25:])+'</div>',unsafe_allow_html=True)
                prog=st.progress(0,"Starting…")
                log(f"Processing {total:,} · engine={S.model_name}","ai")
                all_res=[]; t0=time.perf_counter()
                for i,row in enumerate(data.itertuples(index=False)):
                    desc=str(getattr(row,dc.replace(" ","_"),"")).strip()
                    code=str(getattr(row,cc2.replace(" ","_"),"") if cc2!="(none)" else f"R{i+1}").strip()
                    if not desc or desc.lower()=="nan": continue
                    res=medmatch(desc,top_n=1,do_enrich=de)
                    if "error" in res or not res.get("results"):
                        all_res.append({"INPUT_CODE":code,"INPUT_DESCRIPTION":desc,"NPC_CODE":"","NPC_DESCRIPTION":"","CATEGORY":"","MATCH_SCORE":0,"MATCH_TYPE":"NEW_SOP","VALIDATION_STATUS":"REVIEW","VALIDATION_COMMENT":"No match","FDA_PRODUCT_CODE":"","FDA_DEFINITION":""})
                    else:
                        t=res["results"][0]; ed=res.get("enrich",{})
                        all_res.append({"INPUT_CODE":code,"INPUT_DESCRIPTION":desc,"NPC_CODE":t["npc_code"],"NPC_DESCRIPTION":t["npc_desc"],"CATEGORY":t["category"],"MATCH_SCORE":t["score"],"MATCH_TYPE":t["match_type"],"VALIDATION_STATUS":t["val_status"],"VALIDATION_COMMENT":" | ".join(t["val_issues"]),"FDA_PRODUCT_CODE":ed.get("product_code",""),"FDA_DEFINITION":ed.get("definition","")[:200]})
                    prog.progress((i+1)/total,f"Processing {i+1:,}/{total:,}")
                    if i%50==0 and i>0:
                        nm=sum(1 for r in all_res if r["MATCH_TYPE"]!="NEW_SOP")
                        log(f"  {i+1:,}/{total:,} · {nm:,} matched · {time.perf_counter()-t0:.1f}s","ok")
                prog.progress(1.0,"Done ✓")
                S.batch_results=pd.DataFrame(all_res)
                nm=sum(1 for r in all_res if r["MATCH_TYPE"]!="NEW_SOP")
                nr=sum(1 for r in all_res if r["VALIDATION_STATUS"]=="REVIEW")
                log(f"Done — {len(all_res):,} products in {time.perf_counter()-t0:.1f}s","ok")
                log(f"Matched: {nm:,} · Review: {nr:,} · Unmatched: {len(all_res)-nm:,}","ai")
    with c2:
        if S.batch_results is not None:
            df_r=S.batch_results; total=len(df_r)
            st.markdown('<div class="sh">Results</div>', unsafe_allow_html=True)
            for v,l,c in [(total,"Total","sb"),(int((df_r["MATCH_TYPE"]!="NEW_SOP").sum()),"Matched","sg"),(int((df_r["VALIDATION_STATUS"]=="REVIEW").sum()),"Review","sy"),(int((df_r["MATCH_TYPE"]=="NEW_SOP").sum()),"Unmatched","sp")]:
                st.markdown(f'<div class="sc" style="margin-bottom:.4rem"><div class="sv {c}">{v:,}</div><div class="sl">{l}</div></div>', unsafe_allow_html=True)
            st.markdown("")
            vc=df_r["MATCH_TYPE"].value_counts().reset_index(); vc.columns=["Type","Count"]
            st.bar_chart(vc.set_index("Type"),color="#00d4ff",height=150)
            xlsx=to_excel(df_r.to_dict("records"))
            fn=f"medmatch_batch_{datetime.now():%Y%m%d_%H%M}.xlsx"
            st.download_button("⬇ Export Excel",data=xlsx,file_name=fn,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.download_button("⬇ Download CSV",data=df_r.to_csv(index=False).encode(),file_name=fn.replace(".xlsx",".csv"),mime="text/csv")
    if S.batch_results is not None:
        st.markdown('<div class="sh">Results Table</div>', unsafe_allow_html=True)
        df_r=S.batch_results
        f1,f2,f3=st.columns(3)
        with f1: fm=st.selectbox("Match Type",["All"]+sorted(df_r["MATCH_TYPE"].unique().tolist()))
        with f2: fv=st.selectbox("Validation",["All","VALID","REVIEW"])
        with f3: fs=st.text_input("Search","",placeholder="filter…",key="bsrch")
        view=df_r.copy()
        if fm!="All": view=view[view["MATCH_TYPE"]==fm]
        if fv!="All": view=view[view["VALIDATION_STATUS"]==fv]
        if fs: view=view[view["INPUT_DESCRIPTION"].str.upper().str.contains(fs.upper(),na=False)]
        def style_b(row):
            mt=row.get("MATCH_TYPE",""); vs=row.get("VALIDATION_STATUS","")
            if vs=="REVIEW": return ["background:rgba(255,71,87,.07)"]*len(row)
            if mt=="EXACT": return ["background:rgba(0,230,118,.04)"]*len(row)
            if mt=="NEW_SOP": return ["background:rgba(255,179,0,.05)"]*len(row)
            return [""]*len(row)
        sc=["INPUT_CODE","INPUT_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION","MATCH_SCORE","MATCH_TYPE","VALIDATION_STATUS","FDA_PRODUCT_CODE"]
        sc=[c for c in sc if c in view.columns]
        st.dataframe(view[sc].style.apply(style_b,axis=1).format({"MATCH_SCORE":"{:.0f}"}),use_container_width=True,height=460)
        st.caption(f"{len(view):,} / {len(df_r):,} rows")
