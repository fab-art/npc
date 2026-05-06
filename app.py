"""
IHS–NPC Harmonizer v3 — Manual Validation + Internet Search
Rwanda FDA | SOP ODDG/RES/SOP/004
"""
import os, json, re, io, warnings
from datetime import datetime
import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz, process as rfprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
try:
    import anthropic as _anthropic
    ANTHROPIC_OK = True
except ImportError:
    ANTHROPIC_OK = False
warnings.filterwarnings("ignore")

st.set_page_config(page_title="IHS–NPC Harmonizer", page_icon="🏥", layout="wide", initial_sidebar_state="expanded")

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Syne:wght@400;700;800&family=Inter:wght@300;400;500;600&display=swap');
:root{--bg:#0d1117;--surface:#161b22;--s2:#21262d;--border:#30363d;
  --accent:#00b347;--a2:#00e676;--warn:#f0a500;--danger:#f85149;
  --info:#388bfd;--purple:#8957e5;--text:#e6edf3;--muted:#8b949e;
  --mono:'JetBrains Mono',monospace;--head:'Syne',sans-serif;--body:'Inter',sans-serif;}
.stApp{background:var(--bg);color:var(--text);font-family:var(--body);}
section[data-testid="stSidebar"]{background:var(--surface)!important;border-right:1px solid var(--border);}
.stButton>button{background:var(--accent);color:#000;font-family:var(--head);font-weight:700;
  border:none;border-radius:6px;padding:.45rem 1.2rem;transition:all .15s;}
.stButton>button:hover{background:var(--a2);transform:translateY(-1px);box-shadow:0 4px 20px rgba(0,179,71,.35);}
.stButton>button:disabled{background:var(--s2);color:var(--muted);}
.stTextInput>div>div>input,.stTextArea>div>textarea{background:var(--s2)!important;border:1px solid var(--border)!important;border-radius:6px!important;color:var(--text)!important;}
.stSelectbox>div>div,.stMultiSelect>div>div{background:var(--s2)!important;border:1px solid var(--border)!important;border-radius:6px!important;color:var(--text)!important;}
div[data-testid="stFileUploader"]{border:1.5px dashed var(--border);border-radius:8px;background:var(--surface);}
.stProgress>div>div{background:var(--accent);}
hr{border-color:var(--border);}
.ptitle{font-family:var(--head);font-size:2.1rem;font-weight:800;
  background:linear-gradient(135deg,var(--a2),var(--info));
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
  letter-spacing:-.02em;margin-bottom:.1rem;}
.psub{color:var(--muted);font-size:.88rem;margin-bottom:1.2rem;}
.sec{font-family:var(--head);font-weight:700;font-size:1rem;
  border-left:3px solid var(--accent);padding-left:.65rem;margin:1rem 0 .55rem;}
.stbadge{display:inline-block;font-family:var(--mono);font-size:.63rem;font-weight:700;
  padding:2px 6px;border-radius:3px;background:var(--s2);color:var(--muted);margin-bottom:.2rem;}
.stbadge.a{background:rgba(0,179,71,.15);color:var(--a2);}
.stbadge.d{background:rgba(0,179,71,.08);color:var(--accent);}
.kc{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:.85rem 1rem;}
.kv{font-family:var(--head);font-size:1.8rem;font-weight:800;}
.kl{font-size:.7rem;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;}
.kg{color:var(--a2);}.kb{color:#79c0ff;}.ky{color:var(--warn);}.kr{color:var(--danger);}.kp{color:#d2a8ff;}.km{color:var(--muted);}
.mp{display:inline-block;font-family:var(--mono);font-size:.66rem;font-weight:700;padding:2px 6px;border-radius:3px;}
.pE{background:rgba(0,230,118,.15);color:#00e676;}.pB{background:rgba(56,139,253,.15);color:#79c0ff;}
.pS{background:rgba(240,165,0,.15);color:#f0a500;}.pN{background:rgba(139,148,158,.1);color:#8b949e;}
.pH{background:rgba(0,230,118,.15);color:#00e676;}.pM{background:rgba(240,165,0,.15);color:#f0a500;}
.pL{background:rgba(139,148,158,.1);color:#8b949e;}.pCR{background:rgba(248,81,73,.2);color:#f85149;}
.pVA{background:rgba(0,230,118,.12);color:#00e676;}.pRV{background:rgba(248,81,73,.12);color:#f85149;}
.pCO{background:rgba(0,230,118,.2);color:#00e676;}.pRJ{background:rgba(248,81,73,.2);color:#f85149;}
.pED{background:rgba(137,87,229,.2);color:#d2a8ff;}.pPE{background:rgba(139,148,158,.1);color:#8b949e;}
.vc{background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:.9rem 1.1rem;margin-bottom:.65rem;transition:border-color .12s;}
.vc:hover{border-color:var(--accent);}
.vc.sCO{border-left:4px solid var(--a2);}.vc.sRJ{border-left:4px solid var(--danger);}
.vc.sED{border-left:4px solid var(--purple);}.vc.sPE{border-left:4px solid var(--border);}
.vr{font-family:var(--mono);font-size:.7rem;color:var(--muted);margin-bottom:.25rem;}
.vd{font-family:var(--head);font-weight:600;font-size:.93rem;}
.vn{font-size:.8rem;color:#79c0ff;margin:.22rem 0;}
.vw{font-size:.76rem;color:var(--warn);margin-top:.22rem;}
.sc{background:#010409;border:1px solid var(--border);border-radius:8px;
  padding:.9rem 1.1rem;font-size:.83rem;margin-top:.55rem;line-height:1.55;}
.sl{font-family:var(--mono);font-size:.63rem;text-transform:uppercase;letter-spacing:.09em;color:var(--muted);margin-bottom:.12rem;}
.sv{color:var(--text);margin-bottom:.45rem;}
.aok{color:var(--a2);font-weight:700;}.acaution{color:var(--warn);font-weight:700;}.awrong{color:var(--danger);font-weight:700;}.aunk{color:var(--muted);font-weight:700;}
.con{background:#010409;border:1px solid var(--border);border-radius:8px;
  padding:.8rem 1rem;font-family:var(--mono);font-size:.74rem;color:#58a6ff;
  max-height:240px;overflow-y:auto;line-height:1.85;}
.cok{color:var(--a2);}.cwarn{color:var(--warn);}.cerr{color:var(--danger);}.cinfo{color:#58a6ff;}
.fc{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:.85rem 1rem;margin-bottom:.55rem;}
.fc.r{border-left:3px solid var(--accent);}.fc.o{border-left:3px solid var(--s2);}.fc.ok{border-color:var(--accent);background:rgba(0,179,71,.06);}
.fl{font-family:var(--head);font-weight:700;font-size:.9rem;}
.fn{font-size:.73rem;color:var(--muted);margin-top:.1rem;}
</style>""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
STEPS = ["Upload","Map Columns","Configure","Process","Review","Validate","Export"]

def _init():
    for k,v in {"step":0,"files":{},"col_map":{},"config":{},"results":None,
                "logs":[],"run_done":False,"validations":{},"search_cache":{},
                "api_key":os.environ.get("ANTHROPIC_API_KEY","")}.items():
        if k not in st.session_state: st.session_state[k]=v
_init(); S=st.session_state

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p style="font-family:\'Syne\',sans-serif;font-size:1.05rem;font-weight:800;color:#00e676;letter-spacing:.06em;margin-bottom:0;">🏥 IHS–NPC</p>',unsafe_allow_html=True)
    st.markdown('<p style="font-size:.68rem;color:#8b949e;margin-top:0;margin-bottom:1.1rem;">Harmonizer v3 · Rwanda FDA</p>',unsafe_allow_html=True)
    for i,name in enumerate(STEPS):
        cls = "d" if i<S.step else ("a" if i==S.step else "")
        icon = "✓ " if i<S.step else ("→ " if i==S.step else "  ")
        col = "#00b347" if i<S.step else ("#e6edf3" if i==S.step else "#484f58")
        wt  = "700" if i==S.step else "400"
        st.markdown(f'<div class="stbadge {cls}">{icon}{i+1}</div><div style="font-family:\'Syne\',sans-serif;font-size:.82rem;font-weight:{wt};color:{col};margin-bottom:.4rem;">{name}</div>',unsafe_allow_html=True)
    st.markdown("---")
    st.markdown('<div style="font-size:.7rem;color:#8b949e;font-family:\'JetBrains Mono\',monospace;margin-bottom:.2rem;">ANTHROPIC API KEY</div>',unsafe_allow_html=True)
    ak=st.text_input("ak",value=S.api_key,type="password",label_visibility="collapsed",placeholder="sk-ant-... (for internet search)")
    if ak: S.api_key=ak
    st.markdown(f'<div style="font-size:.68rem;color:{"#00e676" if S.api_key else "#8b949e"};">{"● API key set — search enabled" if S.api_key else "○ No key — search disabled"}</div>',unsafe_allow_html=True)
    st.markdown("---")
    if st.button("↩ Reset",use_container_width=True):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()
    st.markdown('<div style="font-size:.65rem;color:#484f58;margin-top:1.5rem;">SOP ODDG/RES/SOP/004<br>v3 · Validation + Search</div>',unsafe_allow_html=True)

# ── Knowledge bases ───────────────────────────────────────────────────────────
BRAND_ALIASES={"CORTEX SCREW":"CORTICAL SCREW","CORTEX SCREWS":"CORTICAL SCREWS","CONCELOUS":"CANCELLOUS","CONCEULOUS":"CANCELLOUS","CANSELLOUS":"CANCELLOUS","CAPROSYN":"POLYGLYCONATE","VICRYL":"POLYGLACTIN 910","MONOCRYL":"POLIGLECAPRONE 25","PROLENE":"POLYPROPYLENE","PDS":"POLYDIOXANONE","SURGICEL":"OXIDIZED REGENERATED CELLULOSE","SURGICELL":"OXIDIZED REGENERATED CELLULOSE","HEMOLOCK":"HEM-O-LOK","HEMOLOC":"HEM-O-LOK","POLYSORB":"POLYGLYCOLIC ACID LACTIDE","BIOSYN":"GLYCOMER 631","VELOSORB":"POLYGLYCOLIC ACID LACTIDE","MAXON":"POLYGLYCOLATE","SURGIPRO":"POLYPROPYLENE","TICRON":"COATED POLYESTER","SURGIDAC":"POLYESTER","SOFSILK":"SILK","SURGILON":"BRAIDED NYLON","MONOSOF":"NYLON MONOFILAMENT","LIGASURE":"VESSEL SEALING SYSTEM","INSUFLATION":"INSUFFLATION","REDON":"CLOSED WOUND SUCTION DRAIN","NOVAFIL":"POLYBUTESTER","VASCUFIL":"POLYPROPYLENE VASCULAR"}
ANAT_INCOMPAT={"DISTAL RADIUS":["DHS","135","HIP"],"VOLAR LOCKING":["DHS","135","HIP"],"VOLAR LCP":["DHS","135","HIP"],"STALOID":["DHS","135","HIP"],"FIBULA":["DHS","135","HIP"],"TIBIA DISTAL":["DHS","135","HIP"],"TIBIA PROXIMAL":["DHS","135","HIP"],"TIBIA LOCKING":["DHS","135","HIP"],"CLAVICLE":["DHS","135","TIBIA","HIP"],"CLAVICULA":["DHS","135","TIBIA","HIP"],"PROXIMAL HUMERUS":["DHS","135","HIP","TIBIA"],"DISTAL HUMERUS":["DHS","135","HIP","TIBIA"],"FILS DE CERCLAGE":["TIGHTNER","TIGHTENER"],"CERCLAGE WIRE":["TIGHTNER","TIGHTENER"],"DRAIN DE REDON":["CHEST DRAIN","THORACIC","INTERCOSTAL"],"REDON":["CHEST DRAIN","THORACIC","INTERCOSTAL"]}
FAM_PAT={"SCREW":r"\bSCREW\b|\bCORTEX\b|\bCORTICAL\b|\bCANCELLOUS\b|\bINTERLOCKING\b","PLATE":r"\bPLATE\b|\bLCP\b|\bDCP\b|\bDHS\b","NAIL":r"\bNAIL\b|\bPFNA\b","KWIRE":r"\bKIRSCHNER\b|\bK-WIRE\b","CERCLAGE":r"\bCERCLAGE\b|\bFILS DE CERCLAGE\b","SUTURE":r"\bSUTURE\b|POLYSORB|BIOSYN|CAPROSYN|SURGIPRO|TICRON|VELOSORB|SOFSILK|SURGILON|POLYCRYL|MAXON|CHROMIC GUT|PLAIN GUT|\bNYLON\b|V-LOC|MONOCRYL|VICRYL|PROLENE|MONOSOF","TROCAR":r"\bTROCAR\b|\bCANNULA\b","STAPLER":r"\bSTAPLER\b|\bGIA\b|EEA|CEEA|LINEAR CUTTER","DRAIN":r"\bDRAIN\b|\bREDON\b","ELECTRODE":r"\bELECTRODE\b|\bBIPOLAR\b","CLIP":r"\bHEMOLOCK\b|\bENDOCLIP\b","MESH":r"\bMESH\b","BLADE":r"\bBLADE\b"}
IMPL_PAT=r"\bSCREW\b|\bPLATE\b|\bNAIL\b|\bKIRSCHNER\b|\bK-WIRE\b|\bCERCLAGE\b|\bCANCELLOUS\b|\bCORTEX\b|\bCORTICAL\b|\bINTERLOCKING\b|\bANCHOR\b|\bMESH\b|\bPFNA\b"
G500={"SCREW":"CANC","PLATE":"MISP","NAIL":"NAIL","KWIRE":"KWIR","CERCLAGE":"KWIR","MESH":"MESH"}
G200={"SUTURE":"SUTS","TROCAR":"TROC","STAPLER":"STAP","DRAIN":"DREN","ELECTRODE":"ELCT","CLIP":"ECLP","BLADE":"BLAD"}

# ── Preprocessing ─────────────────────────────────────────────────────────────
def norm(t):
    if pd.isna(t): return ""
    s=str(t).upper().strip()
    s=re.sub(r"(\d+),(\d+)",r"\1.\2",s); s=re.sub(r"[×✕]","X",s); s=re.sub(r"\*","",s)
    s=re.sub(r"\.0\s*MM","MM",s); s=re.sub(r"(\d+\.?\d*)\s*MM",r"\1MM",s)
    s=re.sub(r"(\d+\.?\d*)\s*ML",r"\1ML",s); s=re.sub(r"\b(?:CH|FG)\s*(\d+)",r"FR\1",s)
    s=re.sub(r"\b(\d+)\s*-\s*HOLES?\b",r"\1 HOLES",s); s=re.sub(r"\bHOLE\b","HOLES",s)
    s=re.sub(r"\bLH\b","LEFT",s); s=re.sub(r"\bRH\b","RIGHT",s)
    s=re.sub(r"\bS\.T\.\b|\bS/T\b|\bSELF-TAPPING\b","SELF TAPPING",s)
    return re.sub(r"\s{2,}"," ",s).strip(" ,.-")

def gen_norm(t):
    s=norm(t)
    for b,g in BRAND_ALIASES.items():
        s=re.sub(r"\b"+re.escape(b)+r"\b",g,s)
    return s

def specs(t):
    u=norm(t)
    dims=re.findall(r"(\d+\.?\d*)MM",u); ml=re.findall(r"(\d+\.?\d*)ML",u)
    fr=re.findall(r"FR(\d+)",u); holes=re.search(r"(\d+)\s*HOLES?",u)
    return {"mm":sorted([float(d) for d in dims]),"ml":[float(m) for m in ml],
            "fr":[int(f) for f in fr],"holes":int(holes.group(1)) if holes else None}

def fam(t):
    u=norm(t)
    for f,p in FAM_PAT.items():
        if re.search(p,u): return f
    return "OTHER"

def is_impl(t): return bool(re.search(IMPL_PAT,norm(t)))

# ── Validation ────────────────────────────────────────────────────────────────
def validate(rd,nd):
    issues=[]; ur=norm(rd); un=norm(nd)
    for kw,bads in ANAT_INCOMPAT.items():
        if kw.upper() in ur:
            for bad in bads:
                if bad.upper() in un: issues.append(f"Anatomy mismatch: '{kw}'↔'{bad}' (AO Foundation/PMC)")
    rf=fam(rd); nf=fam(nd)
    if rf!="OTHER" and nf!="OTHER" and rf!=nf: issues.append(f"Family mismatch: RHIC={rf} NPC={nf}")
    rs=specs(rd); ns=specs(nd)
    if rs["mm"] and ns["mm"] and abs(rs["mm"][0]-ns["mm"][0])>=0.15: issues.append(f"Size: RHIC {rs['mm'][0]}mm vs NPC {ns['mm'][0]}mm")
    if rs["ml"] and ns["ml"] and abs(rs["ml"][0]-ns["ml"][0])>0.5: issues.append(f"Volume: RHIC {rs['ml'][0]}ml vs NPC {ns['ml'][0]}ml")
    if rs["fr"] and ns["fr"] and rs["fr"][0]!=ns["fr"][0]: issues.append(f"FR size: FR{rs['fr'][0]} vs FR{ns['fr'][0]}")
    if rs["holes"] and ns["holes"] and rs["holes"]!=ns["holes"]: issues.append(f"Holes: {rs['holes']} vs {ns['holes']}")
    if (rs["ml"] and ns["mm"] and not ns["ml"]) or (rs["mm"] and ns["ml"] and not ns["mm"]): issues.append("Unit crossover: ml vs mm")
    return ("REVIEW"," | ".join(issues)) if issues else ("VALID","")

# ── Internet search ───────────────────────────────────────────────────────────
SEARCH_SYS="""You are a senior medical device classification specialist with expertise in WHO MeDevIS/GMDN, AO Foundation trauma classification, FDA 21 CFR Part 888, and Rwanda NPC SOP ODDG/RES/SOP/004.

Respond ONLY with a JSON object (no markdown, no preamble):
{
  "product_type": "brief device type (<=6 words)",
  "generic_name": "GMDN-aligned generic name",
  "clinical_use": "one sentence clinical use",
  "product_family": "SCREW|PLATE|NAIL|SUTURE|TROCAR|STAPLER|DRAIN|OTHER",
  "match_assessment": "CORRECT|LIKELY CORRECT|UNCERTAIN|LIKELY INCORRECT|INCORRECT",
  "match_reasoning": "2-3 sentences citing sources found",
  "specification_notes": "key specs confirmed or discrepancies",
  "recommendation": "one actionable sentence for the data officer",
  "sources": ["url1","url2"]
}"""

def search_product(rd,nd,api_key):
    ck=f"{rd}|||{nd}"
    if ck in S.search_cache: return S.search_cache[ck]
    if not ANTHROPIC_OK: return {"error":"anthropic not installed: pip install anthropic"}
    if not api_key: return {"error":"No API key. Enter key in sidebar."}
    try:
        client=_anthropic.Anthropic(api_key=api_key)
        msg=f"Investigate this medical device mapping:\n\nIHS PRODUCT: {rd}\nMATCHED NPC: {nd}\n\nSearch the web for: 1) What is this product — type, generic name, clinical use? 2) Is the NPC entry clinically appropriate? 3) Any critical anatomy, size, or category differences?\n\nReturn JSON assessment."
        resp=client.messages.create(model="claude-sonnet-4-20250514",max_tokens=1200,system=SEARCH_SYS,
            tools=[{"type":"web_search_20250305","name":"web_search"}],
            messages=[{"role":"user","content":msg}])
        text=""
        for block in resp.content:
            if hasattr(block,"text"): text+=block.text
        text=text.strip()
        if text.startswith("```"): text=re.sub(r"```[a-z]*\n?","",text).strip("`").strip()
        result=json.loads(text); S.search_cache[ck]=result; return result
    except json.JSONDecodeError as e: return {"error":f"JSON parse error: {e}","raw":text[:500]}
    except Exception as e: return {"error":str(e)}

def render_result(res):
    if "error" in res:
        st.error(f"Search error: {res['error']}")
        if "raw" in res: st.code(res["raw"][:400])
        return
    assess=res.get("match_assessment","UNCERTAIN")
    acls={"CORRECT":"aok","LIKELY CORRECT":"aok","UNCERTAIN":"acaution","LIKELY INCORRECT":"awrong","INCORRECT":"awrong"}.get(assess,"aunk")
    icon={"CORRECT":"✅","LIKELY CORRECT":"✅","UNCERTAIN":"⚠️","LIKELY INCORRECT":"🔴","INCORRECT":"🔴"}.get(assess,"❓")
    st.markdown(f"""<div class="sc">
<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:.7rem;">
  <div>
    <div class="sl">Product Type</div><div class="sv"><strong>{res.get("product_type","—")}</strong></div>
    <div class="sl" style="margin-top:.4rem;">Generic Name (GMDN)</div><div class="sv" style="color:#79c0ff;">{res.get("generic_name","—")}</div>
    <div class="sl" style="margin-top:.4rem;">Product Family</div><div class="sv">{res.get("product_family","—")}</div>
  </div>
  <div>
    <div class="sl">Match Assessment</div>
    <div class="sv {acls}" style="font-size:1.05rem;">{icon} {assess}</div>
    <div class="sl" style="margin-top:.4rem;">Clinical Use</div>
    <div class="sv" style="color:#8b949e;font-size:.8rem;">{res.get("clinical_use","—")}</div>
  </div>
</div>
<div class="sl">Reasoning</div>
<div class="sv" style="font-size:.81rem;line-height:1.5;border-left:2px solid #30363d;padding-left:.55rem;">{res.get("match_reasoning","—")}</div>
<div class="sl" style="margin-top:.6rem;">Specification Notes</div>
<div class="sv" style="font-size:.8rem;">{res.get("specification_notes","—")}</div>
<div style="background:rgba(0,179,71,.06);border:1px solid rgba(0,179,71,.2);border-radius:6px;padding:.45rem .65rem;margin-top:.6rem;">
  <span style="font-family:'JetBrains Mono',monospace;font-size:.62rem;color:#8b949e;text-transform:uppercase;letter-spacing:.08em;">Recommendation</span><br>
  <span style="font-size:.82rem;color:#00e676;">{res.get("recommendation","—")}</span>
</div></div>""",unsafe_allow_html=True)
    srcs=res.get("sources",[])
    if srcs:
        with st.expander(f"📎 {len(srcs)} source(s)"):
            for s in srcs: st.markdown(f"- [{s}]({s})" if str(s).startswith("http") else f"- {s}")

# ── Matching engine ───────────────────────────────────────────────────────────
def build_npc_idx(df,dc,cc):
    idx={}
    for _,row in df.iterrows():
        d=str(row[dc]).strip(); c=str(row[cc]).strip(); e={"code":c,"desc":d}
        idx[norm(d)]=e; idx[gen_norm(d)]=e
    return idx

def build_phc_idx(df,dc,nc,ic):
    idx={}
    for _,row in df.iterrows():
        d=str(row[dc]).strip()
        idx[gen_norm(d)]={"desc":d,"npc_code":str(row[nc]).strip() if nc else "","ihbs_code":str(row[ic]).strip() if ic and ic in df.columns else ""}
    return idx

def match_one(rd,ni,pi,nt,pt):
    n_=norm(rd); g_=gen_norm(rd)
    for k in [n_,g_]:
        if k in ni:
            e=ni[k]; vs,vc=validate(rd,e["desc"])
            mt="BRAND_MATCH" if k==g_ and k!=n_ else "EXACT"
            return _r(rd,e["code"],e["desc"],100,"NPC",mt,"HIGH",vs,vc or ("Brand/terminology normalised" if mt=="BRAND_MATCH" else ""))
    ch=list(ni.keys())
    if ch:
        best=rfprocess.extractOne(g_,ch,scorer=fuzz.token_sort_ratio)
        if best and best[1]>=nt:
            e=ni[best[0]]; sc=best[1]
            mt="EXACT" if sc>=95 else ("BRAND_MATCH" if sc>=85 else "SPEC_DIFF")
            cf="HIGH" if sc>=80 else ("MEDIUM" if sc>=nt else "LOW")
            vs,vc=validate(rd,e["desc"])
            return _r(rd,e["code"],e["desc"],sc,"NPC",mt,cf,vs,vc)
    if pi:
        pc=list(pi.keys()); best=rfprocess.extractOne(g_,pc,scorer=fuzz.token_sort_ratio)
        if best and best[1]>=pt:
            e=pi[best[0]]; sc=best[1]
            cf="HIGH" if sc>=80 else ("MEDIUM" if sc>=pt else "LOW")
            vs,vc=validate(rd,e["desc"])
            return _r(rd,e.get("npc_code",""),e["desc"],sc,"PHC","SPEC_DIFF",cf,vs,vc,ihbs=e.get("ihbs_code",""))
    f_=fam(rd)
    code=f"500{G500.get(f_,'MISC')}NEW" if is_impl(rd) else f"200{G200.get(f_,'CONS')}NEW"
    return _r(rd,code,rd,0,"UNMATCHED","NEW_SOP","LOW","REVIEW","No NPC/PHC match — new SOP code required")

def _r(d,nc,nd,sc,src,mt,cf,vs,vc,ihbs=""):
    return {"ORIGINAL_DESCRIPTION":d,"NPC_CODE":nc,"NPC_DESCRIPTION":nd,"IHBS_CODE":ihbs,
            "MATCH_SCORE":sc,"MATCH_SOURCE":src,"MATCH_TYPE":mt,"CONFIDENCE":cf,
            "VALIDATION_STATUS":vs,"VALIDATION_COMMENT":vc,"PRODUCT_FAMILY":fam(d),
            "MANUAL_STATUS":"PENDING","MANUAL_NOTE":"","EDITED_NPC_CODE":""}

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(df):
    wb=Workbook(); thin=Side(style="thin",color="D0D0D0"); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws=wb.active; ws.title="Harmonized Mapping"; cols=list(df.columns); ws.append(cols)
    hf=PatternFill("solid",fgColor="1F3864")
    for cell in ws[1]:
        cell.fill=hf; cell.font=Font(name="Calibri",bold=True,color="FFFFFF",size=10)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cell.border=bdr
    ws.row_dimensions[1].height=36
    fs={"NPC":PatternFill("solid",fgColor="E2EFDA"),"PHC":PatternFill("solid",fgColor="DDEBF7"),"UNMATCHED":PatternFill("solid",fgColor="FFF2CC")}
    fm={"CONFIRMED":PatternFill("solid",fgColor="C6EFCE"),"REJECTED":PatternFill("solid",fgColor="FCE4D6"),"EDITED":PatternFill("solid",fgColor="E6D9F7"),"PENDING":PatternFill("solid",fgColor="F5F5F5")}
    fv={"VALID":PatternFill("solid",fgColor="C6EFCE"),"REVIEW":PatternFill("solid",fgColor="FCE4D6")}
    sc=cols.index("MATCH_SOURCE")+1 if "MATCH_SOURCE" in cols else None
    mc=cols.index("MANUAL_STATUS")+1 if "MANUAL_STATUS" in cols else None
    vc=cols.index("VALIDATION_STATUS")+1 if "VALIDATION_STATUS" in cols else None
    for rv in df.itertuples(index=False):
        rl=list(rv); ws.append(rl); xl=ws[ws.max_row]
        src=rl[sc-1] if sc else ""; man=rl[mc-1] if mc else ""; val=rl[vc-1] if vc else ""
        rf_=fs.get(src,PatternFill("solid",fgColor="F5F5F5"))
        for cell in xl: cell.fill=rf_; cell.font=Font(name="Calibri",size=9); cell.alignment=Alignment(vertical="center"); cell.border=bdr
        if mc: xl[mc-1].fill=fm.get(man,rf_)
        if vc: xl[vc-1].fill=fv.get(val,rf_)
    for col,w in {"A":22,"B":65,"C":18,"D":65,"E":18,"F":14,"G":16,"H":18,"I":16,"J":16,"K":65,"L":18,"M":16,"N":20,"O":65}.items():
        if ord(col)-64<=ws.max_column: ws.column_dimensions[col].width=w
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    ws2=wb.create_sheet("Validation Summary"); total=len(df)
    sv=df["MATCH_SOURCE"].value_counts() if "MATCH_SOURCE" in df.columns else pd.Series()
    mv=df["MANUAL_STATUS"].value_counts() if "MANUAL_STATUS" in df.columns else pd.Series()
    vv=df["VALIDATION_STATUS"].value_counts() if "VALIDATION_STATUS" in df.columns else pd.Series()
    mtv=df["MATCH_TYPE"].value_counts() if "MATCH_TYPE" in df.columns else pd.Series()
    rows_=[["IHS–NPC Harmonization · Validation Report","","",""],
           [f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Rwanda FDA  |  SOP ODDG/RES/SOP/004","","",""],
           ["","","",""],["MATCH SOURCE","COUNT","% TOTAL",""],
           ["NPC (Primary)",int(sv.get("NPC",0)),f"{sv.get('NPC',0)/total*100:.1f}%",""],
           ["PHC (Fallback)",int(sv.get("PHC",0)),f"{sv.get('PHC',0)/total*100:.1f}%",""],
           ["Unmatched",int(sv.get("UNMATCHED",0)),f"{sv.get('UNMATCHED',0)/total*100:.1f}%",""],
           ["TOTAL",total,"100%",""],["","","",""],
           ["MANUAL VALIDATION","COUNT","% TOTAL",""],
           ["CONFIRMED",int(mv.get("CONFIRMED",0)),f"{mv.get('CONFIRMED',0)/total*100:.1f}%",""],
           ["REJECTED",int(mv.get("REJECTED",0)),f"{mv.get('REJECTED',0)/total*100:.1f}%",""],
           ["EDITED",int(mv.get("EDITED",0)),f"{mv.get('EDITED',0)/total*100:.1f}%",""],
           ["PENDING",int(mv.get("PENDING",0)),f"{mv.get('PENDING',0)/total*100:.1f}%","Awaiting review"],
           ["","","",""],["VALIDATION RULES","COUNT","% TOTAL",""],
           ["VALID",int(vv.get("VALID",0)),f"{vv.get('VALID',0)/total*100:.1f}%",""],
           ["REVIEW",int(vv.get("REVIEW",0)),f"{vv.get('REVIEW',0)/total*100:.1f}%","Flagged"],
           ["","","",""],["MATCH TYPE","COUNT","% TOTAL",""],
           ["EXACT",int(mtv.get("EXACT",0)),f"{mtv.get('EXACT',0)/total*100:.1f}%",""],
           ["BRAND_MATCH",int(mtv.get("BRAND_MATCH",0)),f"{mtv.get('BRAND_MATCH',0)/total*100:.1f}%",""],
           ["SPEC_DIFF",int(mtv.get("SPEC_DIFF",0)),f"{mtv.get('SPEC_DIFF',0)/total*100:.1f}%",""],
           ["NEW_SOP",int(mtv.get("NEW_SOP",0)),f"{mtv.get('NEW_SOP',0)/total*100:.1f}%",""]]
    hf2=PatternFill("solid",fgColor="1F3864"); sf2=PatternFill("solid",fgColor="1A5C1A")
    for i,rv in enumerate(rows_,1):
        ws2.append(rv); xl=ws2[i]
        is_t=i<=2; is_h=rv[0] in ["MATCH SOURCE","MANUAL VALIDATION","VALIDATION RULES","MATCH TYPE"]
        for cell in xl:
            if is_t: cell.fill=hf2; cell.font=Font(name="Calibri",bold=True,color="FFFFFF",size=11)
            elif is_h: cell.fill=sf2; cell.font=Font(name="Calibri",bold=True,color="FFFFFF",size=10)
            else: cell.font=Font(name="Calibri",size=10)
            cell.alignment=Alignment(vertical="center")
    ws2.column_dimensions["A"].width=32; ws2.column_dimensions["B"].width=12
    ws2.column_dimensions["C"].width=12; ws2.column_dimensions["D"].width=40
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

# ═════════════════════════════════════════════════════════════════════════════
# STEP PAGES
# ═════════════════════════════════════════════════════════════════════════════

def step_upload():
    st.markdown('<h1 class="ptitle">Data Upload</h1>',unsafe_allow_html=True)
    st.markdown('<p class="psub">Upload source files. IHS, NPC, and PHC are required.</p>',unsafe_allow_html=True)
    FDEFS=[("ihs","IHS File","Consumables to be mapped",True,["RHIC Code","Description"]),
           ("npc","NPC File","National Product Catalogue",True,["NPC Code","Description"]),
           ("phc","PHC File","PHC catalogue (fallback)",True,["Description","NPC Code","IHBS Code"]),
           ("rhic","RHIC File","Comparison reference",False,["Description"])]
    for key,label,note,req,ecols in FDEFS:
        loaded=key in S.files and S.files[key] is not None
        cls="ok" if loaded else ("r" if req else "o")
        st.markdown(f'<div class="fc {cls}"><div class="fl">{label} <span style="font-size:.67rem;color:#8b949e;">{"🔴 REQUIRED" if req else "⚪ OPTIONAL"}</span></div><div class="fn">{note} · {", ".join(ecols)}</div></div>',unsafe_allow_html=True)
        up=st.file_uploader(f"Upload {label}",type=["xlsx","xls","csv"],key=f"up_{key}",label_visibility="collapsed")
        if up:
            try:
                if up.name.endswith(".csv"):
                    df_=pd.read_csv(up)
                else:
                    xl_=pd.read_excel(up,sheet_name=None)
                    snames=list(xl_.keys())
                    sheet=st.selectbox(f"Sheet ({label})",snames,key=f"sh_{key}") if len(snames)>1 else snames[0]
                    df_=xl_[sheet]
                df_.columns=[str(c).strip() for c in df_.columns]
                S.files[key]=df_; st.success(f"✓ {len(df_):,} rows · {len(df_.columns)} columns")
            except Exception as e: st.error(str(e))
        elif loaded: st.info(f"✓ Loaded: {len(S.files[key]):,} rows")
    st.markdown("---"); _,c2=st.columns([3,1])
    with c2:
        ok=all(k in S.files for k in ["ihs","npc","phc"])
        if st.button("Next: Map Columns →",disabled=not ok,use_container_width=True): S.step=1; st.rerun()
    if not ok: st.caption("⚠ Upload IHS, NPC and PHC files to continue")

def step_map():
    st.markdown('<h1 class="ptitle">Column Mapping</h1>',unsafe_allow_html=True)
    st.markdown('<p class="psub">Select the relevant columns from each uploaded file.</p>',unsafe_allow_html=True)
    cm=S.col_map
    def sel(fk,role,label,kws,optnone=False):
        df_=S.files.get(fk);
        if df_ is None: return
        base=list(df_.columns); opts=(["(none)"]+base) if optnone else base
        saved=cm.get(fk,{}).get(role)
        def sc_(c): return sum(k.upper() in c.upper() for k in kws)
        sorted_=["(none)"]+sorted(base,key=sc_,reverse=True) if optnone else sorted(base,key=sc_,reverse=True)
        idx=sorted_.index(saved) if saved in sorted_ else 0
        ch=st.selectbox(label,sorted_,index=idx,key=f"cm_{fk}_{role}")
        cm.setdefault(fk,{})[role]=ch if ch!="(none)" else None
        if ch and ch!="(none)": st.caption(f"e.g. {', '.join(str(v) for v in df_[ch].dropna().head(2).tolist())}")
    c1,c2=st.columns(2)
    with c1:
        st.markdown('<div class="sec">IHS File</div>',unsafe_allow_html=True)
        sel("ihs","desc","Description column",["desc","product","consumable","item","name"])
        sel("ihs","code","RHIC Code column (optional)",["rhic","code"],optnone=True)
    with c2:
        st.markdown('<div class="sec">NPC File</div>',unsafe_allow_html=True)
        sel("npc","code","NPC Code column",["npc","code","product code"])
        sel("npc","desc","NPC Description column",["desc","product","npc desc","rw product"])
    st.markdown("")
    c1,c2,c3=st.columns(3)
    with c1:
        st.markdown('<div class="sec">PHC File</div>',unsafe_allow_html=True)
        sel("phc","desc","Description column",["desc","product","name","item"])
    with c2:
        st.markdown(""); sel("phc","npc_code","NPC Code column",["npc","code"])
    with c3:
        st.markdown(""); sel("phc","ihbs_code","IHBS Code column (optional)",["ihbs","code"],optnone=True)
    if "rhic" in S.files:
        st.markdown('<div class="sec">RHIC File</div>',unsafe_allow_html=True)
        sel("rhic","desc","Description column",["desc","product","name"])
    S.col_map=cm
    st.markdown("---"); c1,_,c3=st.columns([1,3,1])
    with c1:
        if st.button("← Back",use_container_width=True): S.step=0; st.rerun()
    with c3:
        ready=(cm.get("ihs",{}).get("desc") and cm.get("npc",{}).get("code") and cm.get("npc",{}).get("desc") and cm.get("phc",{}).get("desc") and cm.get("phc",{}).get("npc_code"))
        if st.button("Next: Configure →",disabled=not ready,use_container_width=True): S.step=2; st.rerun()

def step_configure():
    st.markdown('<h1 class="ptitle">Configuration</h1>',unsafe_allow_html=True)
    st.markdown('<p class="psub">Tune matching thresholds, validation rules, and search behaviour.</p>',unsafe_allow_html=True)
    c1,c2=st.columns([3,2])
    with c1:
        st.markdown('<div class="sec">Match Thresholds</div>',unsafe_allow_html=True)
        nt=st.slider("NPC minimum score",50,100,S.config.get("npc_thresh",65))
        pt=st.slider("PHC fallback minimum score",40,95,S.config.get("phc_thresh",60))
        st.markdown('<div class="sec">Confidence Bands</div>',unsafe_allow_html=True)
        a,b=st.columns(2)
        ht=a.number_input("HIGH floor",60,100,S.config.get("high_thresh",80))
        mt_=b.number_input("MEDIUM floor",40,99,S.config.get("med_thresh",65))
        st.markdown('<div class="sec">Options</div>',unsafe_allow_html=True)
        ub=st.toggle("Apply brand→generic aliases",value=S.config.get("use_brand",True))
        mr=st.number_input("Max rows (0=all)",0,50000,S.config.get("max_rows",0))
    with c2:
        st.markdown('<div class="sec">Validation Rules</div>',unsafe_allow_html=True)
        rules_={"anatomy":"Anatomy incompatibility (DHS≠Distal Radius, cerclage wire≠tightener)","family":"Product family cross-check","size":"Size tolerance ±0.1mm","holes":"Plate hole count","units":"Unit crossover ml vs mm"}
        ar=S.config.get("active_rules",list(rules_.keys())); nar=[]
        for k,lb in rules_.items():
            if st.checkbox(lb,value=k in ar,key=f"rule_{k}"): nar.append(k)
        st.markdown('<div class="sec">Internet Search</div>',unsafe_allow_html=True)
        st.markdown(f'<div style="font-size:.78rem;color:{"#00e676" if S.api_key else "#8b949e"};">{"● API key set — search enabled" if S.api_key else "○ Enter API key in sidebar to enable"}</div>',unsafe_allow_html=True)
        auto_s=st.toggle("Auto-search REVIEW items during validation",value=S.config.get("auto_search",False),disabled=not S.api_key)
    S.config={"npc_thresh":nt,"phc_thresh":pt,"high_thresh":ht,"med_thresh":mt_,"use_brand":ub,"max_rows":mr,"active_rules":nar,"auto_search":auto_s}
    st.markdown("---"); c1,_,c3=st.columns([1,3,1])
    with c1:
        if st.button("← Back",use_container_width=True): S.step=1; st.rerun()
    with c3:
        if st.button("Run Matching →",use_container_width=True): S.step=3; st.rerun()

def step_process():
    st.markdown('<h1 class="ptitle">Processing</h1>',unsafe_allow_html=True)
    if S.run_done and S.results is not None:
        st.success(f"✓ Done — {len(S.results):,} products matched")
        c1,c2=st.columns(2)
        with c1:
            if st.button("← Reconfigure",use_container_width=True): S.run_done=False; S.step=2; st.rerun()
        with c2:
            if st.button("Review Results →",use_container_width=True): S.step=4; st.rerun()
        return
    cfg=S.config; cm=S.col_map
    df_i=S.files["ihs"]; df_n=S.files["npc"]; df_p=S.files["phc"]
    dc=cm["ihs"]["desc"]; cc=cm["ihs"].get("code")
    nc_=cm["npc"]["code"]; nd_=cm["npc"]["desc"]
    pd__=cm["phc"]["desc"]; pn_=cm["phc"]["npc_code"]; pi_=cm["phc"].get("ihbs_code")
    products=df_i[[c for c in [cc,dc] if c]].dropna(subset=[dc])
    if cfg["max_rows"]>0: products=products.head(cfg["max_rows"])
    total=len(products); logs_=[]; lb=st.empty()
    def log_(msg,kind="info"):
        ts=datetime.now().strftime("%H:%M:%S")
        cls_={"ok":"cok","warn":"cwarn","err":"cerr","info":"cinfo"}.get(kind,"cinfo")
        logs_.append(f'<span class="{cls_}">[{ts}] {msg}</span>')
        lb.markdown('<div class="con">'+("<br>".join(logs_[-22:]))+"</div>",unsafe_allow_html=True)
    log_(f"Indexing {total:,} IHS products...")
    ni=build_npc_idx(df_n,nd_,nc_); log_(f"NPC: {len(ni):,} entries","ok")
    pi=build_phc_idx(df_p,pd__,pn_,pi_); log_(f"PHC: {len(pi):,} entries","ok")
    prog=st.progress(0,"Starting..."); results_=[]
    for i,row in enumerate(products.itertuples(index=False)):
        desc_=str(getattr(row,dc.replace(" ","_"),"")).strip()
        rhic_=str(getattr(row,cc.replace(" ","_"),"")).strip() if cc else f"IHS-{i+1:04d}"
        if not desc_ or desc_.lower()=="nan": continue
        m=match_one(desc_,ni,pi,cfg["npc_thresh"],cfg["phc_thresh"])
        sc_=m["MATCH_SCORE"]
        m["CONFIDENCE"]="HIGH" if sc_>=cfg["high_thresh"] else ("MEDIUM" if sc_>=cfg["med_thresh"] else "LOW")
        m["RHIC_CODE"]=rhic_; results_.append(m)
        if i%50==0 or i==total-1:
            prog.progress((i+1)/total,f"Matching {i+1:,}/{total:,}...")
            if i%300==0 and i>0:
                matched=sum(1 for r in results_ if r["MATCH_SOURCE"]!="UNMATCHED")
                log_(f"  {i+1:,}/{total:,} · matched {matched:,} ({matched/len(results_)*100:.0f}%)")
    prog.progress(1.0,"Complete!")
    df_r=pd.DataFrame(results_)
    front=["RHIC_CODE","ORIGINAL_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION","IHBS_CODE","MATCH_SOURCE","MATCH_SCORE","CONFIDENCE","MATCH_TYPE","VALIDATION_STATUS","VALIDATION_COMMENT","PRODUCT_FAMILY","MANUAL_STATUS","MANUAL_NOTE","EDITED_NPC_CODE"]
    df_r=df_r[[c for c in front if c in df_r.columns]]
    S.results=df_r; S.logs=logs_; S.run_done=True; S.validations={}
    n_rev=(df_r["VALIDATION_STATUS"]=="REVIEW").sum()
    log_(f"NPC: {(df_r['MATCH_SOURCE']=='NPC').sum():,} · PHC: {(df_r['MATCH_SOURCE']=='PHC').sum():,} · Unmatched: {(df_r['MATCH_SOURCE']=='UNMATCHED').sum():,}","ok")
    log_(f"REVIEW items: {n_rev:,}","warn" if n_rev>0 else "ok"); st.success(f"✓ {len(results_):,} processed")
    c1,c2=st.columns(2)
    with c1:
        if st.button("← Reconfigure",use_container_width=True): S.run_done=False; S.step=2; st.rerun()
    with c2:
        if st.button("Review Results →",use_container_width=True): S.step=4; st.rerun()

def step_review():
    st.markdown('<h1 class="ptitle">Results Review</h1>',unsafe_allow_html=True)
    df=S.results
    if df is None: st.warning("No results yet."); return
    total=len(df)
    n_npc=(df["MATCH_SOURCE"]=="NPC").sum(); n_phc=(df["MATCH_SOURCE"]=="PHC").sum()
    n_unm=(df["MATCH_SOURCE"]=="UNMATCHED").sum(); n_rev=(df["VALIDATION_STATUS"]=="REVIEW").sum()
    cov=(n_npc+n_phc)/total*100
    k1,k2,k3,k4,k5,k6=st.columns(6)
    for col,val,lbl,cls in [(k1,str(total),"Total","km"),(k2,f"{n_npc:,}","NPC Matched","kg"),(k3,f"{n_phc:,}","PHC Fallback","kb"),(k4,f"{n_unm:,}","Unmatched","km"),(k5,f"{n_rev:,}","Needs Review","ky"),(k6,f"{cov:.0f}%","Coverage","kg" if cov>=80 else "ky")]:
        with col: st.markdown(f'<div class="kc"><div class="kv {cls}">{val}</div><div class="kl">{lbl}</div></div>',unsafe_allow_html=True)
    st.markdown("")
    f1,f2,f3,f4=st.columns(4)
    fs_=f1.selectbox("Source",["All"]+sorted(df["MATCH_SOURCE"].dropna().unique().tolist()))
    fc_=f2.selectbox("Confidence",["All"]+sorted(df["CONFIDENCE"].dropna().unique().tolist()))
    fv_=f3.selectbox("Validation",["All"]+sorted(df["VALIDATION_STATUS"].dropna().unique().tolist()))
    ff_=f4.selectbox("Family",["All"]+sorted(df["PRODUCT_FAMILY"].dropna().unique().tolist()))
    srch_=st.text_input("🔍 Search","")
    view=df.copy()
    if fs_!="All": view=view[view["MATCH_SOURCE"]==fs_]
    if fc_!="All": view=view[view["CONFIDENCE"]==fc_]
    if fv_!="All": view=view[view["VALIDATION_STATUS"]==fv_]
    if ff_!="All": view=view[view["PRODUCT_FAMILY"]==ff_]
    if srch_: mask=view.apply(lambda r: srch_.upper() in " ".join(str(v).upper() for v in r.values),axis=1); view=view[mask]
    st.caption(f"Showing {len(view):,} / {total:,}")
    show_=["RHIC_CODE","ORIGINAL_DESCRIPTION","NPC_CODE","NPC_DESCRIPTION","MATCH_SOURCE","MATCH_SCORE","CONFIDENCE","MATCH_TYPE","VALIDATION_STATUS","VALIDATION_COMMENT"]
    show_=[c for c in show_ if c in view.columns]
    def srow_(r):
        if r.get("VALIDATION_STATUS","")=="REVIEW": bg="background-color:rgba(248,81,73,.07)"
        elif r.get("MATCH_SOURCE","")=="UNMATCHED": bg="background-color:rgba(240,165,0,.07)"
        elif r.get("MATCH_SOURCE","")=="PHC": bg="background-color:rgba(56,139,253,.06)"
        else: bg="background-color:rgba(0,179,71,.04)"
        return [bg]*len(r)
    st.dataframe(view[show_].style.apply(srow_,axis=1).format({"MATCH_SCORE":"{:.0f}"}),use_container_width=True,height=380)
    ch1,ch2=st.columns(2)
    with ch1:
        st.markdown('<div class="sec" style="font-size:.85rem;">Match Source</div>',unsafe_allow_html=True)
        vc_=df["MATCH_SOURCE"].value_counts().reset_index(); vc_.columns=["Source","Count"]
        st.bar_chart(vc_.set_index("Source"),color="#00b347",height=150)
    with ch2:
        st.markdown('<div class="sec" style="font-size:.85rem;">Confidence</div>',unsafe_allow_html=True)
        vc_=df["CONFIDENCE"].value_counts().reset_index(); vc_.columns=["Confidence","Count"]
        st.bar_chart(vc_.set_index("Confidence"),color="#388bfd",height=150)
    st.markdown("---"); c1,_,c3=st.columns([1,3,1])
    with c1:
        if st.button("← Back to Process",use_container_width=True): S.step=3; st.rerun()
    with c3:
        if st.button("Validate Matches →",use_container_width=True): S.step=5; st.rerun()

def step_validate():
    st.markdown('<h1 class="ptitle">Match Validation</h1>',unsafe_allow_html=True)
    st.markdown('<p class="psub">Confirm, reject, or investigate each match. Internet search provides live clinical verification.</p>',unsafe_allow_html=True)

    df=S.results
    if df is None: st.warning("No results to validate."); return

    def sync():
        for idx,state in S.validations.items():
            if idx<len(S.results):
                S.results.at[idx,"MANUAL_STATUS"]=state.get("status","PENDING")
                S.results.at[idx,"MANUAL_NOTE"]=state.get("note","")
                S.results.at[idx,"EDITED_NPC_CODE"]=state.get("edited_code","")
    sync()

    total=len(df)
    n_pe=(df["MANUAL_STATUS"]=="PENDING").sum(); n_co=(df["MANUAL_STATUS"]=="CONFIRMED").sum()
    n_rj=(df["MANUAL_STATUS"]=="REJECTED").sum(); n_ed=(df["MANUAL_STATUS"]=="EDITED").sum()
    n_done=n_co+n_rj+n_ed

    st.progress(n_done/total if total else 0, text=f"{n_done}/{total} validated ({n_done/total*100:.0f}%)" if total else "0/0")
    k1,k2,k3,k4,k5=st.columns(5)
    for col,val,lbl,cls in [(k1,str(total),"Total","km"),(k2,str(n_co),"Confirmed","kg"),(k3,str(n_rj),"Rejected","kr"),(k4,str(n_ed),"Edited","kp"),(k5,str(n_pe),"Pending","ky")]:
        with col: st.markdown(f'<div class="kc"><div class="kv {cls}">{val}</div><div class="kl">{lbl}</div></div>',unsafe_allow_html=True)
    st.markdown("")

    # Filters
    f1,f2,f3,f4=st.columns(4)
    show_filt=f1.selectbox("Show",["Needs Review","All Items","Pending","Confirmed","Rejected","Edited"])
    f_fam_=f2.selectbox("Family",["All"]+sorted(df["PRODUCT_FAMILY"].dropna().unique().tolist()))
    f_src_=f3.selectbox("Source",["All"]+sorted(df["MATCH_SOURCE"].dropna().unique().tolist()))
    srch__=f4.text_input("🔍 Search","")
    page_sz=st.select_slider("Items per page",[10,25,50,100],value=25)

    view=df.copy()
    if show_filt=="Needs Review": view=view[view["VALIDATION_STATUS"]=="REVIEW"]
    elif show_filt=="Pending":    view=view[view["MANUAL_STATUS"]=="PENDING"]
    elif show_filt=="Confirmed":  view=view[view["MANUAL_STATUS"]=="CONFIRMED"]
    elif show_filt=="Rejected":   view=view[view["MANUAL_STATUS"]=="REJECTED"]
    elif show_filt=="Edited":     view=view[view["MANUAL_STATUS"]=="EDITED"]
    if f_fam_!="All": view=view[view["PRODUCT_FAMILY"]==f_fam_]
    if f_src_!="All": view=view[view["MATCH_SOURCE"]==f_src_]
    if srch__:
        mask=view.apply(lambda r: srch__.upper() in " ".join(str(v).upper() for v in r.values),axis=1); view=view[mask]

    tv=len(view); st.caption(f"Showing {min(page_sz,tv):,} of {tv:,} filtered items")

    # Bulk actions
    with st.expander("⚡ Bulk Actions"):
        bc1,bc2,bc3=st.columns(3)
        if bc1.button("✓ Confirm all HIGH confidence",use_container_width=True):
            for idx in view[view["CONFIDENCE"]=="HIGH"].index:
                S.validations[idx]={"status":"CONFIRMED","note":"Bulk confirmed (HIGH)","edited_code":""}
            sync(); st.rerun()
        if bc2.button("✓ Confirm all EXACT matches",use_container_width=True):
            for idx in view[view["MATCH_TYPE"]=="EXACT"].index:
                S.validations[idx]={"status":"CONFIRMED","note":"Bulk confirmed (EXACT)","edited_code":""}
            sync(); st.rerun()
        if bc3.button("Reset visible to PENDING",use_container_width=True):
            for idx in view.index:
                S.validations[idx]={"status":"PENDING","note":"","edited_code":""}
            sync(); st.rerun()

    # Auto-search
    if S.config.get("auto_search") and S.api_key and show_filt in ("Needs Review","Pending"):
        ap=view[view["MANUAL_STATUS"]=="PENDING"].head(3)
        if len(ap)>0:
            with st.spinner(f"Auto-searching {len(ap)} items..."):
                for _,row_ in ap.iterrows():
                    ck=f"{row_['ORIGINAL_DESCRIPTION']}|||{row_['NPC_DESCRIPTION']}"
                    if ck not in S.search_cache: search_product(row_["ORIGINAL_DESCRIPTION"],row_["NPC_DESCRIPTION"],S.api_key)

    st.markdown("---")

    # Per-item cards
    for idx,row in view.head(page_sz).iterrows():
        ms=S.validations.get(idx,{"status":row.get("MANUAL_STATUS","PENDING"),"note":"","edited_code":""})
        mst=ms.get("status","PENDING")

        # Pill classes
        mp_map={"EXACT":"pE","BRAND_MATCH":"pB","SPEC_DIFF":"pS","NEW_SOP":"pN"}
        vp_map={"VALID":"pVA","REVIEW":"pRV"}
        man_map={"CONFIRMED":"pCO","REJECTED":"pRJ","EDITED":"pED","PENDING":"pPE"}
        mp_mt=f'<span class="mp {mp_map.get(row["MATCH_TYPE"],"pN")}">{row["MATCH_TYPE"]}</span>'
        mp_vs=f'<span class="mp {vp_map.get(row["VALIDATION_STATUS"],"pPE")}">{row["VALIDATION_STATUS"]}</span>'
        mp_ms=f'<span class="mp {man_map.get(mst,"pPE")}">{mst}</span>'

        has_result=f"{row['ORIGINAL_DESCRIPTION']}|||{row['NPC_DESCRIPTION']}" in S.search_cache
        sc_icon="🔬" if has_result else "🔍"
        sc_lbl="View Search Result" if has_result else "Investigate Online"

        st.markdown(f"""<div class="vc s{mst[:2]}">
<div class="vr">{row.get("RHIC_CODE","—")} &nbsp;·&nbsp; {row.get("PRODUCT_FAMILY","—")} &nbsp;·&nbsp; Score: {int(row.get("MATCH_SCORE",0))} &nbsp;·&nbsp; {mp_mt} {mp_vs} {mp_ms}</div>
<div class="vd">{row["ORIGINAL_DESCRIPTION"]}</div>
<div class="vn">→ <strong>{row["NPC_CODE"]}</strong> &nbsp; {row["NPC_DESCRIPTION"]}</div>
{"<div class='vw'>⚠ "+str(row['VALIDATION_COMMENT'])+"</div>" if str(row.get("VALIDATION_COMMENT","")) else ""}
</div>""",unsafe_allow_html=True)

        ba,bb,bc,bd=st.columns([1,1,1,2])
        with ba:
            if st.button("✓ Confirm",key=f"co_{idx}",use_container_width=True):
                S.validations[idx]={**ms,"status":"CONFIRMED"}; sync(); st.rerun()
        with bb:
            if st.button("✗ Reject",key=f"rj_{idx}",use_container_width=True):
                S.validations[idx]={**ms,"status":"REJECTED"}; sync(); st.rerun()
        with bc:
            if st.button("✏ Edit",key=f"ed_{idx}",use_container_width=True):
                st.session_state[f"em_{idx}"]=not st.session_state.get(f"em_{idx}",False)
        with bd:
            if st.button(f"{sc_icon} {sc_lbl}",key=f"sc_{idx}",use_container_width=True,disabled=(not S.api_key and not has_result)):
                st.session_state[f"so_{idx}"]=not st.session_state.get(f"so_{idx}",False)

        # Edit panel
        if st.session_state.get(f"em_{idx}",False):
            ea,eb=st.columns([2,1])
            nc_ed=ea.text_input("New NPC Code",value=ms.get("edited_code","") or row["NPC_CODE"],key=f"ec_{idx}")
            nn_=ea.text_input("Reason / note",value=ms.get("note",""),key=f"nn_{idx}")
            if eb.button("💾 Save",key=f"sv_{idx}",use_container_width=True):
                S.validations[idx]={"status":"EDITED","note":nn_,"edited_code":nc_ed}
                st.session_state[f"em_{idx}"]=False; sync(); st.rerun()

        # Search panel
        if st.session_state.get(f"so_{idx}",False):
            ck=f"{row['ORIGINAL_DESCRIPTION']}|||{row['NPC_DESCRIPTION']}"
            if ck in S.search_cache:
                render_result(S.search_cache[ck])
            else:
                if not S.api_key:
                    st.warning("Enter your Anthropic API key in the sidebar to enable internet search.")
                else:
                    with st.spinner(f"Searching: {row['ORIGINAL_DESCRIPTION'][:55]}..."):
                        res_=search_product(row["ORIGINAL_DESCRIPTION"],row["NPC_DESCRIPTION"],S.api_key)
                    render_result(res_)
                    if "match_assessment" in res_:
                        assess_=res_["match_assessment"]
                        if assess_ in ("CORRECT","LIKELY CORRECT"):
                            st.info("💡 Search suggests this match is correct — click ✓ Confirm to accept.")
                        elif assess_ in ("LIKELY INCORRECT","INCORRECT"):
                            st.warning("⚠️ Search flagged a likely mismatch — consider ✗ Reject or ✏ Edit Code.")
        st.markdown("")

    if tv>page_sz: st.info(f"Showing {page_sz} of {tv}. Use filters or increase page size to see more.")
    st.markdown("---"); c1,_,c3=st.columns([1,3,1])
    with c1:
        if st.button("← Back to Review",use_container_width=True): S.step=4; st.rerun()
    with c3:
        if st.button("Export Results →",use_container_width=True): sync(); S.step=6; st.rerun()

def step_export():
    st.markdown('<h1 class="ptitle">Export</h1>',unsafe_allow_html=True)
    st.markdown('<p class="psub">Download the complete harmonized and validated dataset.</p>',unsafe_allow_html=True)
    df=S.results
    if df is None: st.warning("No results to export."); return
    total=len(df)
    n_co=(df["MANUAL_STATUS"]=="CONFIRMED").sum(); n_rj=(df["MANUAL_STATUS"]=="REJECTED").sum()
    n_ed=(df["MANUAL_STATUS"]=="EDITED").sum(); n_pe=(df["MANUAL_STATUS"]=="PENDING").sum()
    n_rev=(df["VALIDATION_STATUS"]=="REVIEW").sum()
    cov=((df["MATCH_SOURCE"]=="NPC").sum()+(df["MATCH_SOURCE"]=="PHC").sum())/total*100
    k1,k2,k3,k4,k5,k6=st.columns(6)
    for col,val,lbl,cls in [(k1,str(total),"Total","km"),(k2,str(n_co),"Confirmed","kg"),(k3,str(n_rj),"Rejected","kr"),(k4,str(n_ed),"Edited","kp"),(k5,str(n_pe),"Pending","ky"),(k6,f"{cov:.0f}%","Coverage","kg" if cov>=80 else "ky")]:
        with col: st.markdown(f'<div class="kc"><div class="kv {cls}">{val}</div><div class="kl">{lbl}</div></div>',unsafe_allow_html=True)
    if n_rev>0: st.warning(f"⚠ {n_rev:,} items still flagged REVIEW. {n_pe:,} items still PENDING manual validation.")
    st.markdown("")
    for lbl in ["📋 Harmonized Mapping — all products with NPC code, IHBS code, confidence, manual validation status","📊 Validation Summary — match source, manual status, validation rule, match type breakdown"]:
        st.markdown(f'<div style="padding:.4rem .6rem;border-left:2px solid #00b347;margin-bottom:.3rem;font-size:.83rem;">{lbl}</div>',unsafe_allow_html=True)
    st.markdown("")
    with st.spinner("Building Excel workbook..."):
        xlsx=build_excel(df)
    fname=f"IHS_NPC_Validated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button("⬇ Download Validated Excel",data=xlsx,file_name=fname,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.download_button("⬇ Download CSV",data=df.to_csv(index=False).encode(),file_name=fname.replace(".xlsx",".csv"),mime="text/csv")
    st.markdown("---")
    st.markdown(f'<div style="font-size:.76rem;color:#8b949e;text-align:center;">Generated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} · Rwanda FDA · SOP ODDG/RES/SOP/004 · IHS–NPC Harmonizer v3</div>',unsafe_allow_html=True)
    if st.button("↩ New Session"):
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

# ── Router ────────────────────────────────────────────────────────────────────
[step_upload,step_map,step_configure,step_process,step_review,step_validate,step_export][min(S.step,6)]()
