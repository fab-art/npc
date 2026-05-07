# MedMatch AI — Medical Standardization Engine
### v2.0  ·  Rwanda FDA  ·  SOP ODDG/RES/SOP/004

---

## Quick Start
```bash
pip install -r requirements.txt
streamlit run app.py
```

## HF API Key (enables full semantic mode)
Add to `~/.streamlit/secrets.toml`:
```toml
HF_TOKEN = "hf_..."       # Free at huggingface.co
```
Without it: TF-IDF + RapidFuzz mode (no deps, works offline).

## Architecture
```
Input → Normalize → Brand Alias → Embed (HF API / TF-IDF)
     → Cosine Similarity (Top 20) → Re-rank (Cross-encoder / RapidFuzz)
     → Hybrid Score (70% sem + 30% rerank) → Rule Validation → Top N Output
```

## Files
- `app.py`    — Streamlit frontend
- `engine.py` — Core pipeline (embedding, matching, validation)

## Deployment
```bash
# Docker
docker build -t medmatch .
docker run -p 8501:8501 medmatch

# Streamlit Cloud — push to GitHub, connect at share.streamlit.io
```
