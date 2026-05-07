# MedMatch AI — Deployment Guide
## Rwanda FDA · Medical Standardization Engine v2.0

---

## Quick Start

```bash
# Install all dependencies
pip install -r requirements.txt

# Run locally
streamlit run app.py
```
Opens at **http://localhost:8501**

---

## Hosting

### Streamlit Cloud (recommended — free)
1. Push to GitHub
2. share.streamlit.io → New app → select repo
3. Add secrets (Settings → Secrets — see below)

### Docker
```dockerfile
FROM python:3.11-slim
RUN apt-get update && apt-get install -y gcc g++ && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

### Nginx Reverse Proxy
```nginx
server {
    listen 80;
    server_name medmatch.your-domain.com;
    location / {
        proxy_pass http://localhost:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_read_timeout 120s;
    }
}
```

---

## API Keys (~/.streamlit/secrets.toml)

```toml
# openFDA — optional, raises limit from 1k to 120k req/day
# Get free key at: https://api.fda.gov/
OPENFDA_KEY = "your_openfda_key"

# HuggingFace — only needed if sentence-transformers can't install locally
# Get free token at: https://huggingface.co/settings/tokens
HF_TOKEN = "your_hf_token"
```

---

## Engine Tiers (auto-selected by availability)

| Tier | Engine | Install | Quality | Speed |
|------|--------|---------|---------|-------|
| 1 | **Sentence Transformers** (all-MiniLM-L6-v2) | `pip install sentence-transformers` | ⭐⭐⭐⭐⭐ | ~50ms |
| 1+ | **Cross-Encoder Re-Ranker** (ms-marco-MiniLM) | included above | ⭐⭐⭐⭐⭐ | +20ms |
| 1.5 | **PubMedBERT** (medical domain) | optional | ⭐⭐⭐⭐⭐ | ~80ms |
| 2 | **HF Inference API** (remote) | `HF_TOKEN` in secrets | ⭐⭐⭐⭐ | ~200ms |
| 3 | **TF-IDF** (bag of words) | `pip install scikit-learn` | ⭐⭐⭐ | ~5ms |
| 4 | **RapidFuzz** (edit distance) | `pip install rapidfuzz` | ⭐⭐ | ~2ms |

---

## System Architecture

```
User Input (description)
    │
    ▼
[1] NormalizationEngine
    - lowercase + clean
    - brand → generic alias (50+ mappings)
    - spec extraction (mm, ml, mg, FR, holes)
    │
    ▼
[2] EmbeddingEngine
    - Sentence Transformer → 384-dim vector
    - OR: HF Inference API
    - OR: TF-IDF matrix
    │
    ▼
[3] CandidateRetrieval
    - cosine similarity vs catalog embeddings
    - top-20 candidates returned
    │
    ▼
[4] CrossEncoder Re-Ranking
    - ms-marco-MiniLM-L-6-v2
    - deep pairwise scoring
    - top-5 re-ranked
    │
    ▼
[5] Hybrid Scoring
    - 70% semantic + 30% fuzzy
    │
    ▼
[6] ValidationEngine (rule-based safety)
    - anatomy incompatibility check
    - product family mismatch
    - numeric spec tolerance (±0.1mm)
    - drug dose check
    │
    ▼
[7] EnrichmentEngine (openFDA)
    - device classification
    - drug label lookup
    - cached in session
    │
    ▼
Output:
    - NPC Code + Description
    - Confidence Score (0-100%)
    - Match Type (EXACT/BRAND_MATCH/SPEC_DIFF/NEW_SOP)
    - Validation Status (VALID/REVIEW)
    - openFDA metadata
```

---

## Memory Requirements

| Catalog Size | Embeddings RAM | Total App RAM |
|-------------|---------------|--------------|
| 1,000 items | ~1.5 MB | ~500 MB |
| 5,000 items | ~7.5 MB | ~700 MB |
| 10,000 items | ~15 MB | ~900 MB |
| 50,000 items | ~75 MB | ~1.5 GB |

Runs comfortably on HuggingFace Spaces (2 vCPU, 16 GB RAM).
