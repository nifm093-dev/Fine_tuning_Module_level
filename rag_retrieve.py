"""
RAG Retrieval Pipeline - All NSS Themes
Outputs: nss_retrieved.json  (theme → top-3 chunks + metadata)
"""

import pandas as pd
import numpy as np
import faiss
import json
import re
from sentence_transformers import SentenceTransformer, util

# ── CONFIG ──────────────────────────────────────────────────────────────────
CSV_PATH   = "/content/uwaterloo_2000_classified.csv"
OUTPUT_JSON = "nss_retrieved.json"
TOP_K      = 10   # candidates from FAISS
FINAL_K    = 3    # kept per theme (re-ranked by cosine sim)

NSS_THEMES = [
    "The teaching on my course",
    "Learning opportunities",
    "Assessment and feedback",
    "Academic support",
    "Organisation and management",
    "Learning resources",
    "Student voice",
    "Student union",
    "Mental wellbeing",
    "Freedom of expression",
    "Academic staff and support",
    "Covid-19 pandemic",
]

# ── LOAD ────────────────────────────────────────────────────────────────────
df = pd.read_csv(CSV_PATH)
rag_chunks = [
    f"Course: {row['course_code']} | Review: {row['review_text']} | Themes: {row['nss_labels']}"
    for _, row in df.iterrows()
]
print(f"✅ Loaded {len(rag_chunks)} chunks")

# ── EMBED ───────────────────────────────────────────────────────────────────
model = SentenceTransformer("BAAI/bge-small-en-v1.5")
embeddings = model.encode(rag_chunks, batch_size=64).astype("float32")
index = faiss.IndexFlatL2(embeddings.shape[1])
index.add(embeddings)
print("✅ FAISS index built")

# ── RETRIEVE ─────────────────────────────────────────────────────────────────
def retrieve(theme, top_k=TOP_K, final_k=FINAL_K):
    query = f"{theme.lower()} student feedback university experience"
    q_emb = model.encode([query]).astype("float32")
    D, I  = index.search(q_emb, top_k)

    if len(I[0]) == 0 or I[0][0] == -1:
        print(f"  ⚠️  No matches for '{theme}'")
        return rag_chunks[:final_k]

    ret_embs = embeddings[I[0]]
    sims = util.cos_sim(q_emb, ret_embs)[0].cpu().numpy()
    top_idx = np.argsort(sims)[-final_k:][::-1]

    chunks = []
    for idx in top_idx:
        raw = rag_chunks[I[0][idx]]
        # parse fields back out
        m = re.match(r"Course: (.+?) \| Review: (.+?) \| Themes: (.+)", raw)
        chunks.append({
            "course_code": m.group(1).strip() if m else "N/A",
            "review_text": m.group(2).strip() if m else raw,
            "nss_labels":  m.group(3).strip() if m else "",
            "similarity":  round(float(sims[idx]), 4),
        })

    best = chunks[0]["similarity"]
    print(f"  🎯 '{theme}' → best sim {best:.3f} | "
          f"top course: {chunks[0]['course_code']}")
    return chunks

# ── RUN ALL THEMES ───────────────────────────────────────────────────────────
results = {}
print("\n📋 Retrieving all NSS themes...\n")
for theme in NSS_THEMES:
    results[theme] = retrieve(theme)

# ── SAVE ────────────────────────────────────────────────────────────────────
with open(OUTPUT_JSON, "w") as f:
    json.dump(results, f, indent=2)

print(f"\n✅ Saved → {OUTPUT_JSON}")
print("▶️  Next: node generate_report.js")
