
# NSS RAG Pipeline for Waterloo Student Surveys

Automated report generation pipeline using **semantic retrieval** (bge-small-en-v1.5) on 2000 classified Waterloo course reviews → top-3 chunks per NSS theme → Llama prompts → DOCX reports.

✅ **Supervisor-ready**: Processes 2000 chunks, cosine similarity 0.772-0.861, no errors.

## 🎯 Files

| File | Purpose |
|------|---------|
| `rag_retrieve.py` | **Core RAG**: Embeds Waterloo CSV → retrieves top-3 chunks per NSS theme (Academic staff, Assessment, Covid-19) by cosine similarity |
| `generate_report.js` | **DOCX Export**: Node.js script reads `nss_retrieved.json` → formats professional NSS report |
| `uwaterloo_2000_classified.csv` | **Dataset**: 2000 Waterloo reviews w/ columns `course_code`, `review_text`, `nss_labels` (multi-label) |
| `NSS_Report.docx` | **Sample Output**: Generated report showing STAT 230 "Meh" (0.772 sim), ECON 101 "ez 95" (0.800 sim) |
| `package.json` | Node.js dependencies: `docx` for Word generation |

## 🚀 Quickstart

```bash
# Python RAG (Colab/T4 GPU)
pip install sentence-transformers python-docx pandas
python rag_retrieve.py  # → nss_retrieved.json

# Node.js DOCX
npm install
node generate_report.js nss_retrieved.json  # → NSS_Report.docx
```

## 🏆 Key Results

```
✅ Academic staff: 0.772 sim → STAT 230 "Meh", ECON 101 topics
✅ Assessment: 0.800 sim → ECON 101 "ez 95 w testbank", CS 135 breeze  
✅ Covid-19: 0.861 sim → PD 1 "fk pd", CS 240 "Like"
```

## 📊 Technical Stack
- **Embeddings**: BAAI/bge-small-en-v1.5 (133M params)
- **Retrieval**: Cosine similarity, top-k=3 per NSS theme
- **Prompts**: Ready for Unsloth LoRA Llama inference
- **Export**: Professional DOCX via docx.js

## 🔬 Research Context
**PhD Brunel University**: Automated NSS analysis pipeline (2-stage: NSS → Waterloo training). Addresses "report generation... you still have a kind of sample" → semantic top-3 selection.

## Next: Llama Integration
```python
# Paste RAG prompts directly to fine-tuned model
"NSS Academic staff and support:
Course: STAT 230 | Review: Meh | Themes: Academic staff and support..."
```

**Live Demo**: [Colab RAG Results](https://colab.research.google.com/drive/...)  
**Norah Alshahrani | AI PhD Researcher** | March 2026
