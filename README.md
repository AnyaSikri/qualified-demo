# PSN Generator

Patient Safety Narrative generator for clinical trial subjects. Given a subject ID, ADAE Excel file, and protocol PDF, produces a multi-section draft safety letter with full traceability — every sentence maps back to a source row in the Excel or a section in the PDF.

## Architecture (three pillars)

1. **RAG (context retrieval)** — pulls relevant ADAE rows and protocol sections per narrative section. Structural retrieval over the protocol PDF (no vector store — single document, deterministic citations required).
2. **AI Synthesis** — GPT-4o generates prose per section, grounded on retrieved context.
3. **Traceability** — every section returns structured source pointers (ADAE row + columns, protocol section + pages).

## Quickstart

```bash
pip install -r requirements.txt
cp .env.example .env  # add your OPENAI_API_KEY
python build_synthetic_data.py
streamlit run app.py
```

## File map

- `app.py` — Streamlit UI
- `generator.py` — Orchestration (loops sections, calls retrieval + LLM)
- `template.py` — PSN template config
- `data_retrieval/adae.py` — Excel retrieval
- `data_retrieval/protocol.py` — PDF section retrieval
- `llm.py` — OpenAI call wrapper
- `docx_builder.py` — Word + JSON output
- `build_synthetic_data.py` — Generates synthetic ADAE.xlsx and mock_protocol.pdf
