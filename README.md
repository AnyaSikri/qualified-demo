# PSN Generator

Drafts a regulatory **Patient Safety Narrative** (PSN — the "letter to FDA" describing a single clinical trial subject's safety event) from two inputs: an SDTM-style ADAE Excel file and a study protocol PDF. Every generated sentence maps back to a source row in the Excel or a section in the PDF.

Built as a demo for clinical-trial-grade AI synthesis where **traceability is non-negotiable**.

---

## The three architectural pillars

**1. RAG (context retrieval) — `data_retrieval/`**

For each narrative section, pull the relevant ADAE rows + relevant protocol sections.

- **ADAE retrieval** (`data_retrieval/adae.py`): load the workbook, filter by subject, project the columns the template asks for, and normalize NaN / blank values into an explicit `missing` list. Source row index travels with every value so citations are mechanical, not hallucinated.
- **Protocol retrieval** (`data_retrieval/protocol.py`): parse the PDF once with PyMuPDF, walk numbered headers (`^\d+\.\s+Title$`) by regex, slice each section's body + spanning page numbers. **No vector store** — for a single well-structured document where citations must point at specific pages, structural retrieval gives 100% precise, deterministic references with no embedding store to maintain.

**2. AI synthesis — `llm.py`**

One LLM call per template section (4 total). The prompt explicitly carries:
- the section's `instructions` field as the task
- the retrieved ADAE values (with a `MISSING DATA:` line for any blank fields, so the model writes "not reported" instead of hallucinating)
- the retrieved protocol section bodies

System prompt enforces verbatim MedDRA terms, third-person past tense, formal regulatory tone, and the rule "use only the facts in CONTEXT." Model: **GPT-4o** at `temperature=0.2`.

**3. Traceability — built into the data model**

Every generated section returns the same shape:

```python
{
  "section_id": "event_description",
  "title": "Event Description",
  "text": "On Day 47, the subject developed a Grade 3 serious adverse event...",
  "sources": {
    "adae": [
      {"row": 0, "columns_used": ["AEDECOD", "AESEV", ...],
       "values": {"AEDECOD": "Pneumonia bacterial", ...},
       "missing": []}
    ],
    "protocol": [
      {"section_title": "Adverse Event Definitions", "pages": [7]}
    ]
  }
}
```

The `sources` block is **built from the retrieval inputs, not from model output** — the LLM is never asked to cite. This guarantees every citation is grounded in actual retrieved context.

---

## How to run

```bash
pip install -r requirements.txt
cp .env.example .env          # then put your OPENAI_API_KEY in it
python build_synthetic_data.py
streamlit run app.py
```

Or end-to-end from the CLI:

```bash
python generator.py SUBJ-1042
python docx_builder.py SUBJ-1042   # builds narrative.docx + sources.json
```

Try `SUBJ-0817` to see the gap-flagging demo — that subject has intentionally missing CMDOSE/CMSTDTC fields, and the output writes "not reported" while the citation list omits those columns.

---

## File map

| File | What it does |
| --- | --- |
| `app.py` | Streamlit UI — inputs, template viewer, narrative + sources columns, downloads |
| `generator.py` | Orchestration. Loops template sections, calls retrieval + LLM for each. CLI entrypoint via `python generator.py <SUBJID>` |
| `template.py` | PSN template config — sections, instructions, ADAE field projections, protocol section refs |
| `data_retrieval/adae.py` | ADAE Excel loader, subject filter, per-section field projector with missing-value flagging |
| `data_retrieval/protocol.py` | Protocol PDF loader, numbered-section parser, title lookup with page references |
| `llm.py` | GPT-4o wrapper. `generate_section()` assembles prompt + deterministic source payload |
| `docx_builder.py` | Word document + sources JSON builders |
| `build_synthetic_data.py` | Generates `data/ADAE.xlsx` (10 rows, 8 subjects) and `data/mock_protocol.pdf` (8 numbered sections) |
| `data/` | Generated synthetic inputs and per-subject outputs (gitignored) |

---

## Synthetic data

All inputs are synthetic. The ADAE workbook follows SDTM column conventions (`USUBJID`, `AEDECOD`, `AETOXGR`, `AEACN`, `AEOUT`, `CMTRT`, etc.) across 8 subjects with deliberately varied profiles — Grade 3 pneumonia (drug interrupted, resolved), Grade 5 sepsis (fatal, missing conmed data), Grade 4 neutropenia (dose reduced), and 5 filler subjects with diverse grades and outcomes. The protocol PDF is a mock Phase 2 oncology trial protocol with 8 numbered sections suitable for regex-based section extraction.

---

## Design notes

- **No streaming, no concurrency.** One section, one LLM call, sequential. Keeps the source-tracking model trivially correct and makes the demo log readable.
- **Source citations are mechanical.** The model is never asked "what are your sources?" The sources block is built from the retrieval inputs the caller passed in. This is the only way to guarantee citation integrity in a regulatory context.
- **Structural retrieval over semantic.** A vector store would add infrastructure for no precision gain on a single 8-page document with numbered headers. Worse, embeddings can't promise "this sentence came from page 7" — the regex-based slicer can.
- **Missing-data path is explicit.** NaN / empty cells become a `missing` array, surfaced to the LLM as `MISSING DATA: <fields>`. The system prompt requires "not reported" rather than a guess.
