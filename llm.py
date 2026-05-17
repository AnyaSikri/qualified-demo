"""OpenAI GPT-4o wrapper for per-section narrative synthesis.

One section = one LLM call. The call returns:
  {"text": <string>, "sources": {"adae": [...], "protocol": [...]}}

The `sources` block is built mechanically from the retrieval inputs — the
model is NOT asked to cite sources itself. This guarantees that every
citation is grounded in actual retrieved context, not hallucinated.
"""

import json
import os

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

MODEL = "gpt-4o"

SYSTEM_PROMPT = """You are a regulatory medical writer drafting Patient Safety Narratives \
(PSN) for FDA submission. You write in formal clinical prose suitable for a \
"Dear Investigator" letter. Rules:

- Use only the facts provided in the CONTEXT block. Do not invent dates, \
doses, severities, MedDRA terms, lab values, or outcomes.
- Preserve MedDRA preferred terms verbatim (e.g., "Pneumonia bacterial", \
not "bacterial pneumonia").
- If a requested field is listed in MISSING DATA, write "not reported" \
rather than guessing.
- Write in the third person, past tense. Refer to the individual as \
"the subject."
- Output ONLY the prose for the requested section. No section heading, \
no preamble, no bullet points unless the instructions request a list.
- Be concise. Follow the section's word-count guidance in INSTRUCTIONS."""


def _build_user_prompt(section_config, adae_context, protocol_context):
    """Assemble the user message. Pure function — easy to unit-test / inspect."""
    lines = [f"SECTION: {section_config['title']}", ""]
    lines.append("INSTRUCTIONS:")
    lines.append(section_config["instructions"])
    lines.append("")

    lines.append("CONTEXT — ADAE rows (clinical trial safety data):")
    if not adae_context:
        lines.append("  (no ADAE rows retrieved)")
    for rec in adae_context:
        lines.append(f"  row {rec['row']}:")
        if rec["values"]:
            for k, v in rec["values"].items():
                lines.append(f"    {k}: {v}")
        else:
            lines.append("    (all requested fields blank for this row)")
        if rec["missing"]:
            lines.append(f"    MISSING DATA: {', '.join(rec['missing'])}")
    lines.append("")

    lines.append("CONTEXT — Protocol sections:")
    if not protocol_context:
        lines.append("  (no protocol sections retrieved for this narrative section)")
    for sec in protocol_context:
        lines.append(f"  § {sec['number']}. {sec['title']} (pages {sec['pages']}):")
        for line in sec["text"].splitlines():
            if line.strip():
                lines.append(f"    {line.strip()}")
    lines.append("")
    lines.append("Now write the section prose:")
    return "\n".join(lines)


def _build_sources(section_config, adae_context, protocol_context):
    """Build the traceability payload deterministically from retrieval output."""
    return {
        "adae": [
            {
                "row": rec["row"],
                "columns_used": [c for c in rec["columns_used"] if c in rec["values"]],
                "values": rec["values"],
                "missing": rec["missing"],
            }
            for rec in adae_context
        ],
        "protocol": [
            {"section_title": sec["title"], "pages": sec["pages"]}
            for sec in protocol_context
        ],
    }


def generate_section(section_config, adae_context, protocol_context, client=None):
    """Call GPT-4o for one section and return text + structured sources."""
    user_prompt = _build_user_prompt(section_config, adae_context, protocol_context)
    sources = _build_sources(section_config, adae_context, protocol_context)

    client = client or OpenAI()
    response = client.chat.completions.create(
        model=MODEL,
        temperature=0.2,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    )
    text = response.choices[0].message.content.strip()

    return {
        "section_id": section_config["id"],
        "title": section_config["title"],
        "text": text,
        "sources": sources,
    }


# ---------------------------------------------------------------------------
# Standalone smoke test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from template import NARRATIVE_TEMPLATE

    section = NARRATIVE_TEMPLATE["sections"][1]  # event_description

    adae_context = [
        {
            "row": 0,
            "columns_used": ["AEDECOD", "AESEV", "AESTDTC", "AEENDTC", "AESER", "AETOXGR"],
            "values": {
                "AEDECOD": "Pneumonia bacterial",
                "AESEV": "SEVERE",
                "AESTDTC": "2025-03-14",
                "AEENDTC": "2025-03-29",
                "AESER": "Y",
                "AETOXGR": 3,
            },
            "missing": [],
        },
    ]
    protocol_context = [
        {
            "number": 6,
            "title": "Adverse Event Definitions",
            "pages": [7],
            "text": "An adverse event (AE) is any untoward medical occurrence...\n"
                    "A serious adverse event (SAE) is any AE that results in death, "
                    "is life threatening, requires inpatient hospitalization...",
        },
    ]

    print("=== assembled prompt ===")
    print(_build_user_prompt(section, adae_context, protocol_context))
    print("\n=== deterministic sources payload ===")
    print(json.dumps(_build_sources(section, adae_context, protocol_context), indent=2))

    if os.getenv("OPENAI_API_KEY"):
        print("\n=== live GPT-4o call ===")
        result = generate_section(section, adae_context, protocol_context)
        print(result["text"])
    else:
        print("\n(OPENAI_API_KEY not set — skipping live call. "
              "Copy .env.example to .env and add your key to exercise the API path.)")
