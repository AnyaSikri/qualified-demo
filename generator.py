"""Orchestration — loops template sections, calls retrieval + LLM, returns
the full structured narrative.

This is the single seam the Streamlit UI and CLI both talk to. Retrieval is
sequential (fast, pure pandas/PDF). The 4 GPT-4o calls fan out in parallel —
each section's prompt is independent of every other section, so we wait
roughly max(individual call) instead of sum.
"""

import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from openai import OpenAI

from data_retrieval import adae as adae_mod
from data_retrieval import protocol as protocol_mod
from llm import generate_section


def _default_log(msg):
    print(msg, flush=True)


def generate_narrative(subject_id, adae_path, protocol_path, template, log=None):
    """Produce a full PSN for one subject.

    Returns:
        {
          "subject_id": "...",
          "template": {"name": "...", "version": "..."},
          "sections": [
            {"section_id": "...", "title": "...", "text": "...",
             "sources": {"adae": [...], "protocol": [...]}},
            ...
          ]
        }
    """
    log = log or _default_log

    log(f"[load] ADAE      ← {adae_path}")
    df = adae_mod.load_adae(adae_path)
    log(f"       {len(df)} rows, {df['USUBJID'].nunique()} subjects")

    log(f"[load] protocol  ← {protocol_path}")
    pages = protocol_mod.load_protocol(protocol_path)
    parsed_sections = protocol_mod.parse_sections(pages)
    log(f"       {len(pages)} pages, {len(parsed_sections)} sections parsed")

    log(f"[filter] subject={subject_id}")
    events = adae_mod.get_subject_events(df, subject_id)
    if events.empty:
        raise ValueError(f"No ADAE rows found for subject {subject_id!r}")
    log(f"         {len(events)} event row(s) for this subject")

    # --- Retrieval (sequential, fast) -------------------------------------
    contexts = []
    for section_config in template["sections"]:
        adae_context = adae_mod.get_fields_for_section(events, section_config)
        protocol_context = protocol_mod.get_sections_for(
            parsed_sections, section_config["protocol_sections"]
        )
        rows = [rec["row"] for rec in adae_context]
        prot_labels = [f"§{p['number']} {p['title']} p.{p['pages']}"
                       for p in protocol_context]
        log(f"[retrieve {section_config['id']}] "
            f"ADAE rows={rows}; protocol={prot_labels or '(none)'}")
        contexts.append((section_config, adae_context, protocol_context))

    # --- Synthesis (parallel) ---------------------------------------------
    client = OpenAI()  # shared across all 4 calls
    n = len(contexts)
    log(f"[synth] firing {n} GPT-4o calls in parallel...")
    t0 = time.perf_counter()

    results_by_id = {}
    with ThreadPoolExecutor(max_workers=n) as pool:
        futures = {
            pool.submit(
                generate_section, section_config, adae_ctx, prot_ctx, client
            ): section_config["id"]
            for section_config, adae_ctx, prot_ctx in contexts
        }
        for future in as_completed(futures):
            sid = futures[future]
            t_section = time.perf_counter() - t0
            result = future.result()
            results_by_id[sid] = result
            log(f"  ✓ {sid}  ({len(result['text'])} chars, t+{t_section:.1f}s)")

    elapsed = time.perf_counter() - t0
    log(f"[synth] all {n} sections done in {elapsed:.1f}s wall-clock")

    # Reorder results to template section order.
    sections_out = [results_by_id[s["id"]] for s in template["sections"]]

    return {
        "subject_id": subject_id,
        "template": {"name": template["name"], "version": template["version"]},
        "sections": sections_out,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from template import NARRATIVE_TEMPLATE

    here = os.path.dirname(os.path.abspath(__file__))
    default_adae = os.path.join(here, "data", "ADAE.xlsx")
    default_protocol = os.path.join(here, "data", "mock_protocol.pdf")

    subject_id = sys.argv[1] if len(sys.argv) > 1 else "SUBJ-1042"
    adae_path = sys.argv[2] if len(sys.argv) > 2 else default_adae
    protocol_path = sys.argv[3] if len(sys.argv) > 3 else default_protocol

    if not os.getenv("OPENAI_API_KEY"):
        sys.exit(
            "OPENAI_API_KEY is not set. Copy .env.example to .env and add "
            "your key, then re-run."
        )

    narrative = generate_narrative(
        subject_id, adae_path, protocol_path, NARRATIVE_TEMPLATE
    )

    print("\n" + "=" * 72)
    print(f"PSN — subject {narrative['subject_id']}")
    print(f"Template: {narrative['template']['name']} v{narrative['template']['version']}")
    print("=" * 72)
    for section in narrative["sections"]:
        print(f"\n## {section['title']}")
        print(section["text"])
        adae_cites = [
            f"row {a['row']} ({', '.join(a['columns_used'])})"
            for a in section["sources"]["adae"]
        ]
        prot_cites = [
            f"§ {p['section_title']} p.{','.join(map(str, p['pages']))}"
            for p in section["sources"]["protocol"]
        ]
        print(f"   sources — ADAE: {'; '.join(adae_cites) or '(none)'}")
        print(f"   sources — Protocol: {'; '.join(prot_cites) or '(none)'}")

    out_path = os.path.join(here, "data", f"narrative_{subject_id}.json")
    with open(out_path, "w") as f:
        json.dump(narrative, f, indent=2, default=str)
    print(f"\nfull JSON written to {out_path}")
