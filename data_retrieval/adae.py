"""ADAE Excel retrieval.

Loads an SDTM-style ADAE workbook, filters to one subject, and projects the
fields requested by a template section. Every projection carries back the
source row index and the columns actually used so the orchestrator can build
traceable citations.
"""

import math

import pandas as pd


def load_adae(path):
    """Return the ADAE sheet as a DataFrame with original row indices preserved."""
    df = pd.read_excel(path, sheet_name=0)
    df.index = range(len(df))  # explicit, contiguous; row index == citation row
    return df


def get_subject_events(df, subject_id):
    """Return all rows for a subject, preserving original row indices."""
    return df[df["USUBJID"] == subject_id].copy()


def _is_missing(value):
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _clean(value):
    """Normalize NaN / empty strings to None so downstream prompt builders can
    flag gaps explicitly instead of leaking 'nan' into LLM context."""
    return None if _is_missing(value) else value


def get_fields_for_section(events, section_config):
    """Project section-relevant fields out of the subject's rows.

    Returns a list of per-row dicts:
        {"row": <int>, "columns_used": [...], "values": {col: val, ...},
         "missing": [cols that were blank]}
    The orchestrator passes `values` to the LLM and `row` / `columns_used` to
    the traceability layer.
    """
    fields = [f for f in section_config["adae_fields"] if f in events.columns]
    records = []
    for row_idx, row in events.iterrows():
        values, missing = {}, []
        for field in fields:
            cleaned = _clean(row[field])
            if cleaned is None:
                missing.append(field)
            else:
                values[field] = cleaned
        records.append({
            "row": int(row_idx),
            "columns_used": fields,
            "values": values,
            "missing": missing,
        })
    return records


# ---------------------------------------------------------------------------
# Standalone smoke test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import json
    import os
    import sys

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from template import NARRATIVE_TEMPLATE  # noqa: E402

    here = os.path.dirname(os.path.abspath(__file__))
    adae_path = os.path.join(here, "..", "data", "ADAE.xlsx")

    print(f"loading {adae_path}")
    df = load_adae(adae_path)
    print(f"  {len(df)} rows, columns: {list(df.columns)}")

    for subject_id in ["SUBJ-1042", "SUBJ-0817"]:
        print(f"\n--- {subject_id} ---")
        events = get_subject_events(df, subject_id)
        print(f"  {len(events)} event row(s)")
        for section in NARRATIVE_TEMPLATE["sections"]:
            picked = get_fields_for_section(events, section)
            print(f"  section={section['id']}")
            print("    " + json.dumps(picked, indent=2, default=str).replace("\n", "\n    "))
