"""Streamlit UI for the PSN generator.

Layout:
  - Top row: subject ID, ADAE uploader, protocol uploader
  - Middle: collapsed template viewer + Generate button
  - Bottom: 60/40 split — narrative on the left, source panel on the right
  - Below: download buttons for narrative.docx and sources.json
"""

import io
import json
import os
import tempfile

import streamlit as st

from docx_builder import build_narrative_docx, build_sources_json
from generator import generate_narrative
from template import NARRATIVE_TEMPLATE

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ADAE = os.path.join(HERE, "data", "ADAE.xlsx")
DEFAULT_PROTOCOL = os.path.join(HERE, "data", "mock_protocol.pdf")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resolve_input(uploaded_file, use_default, default_path, suffix):
    """Return a filesystem path for a Streamlit input — either the default
    synthetic file or a temp copy of the upload. Streamlit's UploadedFile
    can't be opened by pandas/PyMuPDF directly without writing it out."""
    if use_default:
        if not os.path.exists(default_path):
            st.error(
                f"Default file not found at {default_path}. "
                "Run `python build_synthetic_data.py` first."
            )
            st.stop()
        return default_path
    if uploaded_file is None:
        st.error("Upload a file or check 'Use default'.")
        st.stop()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.flush()
    tmp.close()
    return tmp.name


def _fmt_adae_cite(adae_source):
    cols = ", ".join(adae_source["columns_used"]) or "(no fields)"
    return f"ADAE row {adae_source['row']} ({cols})"


def _fmt_protocol_cite(protocol_source):
    pages = ",".join(str(p) for p in protocol_source["pages"]) or "?"
    return f"Protocol § {protocol_source['section_title']} p.{pages}"


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="PSN Generator", layout="wide")
    st.title("Patient Safety Narrative Generator")
    st.caption(
        "Draft regulatory safety letter for a clinical trial subject. "
        "Every generated sentence maps back to a source row in the ADAE Excel "
        "or a section in the protocol PDF."
    )

    # --- Inputs row ----------------------------------------------------------
    col_subj, col_adae, col_prot = st.columns([1, 1, 1])

    with col_subj:
        st.subheader("Subject")
        subject_id = st.text_input("Subject ID", value="SUBJ-1042")

    with col_adae:
        st.subheader("ADAE")
        use_default_adae = st.checkbox("Use default synthetic ADAE", value=True)
        adae_upload = st.file_uploader(
            "ADAE Excel (.xlsx)", type=["xlsx"], disabled=use_default_adae
        )

    with col_prot:
        st.subheader("Protocol")
        use_default_prot = st.checkbox("Use default mock protocol", value=True)
        prot_upload = st.file_uploader(
            "Protocol PDF", type=["pdf"], disabled=use_default_prot
        )

    # --- Template viewer + generate ----------------------------------------
    with st.expander(
        f"Template: {NARRATIVE_TEMPLATE['name']} v{NARRATIVE_TEMPLATE['version']}",
        expanded=False,
    ):
        for section in NARRATIVE_TEMPLATE["sections"]:
            st.markdown(f"**{section['title']}** &nbsp; `{section['id']}`")
            st.markdown(f"_{section['instructions']}_")
            st.markdown(
                f"ADAE fields: `{', '.join(section['adae_fields'])}` &nbsp;|&nbsp; "
                f"Protocol: `{', '.join(section['protocol_sections']) or '—'}`"
            )
            st.markdown("---")

    generate = st.button("Generate Narrative", type="primary", use_container_width=True)

    # --- Generation ---------------------------------------------------------
    if generate:
        adae_path = _resolve_input(adae_upload, use_default_adae, DEFAULT_ADAE, ".xlsx")
        prot_path = _resolve_input(prot_upload, use_default_prot, DEFAULT_PROTOCOL, ".pdf")

        status = st.status("Generating narrative...", expanded=True)
        log_lines = []

        def log(msg):
            log_lines.append(msg)
            status.write(msg)

        try:
            narrative = generate_narrative(
                subject_id.strip(), adae_path, prot_path, NARRATIVE_TEMPLATE, log=log
            )
        except Exception as e:
            status.update(label=f"Failed: {e}", state="error")
            st.exception(e)
            return

        status.update(label="Narrative generated.", state="complete")
        st.session_state["narrative"] = narrative

    # --- Render results -----------------------------------------------------
    narrative = st.session_state.get("narrative")
    if not narrative:
        return

    st.divider()
    st.subheader(
        f"PSN — Subject {narrative['subject_id']}  "
        f"·  {narrative['template']['name']} v{narrative['template']['version']}"
    )

    left, right = st.columns([6, 4])

    with left:
        for section in narrative["sections"]:
            st.markdown(f"### {section['title']}")
            st.write(section["text"])

    with right:
        st.markdown("#### Sources")
        for section in narrative["sections"]:
            with st.container(border=True):
                st.markdown(f"**{section['title']}**")
                adae_cites = [_fmt_adae_cite(a) for a in section["sources"]["adae"]]
                prot_cites = [_fmt_protocol_cite(p) for p in section["sources"]["protocol"]]
                if adae_cites:
                    st.markdown("• " + "  \n• ".join(adae_cites))
                if prot_cites:
                    st.markdown("• " + "  \n• ".join(prot_cites))
                if not adae_cites and not prot_cites:
                    st.markdown("_(no sources)_")

    # --- Downloads ----------------------------------------------------------
    st.divider()
    col_dl1, col_dl2 = st.columns(2)

    docx_buf = io.BytesIO()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        build_narrative_docx(narrative, tmp.name)
        with open(tmp.name, "rb") as f:
            docx_buf.write(f.read())
        os.unlink(tmp.name)
    docx_buf.seek(0)

    sources_payload = json.dumps(
        {
            "subject_id": narrative["subject_id"],
            "template": narrative["template"],
            "sections": narrative["sections"],
        },
        indent=2,
        default=str,
    )

    with col_dl1:
        st.download_button(
            "Download narrative.docx",
            data=docx_buf,
            file_name=f"narrative_{narrative['subject_id']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col_dl2:
        st.download_button(
            "Download sources.json",
            data=sources_payload,
            file_name=f"sources_{narrative['subject_id']}.json",
            mime="application/json",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
