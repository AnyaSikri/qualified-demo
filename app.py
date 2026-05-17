"""Streamlit UI for the PSN generator.

Layout:
  - Top row: subject ID, ADAE uploader, protocol uploader
  - Middle: collapsed template viewer + Generate button
  - Bottom: 60/40 split — narrative on the left, source panel on the right
  - Below: download buttons for narrative.docx and sources.json
"""

import html
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

# Pastel accents per template section — used for narrative left borders and
# matching source cards so prose and citation stay visually linked.
SECTION_ACCENTS = {
    "patient_info":      {"bar": "#A8C5E0", "bg": "#EFF6FB"},  # powder blue
    "event_description": {"bar": "#C8A8D8", "bg": "#F4ECF7"},  # dusty lavender
    "treatment_outcome": {"bar": "#9FBFA0", "bg": "#ECF4ED"},  # sage
    "conmeds":           {"bar": "#E8B5A0", "bg": "#FBF0EA"},  # soft peach
}
DEFAULT_ACCENT = {"bar": "#C9C0DA", "bg": "#F1EDF7"}

CUSTOM_CSS = """
<style>
/* Page header */
.psn-header {
    background: linear-gradient(135deg, #EDE7F6 0%, #FBF0EA 100%);
    padding: 1.4rem 1.6rem;
    border-radius: 14px;
    border: 1px solid #E0D9EE;
    margin-bottom: 1.2rem;
}
.psn-header h1 {
    margin: 0;
    color: #3D3D5C;
    font-size: 1.7rem;
    letter-spacing: -0.01em;
}
.psn-header p {
    margin: 0.35rem 0 0 0;
    color: #6B6B85;
    font-size: 0.95rem;
}

/* Input section labels */
.psn-input-label {
    font-weight: 600;
    color: #5C5478;
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 0.4rem;
}

/* Narrative section card */
.psn-section {
    background: var(--psn-bg);
    border-left: 4px solid var(--psn-bar);
    padding: 1rem 1.2rem;
    border-radius: 8px;
    margin-bottom: 1rem;
}
.psn-section h3 {
    margin: 0 0 0.5rem 0;
    color: #3D3D5C;
    font-size: 1.1rem;
}
.psn-section p {
    margin: 0;
    color: #2D3748;
    line-height: 1.6;
}

/* Source card */
.psn-source {
    background: #FFFFFF;
    border: 1px solid #E5DFEF;
    border-left: 4px solid var(--psn-bar);
    border-radius: 8px;
    padding: 0.85rem 1rem;
    margin-bottom: 0.75rem;
    font-size: 0.85rem;
}
.psn-source .psn-source-title {
    font-weight: 600;
    color: #3D3D5C;
    margin-bottom: 0.4rem;
    font-size: 0.95rem;
}
.psn-source ul {
    margin: 0;
    padding-left: 1.1rem;
    color: #54546B;
}
.psn-source li {
    margin-bottom: 0.15rem;
}
.psn-source .psn-source-empty {
    color: #9090A8;
    font-style: italic;
}

/* Section divider for results header */
.psn-results-bar {
    background: #F4EFFB;
    padding: 0.7rem 1rem;
    border-radius: 8px;
    border-left: 4px solid #9C89B8;
    color: #3D3D5C;
    font-weight: 600;
    margin: 1rem 0;
}

/* Tighten the default primary button */
div.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #9C89B8 0%, #B19CD9 100%);
    border: none;
    color: white;
    font-weight: 600;
    letter-spacing: 0.02em;
    border-radius: 8px;
    padding: 0.55rem 1rem;
    transition: transform 0.05s ease, box-shadow 0.15s ease;
}
div.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(156, 137, 184, 0.35);
}
</style>
"""


def _accent(section_id):
    return SECTION_ACCENTS.get(section_id, DEFAULT_ACCENT)


def _escape_html(s):
    """Escape user/LLM text before embedding into our custom HTML cards."""
    return html.escape(s, quote=False)


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
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    st.markdown(
        """
        <div class="psn-header">
            <h1>Patient Safety Narrative Generator</h1>
            <p>Draft a regulatory safety letter for a clinical trial subject —
            every sentence maps back to a source row in the ADAE Excel or a
            section in the protocol PDF.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # --- Inputs row ----------------------------------------------------------
    col_subj, col_adae, col_prot = st.columns([1, 1, 1])

    with col_subj:
        st.markdown('<div class="psn-input-label">Subject</div>', unsafe_allow_html=True)
        subject_id = st.text_input("Subject ID", value="SUBJ-1042", label_visibility="collapsed")

    with col_adae:
        st.markdown('<div class="psn-input-label">ADAE Excel</div>', unsafe_allow_html=True)
        use_default_adae = st.checkbox("Use default synthetic ADAE", value=True)
        adae_upload = st.file_uploader(
            "ADAE Excel (.xlsx)", type=["xlsx"],
            disabled=use_default_adae, label_visibility="collapsed",
        )

    with col_prot:
        st.markdown('<div class="psn-input-label">Protocol PDF</div>', unsafe_allow_html=True)
        use_default_prot = st.checkbox("Use default mock protocol", value=True)
        prot_upload = st.file_uploader(
            "Protocol PDF", type=["pdf"],
            disabled=use_default_prot, label_visibility="collapsed",
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

    st.markdown(
        f'<div class="psn-results-bar">PSN — Subject {narrative["subject_id"]} '
        f'· {narrative["template"]["name"]} v{narrative["template"]["version"]}</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns([6, 4])

    with left:
        for section in narrative["sections"]:
            accent = _accent(section["section_id"])
            text_html = _escape_html(section["text"]).replace("\n", "<br>")
            st.markdown(
                f'''
                <div class="psn-section" style="--psn-bar:{accent['bar']}; --psn-bg:{accent['bg']};">
                    <h3>{section["title"]}</h3>
                    <p>{text_html}</p>
                </div>
                ''',
                unsafe_allow_html=True,
            )

    with right:
        st.markdown(
            '<div class="psn-input-label" style="margin-top:0.2rem;">Sources</div>',
            unsafe_allow_html=True,
        )
        for section in narrative["sections"]:
            accent = _accent(section["section_id"])
            adae_cites = [_fmt_adae_cite(a) for a in section["sources"]["adae"]]
            prot_cites = [_fmt_protocol_cite(p) for p in section["sources"]["protocol"]]

            items = adae_cites + prot_cites
            if items:
                body = "<ul>" + "".join(
                    f"<li>{_escape_html(item)}</li>" for item in items
                ) + "</ul>"
            else:
                body = '<div class="psn-source-empty">(no sources)</div>'

            st.markdown(
                f'''
                <div class="psn-source" style="--psn-bar:{accent['bar']};">
                    <div class="psn-source-title">{section["title"]}</div>
                    {body}
                </div>
                ''',
                unsafe_allow_html=True,
            )

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
