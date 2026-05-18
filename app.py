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
import textwrap

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

/* Narrative — styled as a single "letter on paper", not a card list */
.psn-document {
    background: #FFFFFF;
    border-radius: 14px;
    border: 1px solid #E5DFEF;
    box-shadow: 0 4px 16px rgba(45, 55, 72, 0.10);
    padding: 2rem 2.2rem 1.5rem 2.2rem;
}
.psn-doc-header {
    border-bottom: 2px solid #EDE7F6;
    padding-bottom: 1rem;
    margin-bottom: 1.4rem;
}
.psn-doc-eyebrow {
    color: #9C89B8;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin-bottom: 0.35rem;
}
.psn-doc-title {
    color: #2D3748;
    font-family: Georgia, "Times New Roman", serif;
    font-size: 1.55rem;
    font-weight: 600;
    margin: 0;
    line-height: 1.25;
}
.psn-doc-meta {
    color: #6B6B85;
    font-size: 0.85rem;
    margin-top: 0.45rem;
}
.psn-doc-meta strong {
    color: #3D3D5C;
}
.psn-doc-section {
    padding: 0.9rem 0 1.1rem 0;
    border-bottom: 1px dashed #EDE7F6;
}
.psn-doc-section:last-of-type {
    border-bottom: none;
    padding-bottom: 0.2rem;
}
.psn-doc-section-head {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    margin-bottom: 0.55rem;
}
.psn-doc-badge {
    flex: none;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: var(--psn-bar);
    color: #FFFFFF;
    font-weight: 700;
    font-size: 0.85rem;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 6px rgba(0,0,0,0.12);
}
.psn-doc-section-title {
    font-family: Georgia, "Times New Roman", serif;
    color: #2D3748;
    font-size: 1.15rem;
    font-weight: 600;
    margin: 0;
}
.psn-doc-section-body {
    font-family: Georgia, "Times New Roman", serif;
    color: #2D3748;
    font-size: 1.0rem;
    line-height: 1.7;
    margin: 0;
    padding-left: 0.1rem;
}
.psn-doc-footer {
    margin-top: 1.4rem;
    padding-top: 0.9rem;
    border-top: 2px solid #EDE7F6;
    color: #9090A8;
    font-size: 0.72rem;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    text-align: center;
}

/* Sources panel — the entire right column, visually distinct from narrative */
.psn-sources-panel {
    background: #2D3748;
    border-radius: 14px;
    padding: 1.2rem 1.1rem 0.6rem 1.1rem;
    border: 1px solid #1F2937;
    box-shadow: 0 4px 16px rgba(45, 55, 72, 0.18);
}
.psn-sources-panel-header {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    margin-bottom: 0.9rem;
    padding-bottom: 0.7rem;
    border-bottom: 1px solid #4A5568;
}
.psn-sources-panel-header .title {
    color: #F7FAFC;
    font-size: 0.95rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.psn-sources-panel-header .subtitle {
    color: #A0AEC0;
    font-size: 0.72rem;
    letter-spacing: 0.04em;
}

/* Source card — dark, compact, monospaced citations */
.psn-source {
    background: #FFFFFF;
    border-radius: 8px;
    border-left: 4px solid var(--psn-bar);
    padding: 0.75rem 0.85rem;
    margin-bottom: 0.65rem;
    font-size: 0.82rem;
}
.psn-source .psn-source-title {
    font-weight: 600;
    color: #3D3D5C;
    margin-bottom: 0.55rem;
    font-size: 0.88rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.psn-source .psn-cite {
    display: flex;
    align-items: flex-start;
    gap: 0.45rem;
    margin: 0.3rem 0;
    font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
    font-size: 0.76rem;
    color: #2D3748;
    line-height: 1.4;
}
.psn-source .psn-chip {
    flex: none;
    display: inline-block;
    padding: 0.08rem 0.42rem;
    border-radius: 4px;
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    font-family: ui-sans-serif, system-ui, sans-serif;
    line-height: 1.4;
}
.psn-source .psn-chip-adae {
    background: #DCEBF7;
    color: #2C5282;
}
.psn-source .psn-chip-protocol {
    background: #E0EDDB;
    color: #2F5233;
}
.psn-source .psn-source-empty {
    color: #9090A8;
    font-style: italic;
    font-size: 0.78rem;
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


def _md_html(markup):
    """Render raw HTML through st.markdown.

    Streamlit's markdown parser treats lines indented 4+ spaces as code
    blocks, so we MUST dedent before passing through. The strip removes
    leading/trailing newlines that would otherwise also confuse the parser.
    """
    st.markdown(textwrap.dedent(markup).strip(), unsafe_allow_html=True)


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

    _md_html("""
        <div class="psn-header">
            <h1>Patient Safety Narrative Generator</h1>
            <p>Draft a regulatory safety letter for a clinical trial subject —
            every sentence maps back to a source row in the ADAE Excel or a
            section in the protocol PDF.</p>
        </div>
    """)

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

    st.write("")  # vertical breathing room before the two-column layout
    left, right = st.columns([6, 4])

    with left:
        section_blocks = []
        for idx, section in enumerate(narrative["sections"], start=1):
            accent = _accent(section["section_id"])
            text_html = _escape_html(section["text"]).replace("\n", "<br>")
            # Built as a single line: any blank line inside the joined output
            # would terminate the outer document's HTML block in markdown,
            # leaking the literal markup of subsequent sections.
            section_blocks.append(
                f'<div class="psn-doc-section">'
                f'<div class="psn-doc-section-head">'
                f'<div class="psn-doc-badge" style="--psn-bar:{accent["bar"]};">{idx}</div>'
                f'<div class="psn-doc-section-title">{_escape_html(section["title"])}</div>'
                f'</div>'
                f'<p class="psn-doc-section-body">{text_html}</p>'
                f'</div>'
            )

        _md_html(f"""
            <div class="psn-document">
                <div class="psn-doc-header">
                    <div class="psn-doc-eyebrow">Patient Safety Narrative &middot; Draft</div>
                    <h2 class="psn-doc-title">Subject {_escape_html(narrative["subject_id"])}</h2>
                    <div class="psn-doc-meta">
                        Template: <strong>{_escape_html(narrative["template"]["name"])}</strong>
                        v{_escape_html(narrative["template"]["version"])}
                    </div>
                </div>
                {"".join(section_blocks)}
                <div class="psn-doc-footer">
                    Draft &mdash; for investigator review, not for distribution
                </div>
            </div>
        """)

    with right:
        cards_html = []
        for section in narrative["sections"]:
            accent = _accent(section["section_id"])
            adae_cites = [_fmt_adae_cite(a) for a in section["sources"]["adae"]]
            prot_cites = [_fmt_protocol_cite(p) for p in section["sources"]["protocol"]]

            cite_lines = []
            for cite in adae_cites:
                cite_lines.append(
                    f'<div class="psn-cite">'
                    f'<span class="psn-chip psn-chip-adae">ADAE</span>'
                    f'<span>{_escape_html(cite.removeprefix("ADAE "))}</span>'
                    f'</div>'
                )
            for cite in prot_cites:
                cite_lines.append(
                    f'<div class="psn-cite">'
                    f'<span class="psn-chip psn-chip-protocol">PROTOCOL</span>'
                    f'<span>{_escape_html(cite.removeprefix("Protocol "))}</span>'
                    f'</div>'
                )
            if not cite_lines:
                cite_lines.append('<div class="psn-source-empty">no citations</div>')

            cards_html.append(
                f'<div class="psn-source" style="--psn-bar:{accent["bar"]};">'
                f'<div class="psn-source-title">{_escape_html(section["title"])}</div>'
                f'{"".join(cite_lines)}'
                f'</div>'
            )

        n_sections = len(narrative["sections"])
        total_adae = sum(len(s["sources"]["adae"]) for s in narrative["sections"])
        total_prot = sum(len(s["sources"]["protocol"]) for s in narrative["sections"])

        _md_html(f"""
            <div class="psn-sources-panel">
                <div class="psn-sources-panel-header">
                    <span class="title">Source Trail</span>
                    <span class="subtitle">{n_sections} sections &middot;
                        {total_adae} ADAE &middot; {total_prot} protocol</span>
                </div>
                {"".join(cards_html)}
            </div>
        """)

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
