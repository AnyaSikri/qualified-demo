"""Generates synthetic ADAE.xlsx and mock_protocol.pdf into ./data/.

Run directly:  python build_synthetic_data.py
"""

import os

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
ADAE_PATH = os.path.join(DATA_DIR, "ADAE.xlsx")
PROTOCOL_PATH = os.path.join(DATA_DIR, "mock_protocol.pdf")

STUDY_ID = "C906289002"


# ---------------------------------------------------------------------------
# ADAE
# ---------------------------------------------------------------------------

ADAE_ROWS = [
    # SUBJ-1042 — Grade 3 pneumonia, hospitalized, drug interrupted, resolved
    {
        "USUBJID": "SUBJ-1042", "AGE": 64, "SEX": "M", "RACE": "WHITE",
        "AEDECOD": "Pneumonia bacterial", "AESEV": "SEVERE",
        "AESTDTC": "2025-03-14", "AEENDTC": "2025-03-29",
        "AESER": "Y", "AETOXGR": 3,
        "AEACN": "DRUG INTERRUPTED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Ceftriaxone", "CMDOSE": "2 g IV daily", "CMSTDTC": "2025-03-15",
    },
    {
        "USUBJID": "SUBJ-1042", "AGE": 64, "SEX": "M", "RACE": "WHITE",
        "AEDECOD": "Pneumonia bacterial", "AESEV": "SEVERE",
        "AESTDTC": "2025-03-14", "AEENDTC": "2025-03-29",
        "AESER": "Y", "AETOXGR": 3,
        "AEACN": "DRUG INTERRUPTED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Acetaminophen", "CMDOSE": "650 mg PO q6h PRN", "CMSTDTC": "2025-03-14",
    },

    # SUBJ-0817 — Grade 5 sepsis, fatal, missing some conmed fields
    {
        "USUBJID": "SUBJ-0817", "AGE": 71, "SEX": "F", "RACE": "BLACK OR AFRICAN AMERICAN",
        "AEDECOD": "Sepsis", "AESEV": "SEVERE",
        "AESTDTC": "2025-02-02", "AEENDTC": "2025-02-09",
        "AESER": "Y", "AETOXGR": 5,
        "AEACN": "DRUG WITHDRAWN", "AEOUT": "FATAL",
        "CMTRT": "Piperacillin-tazobactam", "CMDOSE": "", "CMSTDTC": "",
    },
    {
        "USUBJID": "SUBJ-0817", "AGE": 71, "SEX": "F", "RACE": "BLACK OR AFRICAN AMERICAN",
        "AEDECOD": "Sepsis", "AESEV": "SEVERE",
        "AESTDTC": "2025-02-02", "AEENDTC": "2025-02-09",
        "AESER": "Y", "AETOXGR": 5,
        "AEACN": "DRUG WITHDRAWN", "AEOUT": "FATAL",
        "CMTRT": "Norepinephrine", "CMDOSE": "", "CMSTDTC": "2025-02-03",
    },

    # SUBJ-2156 — Grade 4 neutropenia, dose reduced
    {
        "USUBJID": "SUBJ-2156", "AGE": 58, "SEX": "M", "RACE": "ASIAN",
        "AEDECOD": "Neutropenia", "AESEV": "SEVERE",
        "AESTDTC": "2025-04-10", "AEENDTC": "2025-04-22",
        "AESER": "Y", "AETOXGR": 4,
        "AEACN": "DOSE REDUCED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Filgrastim", "CMDOSE": "300 mcg SC daily", "CMSTDTC": "2025-04-11",
    },

    # Filler 1 — SUBJ-3308: Grade 2 nausea, no action
    {
        "USUBJID": "SUBJ-3308", "AGE": 49, "SEX": "F", "RACE": "WHITE",
        "AEDECOD": "Nausea", "AESEV": "MODERATE",
        "AESTDTC": "2025-01-22", "AEENDTC": "2025-01-25",
        "AESER": "N", "AETOXGR": 2,
        "AEACN": "DOSE NOT CHANGED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Ondansetron", "CMDOSE": "8 mg PO q8h PRN", "CMSTDTC": "2025-01-22",
    },

    # Filler 2 — SUBJ-4471: Grade 3 diarrhea, drug interrupted
    {
        "USUBJID": "SUBJ-4471", "AGE": 67, "SEX": "M", "RACE": "WHITE",
        "AEDECOD": "Diarrhoea", "AESEV": "SEVERE",
        "AESTDTC": "2025-03-02", "AEENDTC": "2025-03-11",
        "AESER": "N", "AETOXGR": 3,
        "AEACN": "DRUG INTERRUPTED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Loperamide", "CMDOSE": "4 mg PO PRN", "CMSTDTC": "2025-03-02",
    },

    # Filler 3 — SUBJ-5589: Grade 1 fatigue, ongoing
    {
        "USUBJID": "SUBJ-5589", "AGE": 55, "SEX": "F", "RACE": "ASIAN",
        "AEDECOD": "Fatigue", "AESEV": "MILD",
        "AESTDTC": "2025-04-01", "AEENDTC": "",
        "AESER": "N", "AETOXGR": 1,
        "AEACN": "DOSE NOT CHANGED", "AEOUT": "NOT RECOVERED/NOT RESOLVED",
        "CMTRT": "", "CMDOSE": "", "CMSTDTC": "",
    },

    # Filler 4 — SUBJ-6612: Grade 3 ALT increased, dose reduced
    {
        "USUBJID": "SUBJ-6612", "AGE": 62, "SEX": "M", "RACE": "WHITE",
        "AEDECOD": "Alanine aminotransferase increased", "AESEV": "SEVERE",
        "AESTDTC": "2025-02-18", "AEENDTC": "2025-03-05",
        "AESER": "N", "AETOXGR": 3,
        "AEACN": "DOSE REDUCED", "AEOUT": "RECOVERED/RESOLVED",
        "CMTRT": "Ursodiol", "CMDOSE": "300 mg PO BID", "CMSTDTC": "2025-02-19",
    },

    # Filler 5 — SUBJ-7724: Grade 4 thrombocytopenia, hospitalized, recovered with sequelae
    {
        "USUBJID": "SUBJ-7724", "AGE": 70, "SEX": "F", "RACE": "WHITE",
        "AEDECOD": "Thrombocytopenia", "AESEV": "SEVERE",
        "AESTDTC": "2025-03-27", "AEENDTC": "2025-04-09",
        "AESER": "Y", "AETOXGR": 4,
        "AEACN": "DRUG INTERRUPTED", "AEOUT": "RECOVERED/RESOLVED WITH SEQUELAE",
        "CMTRT": "Platelet transfusion", "CMDOSE": "1 unit IV", "CMSTDTC": "2025-03-28",
    },
]

ADAE_COLUMNS = [
    "USUBJID", "AGE", "SEX", "RACE",
    "AEDECOD", "AESEV", "AESTDTC", "AEENDTC", "AESER", "AETOXGR",
    "AEACN", "AEOUT",
    "CMTRT", "CMDOSE", "CMSTDTC",
]


def build_adae(out_path):
    df = pd.DataFrame(ADAE_ROWS, columns=ADAE_COLUMNS)
    df.to_excel(out_path, index=False, sheet_name="ADAE")
    return df


# ---------------------------------------------------------------------------
# Protocol PDF
# ---------------------------------------------------------------------------

PROTOCOL_SECTIONS = [
    ("1. Background", [
        f"Study {STUDY_ID} is a Mock Phase 2 oncology trial evaluating investigational "
        "agent MX-209 in adults with relapsed or refractory solid tumors. Prior studies "
        "demonstrated single-agent activity and an acceptable safety profile at the "
        "recommended Phase 2 dose. This protocol is provided for demonstration purposes "
        "only and does not describe a real clinical investigation.",
        "The investigational product targets a tumor-associated signaling pathway that has "
        "been implicated in proliferation and resistance to standard cytotoxic therapy. "
        "Preclinical models indicated tumor growth inhibition across multiple histologies.",
    ]),
    ("2. Objectives", [
        "The primary objective of Study {study} is to estimate the objective response rate "
        "per RECIST v1.1 in the eligible population. Key secondary objectives include "
        "progression-free survival, overall survival, duration of response, and "
        "characterization of the safety and tolerability profile.".format(study=STUDY_ID),
        "Exploratory objectives include pharmacokinetic characterization and assessment of "
        "candidate biomarkers of response and resistance.",
    ]),
    ("3. Study Design", [
        "This is an open-label, single-arm, multicenter Phase 2 study. Approximately 80 "
        "subjects will be enrolled across 12 sites in North America and Europe. Subjects "
        "will receive MX-209 in 28-day cycles until disease progression, unacceptable "
        "toxicity, or withdrawal of consent.",
        "Tumor assessments will be performed at screening and every two cycles (8 weeks). "
        "End-of-treatment and 30-day safety follow-up visits are required for all subjects.",
    ]),
    ("4. Eligibility", [
        "Subjects must be at least 18 years of age with histologically or cytologically "
        "confirmed solid tumor that has progressed on or after standard therapy. ECOG "
        "performance status of 0 or 1 is required. Adequate hematologic, hepatic, and "
        "renal function as defined by the laboratory criteria in Appendix A must be "
        "documented within 14 days prior to enrollment.",
        "Key exclusion criteria include active central nervous system metastases, "
        "uncontrolled intercurrent illness, pregnancy or lactation, and receipt of any "
        "investigational therapy within 28 days prior to the first dose of study drug.",
    ]),
    ("5. Dosing", [
        "MX-209 will be administered orally at the recommended Phase 2 dose of 200 mg "
        "once daily on a continuous schedule. Doses should be taken at approximately the "
        "same time each day with or without food. Missed doses may be taken within 6 "
        "hours of the scheduled time; otherwise the dose should be skipped.",
        "Dose modifications for treatment-emergent toxicity follow the schema in Appendix "
        "B. Two dose reductions are permitted (to 150 mg and 100 mg). Subjects requiring "
        "a third dose reduction will be discontinued from study treatment.",
    ]),
    ("6. Adverse Event Definitions", [
        "An adverse event (AE) is any untoward medical occurrence in a subject "
        "administered the investigational product, whether or not considered related to "
        "the product. AEs will be graded according to NCI CTCAE v5.0 and coded using the "
        "latest version of MedDRA.",
        "A serious adverse event (SAE) is any AE that results in death, is life "
        "threatening, requires inpatient hospitalization or prolongation of existing "
        "hospitalization, results in persistent or significant disability or incapacity, "
        "is a congenital anomaly or birth defect, or is an important medical event based "
        "on medical judgment.",
        "Adverse events of special interest (AESIs) for this protocol include febrile "
        "neutropenia, Grade 3 or higher hepatic enzyme elevations, and any infection of "
        "Grade 3 or higher severity.",
    ]),
    ("7. Safety Monitoring", [
        "All subjects will be monitored continuously for AEs from the time of first dose "
        "through 30 days after the last dose. Investigators will assess and grade events "
        "at each scheduled visit and as clinically indicated.",
        "Serious adverse events must be reported to the Sponsor within 24 hours of "
        "investigator awareness, regardless of causality. The Sponsor will forward "
        "qualifying events to applicable regulatory authorities in accordance with local "
        "expedited reporting requirements.",
        "An independent Data Monitoring Committee will review aggregate safety data at "
        "pre-specified intervals and may recommend protocol modifications or study "
        "suspension based on emerging safety signals.",
    ]),
    ("8. Statistical Analysis", [
        "The primary analysis population for efficacy is the Full Analysis Set, defined "
        "as all enrolled subjects who receive at least one dose of MX-209. The Safety "
        "Analysis Set is identical to the Full Analysis Set for this protocol.",
        "Objective response rate will be summarized with a two-sided 95% Clopper-Pearson "
        "confidence interval. Time-to-event endpoints will be summarized using "
        "Kaplan-Meier methodology. No formal hypothesis testing is planned.",
    ]),
]


def build_protocol(out_path):
    doc = SimpleDocTemplate(
        out_path,
        pagesize=letter,
        leftMargin=0.9 * inch,
        rightMargin=0.9 * inch,
        topMargin=0.9 * inch,
        bottomMargin=0.9 * inch,
        title=f"Study {STUDY_ID} Mock Protocol",
    )
    styles = getSampleStyleSheet()
    heading = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading1"],
        fontSize=14,
        spaceBefore=12,
        spaceAfter=10,
    )
    body = ParagraphStyle(
        "BodyJustified",
        parent=styles["BodyText"],
        fontSize=10.5,
        leading=14,
        spaceAfter=8,
        alignment=4,
    )
    cover_title = ParagraphStyle(
        "CoverTitle",
        parent=styles["Title"],
        fontSize=20,
        spaceAfter=16,
    )

    flow = []
    flow.append(Paragraph(f"Study {STUDY_ID}", cover_title))
    flow.append(Paragraph("Mock Phase 2 Oncology Trial", styles["Heading2"]))
    flow.append(Spacer(1, 0.4 * inch))
    flow.append(Paragraph(
        "Protocol Version 1.0 &middot; Synthetic document for demonstration only.",
        body,
    ))
    flow.append(PageBreak())

    for idx, (title, paragraphs) in enumerate(PROTOCOL_SECTIONS):
        flow.append(Paragraph(title, heading))
        for para in paragraphs:
            flow.append(Paragraph(para, body))
        # Push later sections to fresh pages so the doc spans ~8 pages and
        # parser regex finds headers near page tops.
        if idx < len(PROTOCOL_SECTIONS) - 1:
            flow.append(PageBreak())

    doc.build(flow)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    os.makedirs(DATA_DIR, exist_ok=True)

    print(f"[1/2] Writing ADAE → {ADAE_PATH}")
    df = build_adae(ADAE_PATH)
    print(f"      {len(df)} rows, {df['USUBJID'].nunique()} subjects")

    print(f"[2/2] Writing protocol PDF → {PROTOCOL_PATH}")
    build_protocol(PROTOCOL_PATH)
    size_kb = os.path.getsize(PROTOCOL_PATH) / 1024
    print(f"      {size_kb:.1f} KB")

    print("Done.")
