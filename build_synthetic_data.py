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
        "agent MX-209 in adults with relapsed or refractory solid tumors who have "
        "exhausted standard-of-care therapy. This protocol is provided for demonstration "
        "purposes only and does not describe a real clinical investigation; all subjects, "
        "sites, sponsors, and outcomes referenced herein are fictional.",
        "Despite advances in targeted therapy and immune checkpoint inhibition over the "
        "past decade, a meaningful fraction of patients with advanced solid tumors develop "
        "resistance to available treatments and have limited subsequent options. Median "
        "overall survival following progression on second-line therapy remains under "
        "twelve months across most histologies, and tolerability of cytotoxic salvage "
        "regimens is often poor in heavily pretreated populations.",
        "MX-209 is an orally bioavailable small-molecule inhibitor of a tumor-associated "
        "kinase implicated in proliferation, DNA damage repair, and acquired resistance to "
        "standard cytotoxic therapy. The compound demonstrated single-agent activity "
        "across a panel of patient-derived xenograft models spanning non-small cell lung, "
        "colorectal, ovarian, and triple-negative breast histologies, with tumor growth "
        "inhibition observed at exposures achievable in human dosing.",
        "The first-in-human dose-escalation study (Study C906289001) enrolled 42 subjects "
        "across six dose levels. The maximum tolerated dose was not formally reached; the "
        "recommended Phase 2 dose (RP2D) of 200 mg once daily was selected on the basis "
        "of pharmacokinetic exposure, target engagement biomarkers, and the absence of "
        "dose-limiting toxicities at that level. Common treatment-emergent adverse events "
        "in the Phase 1 study included fatigue, nausea, transaminase elevation, and "
        "neutropenia.",
        "The present study is designed to estimate single-agent efficacy at the RP2D and "
        "to further characterize the safety profile in an enriched population. Findings "
        "will inform the design of subsequent randomized studies and combination strategies.",
    ]),
    ("2. Objectives", [
        f"The primary objective of Study {STUDY_ID} is to estimate the objective response "
        "rate (ORR) of MX-209 monotherapy, defined as the proportion of subjects achieving "
        "a confirmed complete or partial response per RECIST v1.1 as assessed by the "
        "investigator. The ORR will be summarized with a two-sided 95% confidence interval.",
        "Secondary efficacy objectives include estimation of progression-free survival "
        "(PFS), overall survival (OS), duration of response (DoR), and disease control "
        "rate (DCR) defined as the proportion of subjects achieving complete response, "
        "partial response, or stable disease lasting at least 16 weeks. All efficacy "
        "endpoints will be analyzed in the Full Analysis Set.",
        "Secondary safety objectives include characterization of the type, incidence, "
        "severity, timing, seriousness, causality, and resolution of treatment-emergent "
        "adverse events, with particular attention to adverse events of special interest "
        "(AESIs) as defined in Section 6. Dose modifications, treatment interruptions, and "
        "discontinuations due to toxicity will be tabulated.",
        "Exploratory objectives include characterization of the pharmacokinetic profile "
        "at the RP2D under steady-state conditions, assessment of candidate predictive "
        "biomarkers including baseline circulating tumor DNA mutational status and "
        "on-treatment dynamics, and exploration of resistance mechanisms through paired "
        "tumor biopsies collected at baseline and at the time of progression where "
        "clinically feasible.",
    ]),
    ("3. Study Design", [
        "This is a Phase 2, open-label, single-arm, multicenter study evaluating MX-209 "
        "monotherapy in adults with relapsed or refractory advanced solid tumors. "
        "Approximately 80 subjects will be enrolled across 12 investigative sites in "
        "North America and Europe. The study consists of three periods: a screening "
        "period of up to 28 days, an open-label treatment period of repeating 28-day "
        "cycles, and a follow-up period extending to 30 days after the last dose for "
        "safety and quarterly thereafter for survival.",
        "Subjects will receive MX-209 at the RP2D of 200 mg orally once daily on a "
        "continuous schedule. Treatment will continue until radiographic progression, "
        "unacceptable toxicity, withdrawal of consent, investigator decision, or "
        "completion of two years of therapy, whichever occurs first. Subjects who derive "
        "clinical benefit at the two-year mark may continue treatment with sponsor agreement.",
        "Disease assessments by computed tomography or magnetic resonance imaging will be "
        "performed at screening and every eight weeks (every two cycles) for the first "
        "twelve months, and every twelve weeks thereafter, until documented progression. "
        "Responses must be confirmed by a repeat assessment at least four weeks after the "
        "initial observation.",
        "Safety assessments including vital signs, ECOG performance status, physical "
        "examination, hematology, serum chemistry, coagulation, and 12-lead ECG will be "
        "performed at the start of each cycle and as clinically indicated. Concomitant "
        "medications will be recorded at every visit. Quality of life will be assessed "
        "at screening, every other cycle, and at end of treatment using the EORTC QLQ-C30 "
        "instrument.",
        "Translational research samples — peripheral blood for circulating tumor DNA and "
        "pharmacokinetics, and optional tumor biopsies — will be collected per the "
        "schedule in Appendix C. Pharmacokinetic sampling will be performed on Cycle 1 "
        "Day 1, Cycle 1 Day 15, and Cycle 2 Day 1.",
    ]),
    ("4. Eligibility", [
        "To be eligible, subjects must be at least 18 years of age at the time of consent, "
        "have a histologically or cytologically confirmed advanced solid tumor of any "
        "histology, have radiographically measurable disease per RECIST v1.1, and have "
        "documented progression on or intolerance to at least one prior line of "
        "systemic therapy for advanced or metastatic disease.",
        "Subjects must have an ECOG performance status of 0 or 1, a life expectancy of at "
        "least 12 weeks in the investigator's judgment, and provide written informed consent "
        "in accordance with local regulatory and institutional review board requirements. "
        "Women of childbearing potential must have a negative serum pregnancy test within "
        "seven days of the first dose and agree to use highly effective contraception for "
        "the duration of treatment and for 90 days following the last dose.",
        "Adequate organ function must be documented within 14 days prior to first dose: "
        "absolute neutrophil count at least 1.5 x 10^9/L, platelet count at least "
        "100 x 10^9/L, hemoglobin at least 9.0 g/dL without transfusion in the prior two "
        "weeks, total bilirubin at most 1.5 x upper limit of normal (3 x ULN in subjects "
        "with documented Gilbert's syndrome), AST and ALT at most 2.5 x ULN (5 x ULN in "
        "subjects with documented hepatic involvement of malignancy), and creatinine "
        "clearance at least 50 mL/min by Cockcroft-Gault estimation.",
        "Key exclusion criteria include known active central nervous system metastases "
        "(treated and stable CNS disease without ongoing corticosteroid requirement is "
        "permitted following at least 28 days of clinical stability), known leptomeningeal "
        "involvement, uncontrolled intercurrent illness including but not limited to "
        "symptomatic congestive heart failure, unstable angina, myocardial infarction "
        "within six months, clinically significant arrhythmia, or active uncontrolled "
        "infection requiring systemic therapy.",
        "Additional exclusions include receipt of any investigational agent within 28 "
        "days or five half-lives (whichever is shorter) prior to the first dose, major "
        "surgery within 28 days, palliative radiotherapy to a non-target lesion within 14 "
        "days, prior allogeneic stem cell or solid organ transplant, history of another "
        "malignancy within three years (with the exception of adequately treated "
        "non-melanoma skin cancer or carcinoma in situ), pregnancy or breastfeeding, and "
        "any condition that in the investigator's judgment would interfere with study "
        "participation or compromise the interpretation of study data.",
    ]),
    ("5. Dosing", [
        "MX-209 will be supplied as 50 mg and 100 mg film-coated tablets. The recommended "
        "Phase 2 dose is 200 mg administered orally once daily on a continuous schedule. "
        "Doses should be taken at approximately the same time each day, with or without "
        "food, and swallowed whole with water. Tablets must not be crushed, chewed, or "
        "dissolved.",
        "A missed dose may be taken within six hours of the scheduled administration time; "
        "if more than six hours have elapsed, the dose should be skipped and the next dose "
        "taken at the scheduled time the following day. Subjects must not take a double "
        "dose to compensate for a missed dose. Vomiting within 30 minutes of administration "
        "may be followed by a single re-dose; vomiting beyond that window should not be "
        "re-dosed.",
        "Two dose reductions are permitted in response to treatment-emergent toxicity, to "
        "150 mg and then to 100 mg, in accordance with the dose modification schema in "
        "Appendix B. Subjects requiring a third dose reduction will be discontinued from "
        "study treatment. Dose re-escalation is not permitted once a reduction has occurred.",
        "Dose interruptions for non-hematologic toxicity should generally not exceed 21 "
        "consecutive days; interruptions of longer duration will result in treatment "
        "discontinuation unless explicit sponsor concurrence is documented. Hematologic "
        "toxicity may permit interruptions of up to 28 days at investigator discretion.",
        "Drug accountability records must be maintained at each site. Subjects are "
        "instructed to return all unused tablets and empty containers at each on-site "
        "visit. Compliance will be assessed by tablet count and will be documented in the "
        "source record. Subjects with compliance below 75% over two consecutive cycles "
        "will be reviewed by the medical monitor.",
        "Concomitant strong inhibitors and inducers of CYP3A4 are prohibited from "
        "screening through 30 days after the last dose. Moderate inhibitors and inducers "
        "are permitted only when clinically necessary and require dose monitoring per "
        "Appendix D.",
    ]),
    ("6. Adverse Event Definitions", [
        "An adverse event (AE) is any untoward medical occurrence in a subject administered "
        "the investigational product, whether or not considered related to the product. "
        "This definition includes new signs, symptoms, or diseases, as well as worsening "
        "of pre-existing conditions. AEs will be graded according to NCI CTCAE v5.0 and "
        "coded using the latest version of MedDRA available at the time of database lock.",
        "A serious adverse event (SAE) is any AE that meets one or more of the following "
        "criteria: results in death; is life threatening (i.e., places the subject at "
        "immediate risk of death from the event as it occurred); requires inpatient "
        "hospitalization or prolongation of existing hospitalization; results in "
        "persistent or significant disability or incapacity; is a congenital anomaly or "
        "birth defect in the offspring of a study subject; or is an important medical "
        "event that may jeopardize the subject or may require medical or surgical "
        "intervention to prevent one of the above outcomes, based on medical judgment.",
        "Causality assessment will be performed by the investigator for every AE using "
        "a five-category scale: not related, unlikely related, possibly related, probably "
        "related, and definitely related. The Sponsor's medical monitor will perform an "
        "independent causality assessment for all SAEs. In the event of discordance, "
        "both assessments will be retained in the safety database and the more "
        "conservative assessment will be used for expedited regulatory reporting.",
        "Severity grading per CTCAE v5.0: Grade 1 (mild; asymptomatic or mild symptoms, "
        "clinical or diagnostic observations only, intervention not indicated), Grade 2 "
        "(moderate; minimal, local, or non-invasive intervention indicated), Grade 3 "
        "(severe or medically significant but not immediately life threatening; "
        "hospitalization or prolongation indicated), Grade 4 (life threatening; urgent "
        "intervention indicated), and Grade 5 (death related to AE).",
        "Adverse events of special interest (AESIs) for this protocol, requiring expedited "
        "internal review even when not meeting SAE criteria, include: febrile neutropenia "
        "of any grade; Grade 3 or higher hepatic enzyme elevation (AST, ALT, total "
        "bilirubin) including any case meeting Hy's law criteria; any infection of Grade "
        "3 or higher severity; Grade 3 or higher cardiac arrhythmia; QTc prolongation to "
        "greater than 500 ms or by more than 60 ms from baseline; interstitial lung "
        "disease or pneumonitis of any grade; and any Grade 4 or 5 hematologic toxicity.",
        "Pregnancies occurring in a study subject or in the partner of a male subject "
        "during treatment or within 90 days after the last dose must be reported to the "
        "Sponsor within 24 hours of investigator awareness. Pregnancy outcomes will be "
        "followed regardless of whether the pregnancy was discovered during or after "
        "treatment.",
    ]),
    ("7. Safety Monitoring", [
        "All subjects will be monitored continuously for adverse events from the time of "
        "first dose through 30 days after the last dose of study drug. Investigators will "
        "assess and grade events at each scheduled visit and as clinically indicated, "
        "and will document onset, severity, seriousness, causality, action taken with "
        "study drug, concomitant interventions, and outcome.",
        "Serious adverse events must be reported to the Sponsor's pharmacovigilance unit "
        "within 24 hours of investigator awareness, regardless of causality assessment "
        "or relationship to study drug. SAE reporting is performed on the designated SAE "
        "form and submitted via the validated electronic reporting system; paper "
        "submission is permitted only in the event of system unavailability and must be "
        "followed by electronic entry within 48 hours.",
        "The Sponsor will forward qualifying expedited SAEs to applicable regulatory "
        "authorities and to all participating investigators in accordance with ICH E2A "
        "and local expedited reporting requirements. The Sponsor will also distribute "
        "periodic Development Safety Update Reports (DSURs) to investigators and "
        "regulatory authorities on the schedule defined in the safety management plan.",
        "An independent Data Monitoring Committee (DMC) consisting of three external "
        "members — two clinical experts in medical oncology and one biostatistician — "
        "will review aggregate safety data at pre-specified intervals (after the first "
        "20 subjects have completed two cycles, after the first 40 subjects have "
        "completed four cycles, and quarterly thereafter). The DMC may recommend protocol "
        "modifications, dose adjustments, enrollment pause, or study suspension based on "
        "emerging safety signals.",
        "Pre-specified stopping rules include: more than 25% of enrolled subjects "
        "experiencing a Grade 4 or 5 treatment-related adverse event; any single "
        "treatment-related Grade 5 event meeting Hy's law criteria; or two or more "
        "treatment-related Grade 5 events of any cause. Triggering any stopping rule "
        "will result in immediate enrollment pause pending DMC review.",
        "Each site is responsible for local reporting to its Institutional Review Board "
        "or Ethics Committee in accordance with local requirements. Sites must also "
        "ensure that all study personnel involved in safety assessment have completed "
        "current Good Clinical Practice training and protocol-specific training prior "
        "to enrolling subjects.",
    ]),
    ("8. Statistical Analysis", [
        "The primary analysis population for efficacy is the Full Analysis Set (FAS), "
        "defined as all enrolled subjects who receive at least one dose of MX-209. The "
        "Safety Analysis Set is identical to the FAS for this protocol. The "
        "Per-Protocol Set, defined as FAS subjects without major protocol deviations, "
        "will be used for sensitivity analyses of the primary endpoint.",
        "The planned sample size of approximately 80 subjects provides adequate "
        "precision for estimating the primary endpoint. Assuming a true ORR of 25%, a "
        "sample size of 80 yields a two-sided 95% Clopper-Pearson confidence interval "
        "of approximately 16% to 36%, which is considered acceptable for inferring "
        "clinical benefit relative to a null rate of 10% based on historical "
        "performance of available salvage therapies.",
        "Objective response rate will be summarized with a two-sided 95% Clopper-Pearson "
        "exact confidence interval. Time-to-event endpoints (PFS, OS, DoR) will be "
        "summarized using Kaplan-Meier methodology, with median estimates and "
        "corresponding two-sided 95% confidence intervals computed using the "
        "Brookmeyer-Crowley method. Disease control rate will be summarized with a "
        "two-sided 95% Clopper-Pearson confidence interval. No formal hypothesis "
        "testing is planned in this single-arm study.",
        "Adverse events will be summarized by MedDRA system organ class and preferred "
        "term, by maximum CTCAE grade per subject, and by relationship to study drug. "
        "Treatment-emergent AEs are defined as any AE with onset on or after the first "
        "dose of MX-209 through 30 days after the last dose. Laboratory abnormalities "
        "will be summarized by maximum post-baseline grade and by shift tables relative "
        "to baseline grade. Concomitant medications will be coded using the WHO Drug "
        "Dictionary.",
        "Missing data for the primary endpoint will be handled by defining "
        "non-evaluable subjects as non-responders in the primary analysis. Sensitivity "
        "analyses will exclude non-evaluable subjects from the denominator. For "
        "time-to-event endpoints, subjects without an event at the analysis cutoff will "
        "be censored at the date of last adequate disease assessment.",
        "An interim analysis is planned after the first 40 enrolled subjects have "
        "completed at least four cycles of treatment, for the purpose of supporting "
        "regulatory interactions and informing potential expansion cohorts. The interim "
        "analysis is descriptive only and will not include formal stopping criteria "
        "for efficacy or futility; the study will continue to its planned total "
        "enrollment regardless of interim results.",
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
