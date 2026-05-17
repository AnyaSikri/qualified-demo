NARRATIVE_TEMPLATE = {
    "name": "Standard SAE Narrative",
    "version": "1.0",
    "sections": [
        {
            "id": "patient_info",
            "title": "Patient Information",
            "instructions": "In 1-2 sentences, describe the subject's demographics and study assignment.",
            "adae_fields": ["USUBJID", "AGE", "SEX", "RACE"],
            "protocol_sections": ["Eligibility"],
        },
        {
            "id": "event_description",
            "title": "Event Description",
            "instructions": "Describe the adverse event: onset date, severity grade, MedDRA preferred term, seriousness criteria. Preserve MedDRA terms verbatim.",
            "adae_fields": ["AEDECOD", "AESEV", "AESTDTC", "AEENDTC", "AESER", "AETOXGR"],
            "protocol_sections": ["Adverse Event Definitions", "Safety Monitoring"],
        },
        {
            "id": "treatment_outcome",
            "title": "Treatment & Outcome",
            "instructions": "Describe actions taken (drug interruption, dose modification), interventions, and event resolution.",
            "adae_fields": ["AEACN", "AEOUT", "AETOXGR"],
            "protocol_sections": ["Safety Monitoring"],
        },
        {
            "id": "conmeds",
            "title": "Concomitant Medications",
            "instructions": "List concomitant medications relevant to the event timeframe.",
            "adae_fields": ["CMTRT", "CMDOSE", "CMSTDTC"],
            "protocol_sections": [],
        },
    ],
}
