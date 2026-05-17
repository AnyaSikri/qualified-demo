"""ADAE Excel retrieval — load, filter by subject, project fields per template section."""


def load_adae(path):
    raise NotImplementedError


def get_subject_events(df, subject_id):
    raise NotImplementedError


def get_fields_for_section(events, section_config):
    raise NotImplementedError
