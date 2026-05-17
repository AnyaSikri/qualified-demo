"""Protocol PDF retrieval — load PDF, parse numbered sections, fetch sections by title."""


def load_protocol(path):
    raise NotImplementedError


def parse_sections(pages):
    raise NotImplementedError


def get_sections_for(parsed, section_titles):
    raise NotImplementedError
