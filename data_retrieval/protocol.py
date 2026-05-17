"""Protocol PDF retrieval.

Loads the protocol PDF, parses numbered sections by regex, and returns the
text + page span for the section titles a template section asks for.

Structural retrieval (not semantic): for a single, well-structured document
where citations must be deterministic, regex on numbered headers gives 100%
precise page references with no embedding store to maintain.
"""

import re

import fitz  # PyMuPDF


# Matches headers like "5. Safety Monitoring" appearing on their own line.
# Capture groups: section number, title.
SECTION_HEADER_RE = re.compile(r"^\s*(\d+)\.\s+([A-Z][^\n]{2,80})\s*$", re.MULTILINE)


def load_protocol(path):
    """Return a list of page texts (1-indexed by list position + 1)."""
    doc = fitz.open(path)
    pages = [page.get_text() for page in doc]
    doc.close()
    return pages


def parse_sections(pages):
    """Walk pages once, return list of {number, title, pages, text}.

    `pages` is a list of per-page strings (output of load_protocol).
    Each section's `pages` is a list of 1-indexed page numbers it spans
    (a section starts on the page where its header appears and continues
    until the next header — usually one page since we emit a PageBreak
    after each section, but the code handles multi-page sections too).
    """
    # Collect every header occurrence with its page number.
    headers = []  # list of (page_idx, char_offset_within_page, number, title)
    for page_idx, page_text in enumerate(pages):
        for match in SECTION_HEADER_RE.finditer(page_text):
            headers.append({
                "page": page_idx + 1,
                "offset": match.start(),
                "number": int(match.group(1)),
                "title": match.group(2).strip(),
            })

    # Walk pairs of headers to slice out section bodies.
    sections = []
    for i, header in enumerate(headers):
        start_page = header["page"]
        start_offset = header["offset"]
        if i + 1 < len(headers):
            next_header = headers[i + 1]
            end_page = next_header["page"]
            end_offset = next_header["offset"]
        else:
            end_page = len(pages)
            end_offset = len(pages[-1])

        # Stitch text from start_page..end_page, trimming the first and last.
        if start_page == end_page:
            body = pages[start_page - 1][start_offset:end_offset]
            span_pages = [start_page]
        else:
            chunks = [pages[start_page - 1][start_offset:]]
            for p in range(start_page, end_page - 1):
                chunks.append(pages[p])
            chunks.append(pages[end_page - 1][:end_offset])
            body = "".join(chunks)
            span_pages = list(range(start_page, end_page + 1))
            # If the next header sits at the very top of its page, the
            # current section doesn't actually extend into that page.
            if end_offset == 0:
                span_pages = span_pages[:-1]

        sections.append({
            "number": header["number"],
            "title": header["title"],
            "pages": span_pages,
            "text": body.strip(),
        })

    return sections


def _normalize(s):
    return re.sub(r"\s+", " ", s).strip().lower()


def get_sections_for(parsed, section_titles):
    """Return the parsed sections whose title matches any requested title.

    Matching is case- and whitespace-insensitive. Titles that don't match
    anything are silently skipped — callers can detect this by comparing
    requested vs returned length if they want to flag missing sections.
    """
    if not section_titles:
        return []
    wanted = {_normalize(t) for t in section_titles}
    return [s for s in parsed if _normalize(s["title"]) in wanted]


# ---------------------------------------------------------------------------
# Standalone smoke test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import os
    import sys

    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from template import NARRATIVE_TEMPLATE  # noqa: E402

    here = os.path.dirname(os.path.abspath(__file__))
    pdf_path = os.path.join(here, "..", "data", "mock_protocol.pdf")

    print(f"loading {pdf_path}")
    pages = load_protocol(pdf_path)
    print(f"  {len(pages)} pages")

    parsed = parse_sections(pages)
    print(f"  parsed {len(parsed)} sections:")
    for s in parsed:
        snippet = s["text"].split("\n", 1)[-1][:80].replace("\n", " ")
        print(f"    {s['number']}. {s['title']}  pages={s['pages']}  '{snippet}...'")

    print("\nlookup per template section:")
    for section in NARRATIVE_TEMPLATE["sections"]:
        hits = get_sections_for(parsed, section["protocol_sections"])
        labels = [f"{h['title']} (p.{','.join(map(str, h['pages']))})" for h in hits]
        print(f"  {section['id']}: requested={section['protocol_sections']!r} -> {labels}")
