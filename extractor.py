import re

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


HEADING_PATTERN = re.compile(r"^\d+(?:\.\d+)*\.?\s+\S+")

# Stricter pattern for plain-text parsing: excludes numbered list items like "1. Item"
TEXT_HEADING_PATTERN = re.compile(r"^\d+(?:\.\d+)*\s+\S+")


def _iter_block_items(document):
    """Yield paragraphs and tables from the document body in their original order."""
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def _is_numbered_heading(paragraph_text):
    return bool(HEADING_PATTERN.match(paragraph_text.strip()))


def _table_to_rows(table):
    rows = []
    for row in table.rows:
        rows.append([cell.text.strip() for cell in row.cells])
    return rows


def extract_sections(docx_file):
    """
    Parse a DOCX into numbered sections and keep all tables with their section.

    Returns:
        dict: {
            "<section heading>": {
                "section_title": str,
                "section_text": str,
                "tables": list[list[list[str]]]
            },
            ...
        }
    """
    doc = Document(docx_file)

    sections = {}
    current_title = None
    text_chunks = []

    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue

            if _is_numbered_heading(text):
                current_title = text
                sections[current_title] = {
                    "section_title": current_title,
                    "section_text": "",
                    "tables": [],
                }
                text_chunks = []
                continue

            if current_title:
                text_chunks.append(text)
                sections[current_title]["section_text"] = "\n".join(text_chunks)

        elif isinstance(block, Table) and current_title:
            sections[current_title]["tables"].append(_table_to_rows(block))

    return sections


def _is_text_heading(line):
    """Heading check for plain-text SD documents. Does NOT match numbered list items like '1. item'."""
    return bool(TEXT_HEADING_PATTERN.match(line.strip()))


def extract_sections_from_text(text):
    """
    Parse a plain-text SD document into numbered sections.

    Headings recognised: '1 Title', '1.1 Title', '2.3.4 Title'.
    Lines like '1. List item' are intentionally NOT treated as headings.

    Returns:
        Same structure as extract_sections:
        { "<heading>": {"section_title", "section_text", "tables"}, ... }
    """
    sections = {}
    current_title = None
    text_chunks = []

    for line in str(text).splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        if _is_text_heading(stripped):
            if current_title is not None:
                sections[current_title]["section_text"] = "\n".join(text_chunks)
            current_title = stripped
            sections[current_title] = {
                "section_title": current_title,
                "section_text": "",
                "tables": [],
            }
            text_chunks = []
        elif current_title:
            text_chunks.append(stripped)

    # Flush the last open section.
    if current_title is not None and text_chunks:
        sections[current_title]["section_text"] = "\n".join(text_chunks)

    return sections


def extract_sd(path):
    """Backward-compatible extractor used by the current pipeline."""
    sections = extract_sections(path)

    data = {title: section["section_text"] for title, section in sections.items()}
    data["tables"] = []
    for section in sections.values():
        data["tables"].extend(section["tables"])

    return data