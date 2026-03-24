import re
from typing import Dict, List, Optional

from docx import Document


# Numbered heading pattern e.g. "1. Intro" or "2.3 Details"
_NUMBERED_HEADING_RE = re.compile(r"^\d+(?:\.\d+)*\.?\s+\S+")


def _is_heading_style(style_name: Optional[str]) -> bool:
    if not style_name:
        return False
    normalized = style_name.strip().lower()
    # English (Heading 1), Dutch (Kop 1), and any other Heading variant
    return (
        normalized in {"heading 1", "heading 2", "heading 3"}
        or normalized in {"kop 1", "kop 2", "kop 3"}
        or normalized.startswith("heading ")
    )


def extract_chapters(docxfile) -> List[Dict[str, str]]:
    """
    Extract chapters from a DOCX file.
    A chapter starts with Heading 1/2/3 and includes all subsequent text
    until the next heading.
    """
    doc = Document(str(docxfile))
    chapters: List[Dict[str, str]] = []

    current_title: Optional[str] = None
    current_lines: List[str] = []

    for para in doc.paragraphs:
        text = para.text.strip()
        style_name = para.style.name if para.style is not None else None

        if _is_heading_style(style_name):
            if current_title is not None:
                chapters.append(
                    {
                        "title": current_title,
                        "text": "\n".join(current_lines).strip(),
                    }
                )

            current_title = text if text else "[UNTITLED HEADING]"
            current_lines = []
            continue

        if current_title is not None and text:
            current_lines.append(text)

    if current_title is not None:
        chapters.append(
            {
                "title": current_title,
                "text": "\n".join(current_lines).strip(),
            }
        )

    return chapters
