import json
import logging
import os
from typing import Dict, List, Optional

import ollama
import pandas as pd
from docx import Document


ROOT_FOLDER = r"C:\Users\koengo\Cegeka\Product Management - Product Management Library"
OUTPUT_FILE = "sd_chapter_classification.xlsx"
MODEL_NAME = "qwen2.5:3b-instruct"

# TEST MODE: Set to True to test with only first 2 documents
TEST_MODE = False  # Set to True for testing, False for full run
MAX_FILES_FOR_TEST = 2

ALLOWED_GROUPS = {
    "Executive Summary & Product Overview",
    "Scope Boundaries & Prerequisites",
    "Transition Operations & Governance",
    "Commercial & Risk Management",
    "Internal Presales Alignment",
}


def find_sd_files(root: str) -> List[str]:
    """Find all Service Description DOCX files (starting with SD) recursively under root."""
    docx_files: List[str] = []
    for dirpath, _, filenames in os.walk(root):
        for filename in filenames:
            normalized = filename.strip().lower()
            if normalized.endswith(".docx") and normalized.startswith("sd"):
                docx_files.append(os.path.join(dirpath, filename))
    return docx_files


def _is_heading_style(style_name: Optional[str]) -> bool:
    if not style_name:
        return False
    normalized = style_name.strip().lower()
    return normalized in {"heading 1", "heading 2", "heading 3"}


def extract_chapters(docxfile: str) -> List[Dict[str, str]]:
    """
    Extract chapters from a DOCX file.
    A chapter starts at Heading 1/2/3 and includes all following text
    until the next heading.
    """
    doc = Document(docxfile)
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
        else:
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


def _extract_json_substring(raw_text: str) -> Optional[str]:
    """
    Extract the first JSON-looking substring between the outermost
    curly braces. More robust handling of variations.
    """
    # Remove common markdown code blocks
    text_cleaned = raw_text.replace("```json", "").replace("```", "")
    
    start = text_cleaned.find("{")
    end = text_cleaned.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return None
    return text_cleaned[start : end + 1]


def _parse_classification_json(raw_text: str) -> Optional[Dict[str, str]]:
    """
    Robust parser:
    1) direct JSON parse
    2) fallback: extract substring between { ... } and retry
    3) Handle common issues (trailing commas, quotes, etc.)
    """
    try:
        parsed = json.loads(raw_text)
        if isinstance(parsed, dict):
            return parsed
    except Exception as e:
        logging.debug("Direct JSON parse failed: %s", str(e)[:100])
        pass

    json_substring = _extract_json_substring(raw_text)
    if not json_substring:
        logging.debug("No JSON substring found in: %s", raw_text[:100])
        return None

    try:
        parsed = json.loads(json_substring)
        if isinstance(parsed, dict):
            return parsed
    except Exception as e:
        logging.debug("JSON substring parse failed: %s | Substring: %s", str(e)[:100], json_substring[:100])
        return None

    return None


def classify_with_ollama(title: str, text: str) -> Dict[str, str]:
    """
    Classify chapter content using local Ollama model.
    Expected output JSON:
    {"group":"...", "reason":"..."}
    """
    prompt = f"""
You classify chapters from service description documents.

Allowed groups (choose ONE exactly as written):
- "Executive Summary & Product Overview"
- "Scope Boundaries & Prerequisites"
- "Transition Operations & Governance"
- "Commercial & Risk Management"
- "Internal Presales Alignment"

Return ONLY valid JSON with this exact schema:
{{
  "group": "...",
  "reason": "..."
}}

IMPORTANT: The group value MUST be EXACTLY one of the options above without numbering.

Chapter title:
{title}

Chapter text:
{text}
""".strip()

    try:
        response = ollama.chat(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            options={
                "temperature": 0,
            },
        )

        # Handle ChatResponse object (not a dict)
        # response.message.content contains the actual response
        try:
            raw_content = response.message.content if hasattr(response, 'message') else ""
        except Exception as e:
            logging.error("Failed to extract content from response: %s", e)
            raw_content = ""

        if not raw_content or not raw_content.strip():
            logging.warning("Empty response from Ollama for title: %s", title)
            return {"group": "ERROR", "reason": "Empty response from Ollama"}

        logging.debug("Raw Ollama response for '%s': %s", title[:50], raw_content[:200])

        parsed = _parse_classification_json(raw_content)
        if not parsed:
            logging.warning("Failed to parse JSON for title: %s | Response: %s", title[:50], raw_content[:200])
            return {"group": "ERROR", "reason": f"JSON parse failed"}

        group = str(parsed.get("group", "")).strip()
        reason = str(parsed.get("reason", "")).strip()

        # Remove numbering prefix if present (e.g., "1. Executive Summary..." -> "Executive Summary...")
        if group and group[0].isdigit():
            # Remove leading "N. " pattern
            parts = group.split(". ", 1)
            if len(parts) > 1:
                group = parts[1]

        if group not in ALLOWED_GROUPS:
            logging.warning("Group '%s' not in allowed groups for title: %s", group, title[:50])
            # Find similar group
            similar = [g for g in ALLOWED_GROUPS if any(word in group.lower() for word in g.lower().split())]
            if similar:
                logging.info("Using similar group '%s' instead of '%s'", similar[0], group)
                group = similar[0]
            else:
                return {"group": "ERROR", "reason": f"Invalid group: {group}"}

        if not reason:
            reason = "No reason provided"

        return {"group": group, "reason": reason}

    except Exception as exc:
        logging.error("Ollama classification failed: %s", exc)
        return {"group": "ERROR", "reason": f"Exception: {str(exc)[:100]}"}


def main() -> None:
    logging.basicConfig(
        level=logging.DEBUG,  # Changed from INFO to DEBUG
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    logging.info("Searching for DOCX files under: %s", ROOT_FOLDER)
    docx_files = find_sd_files(ROOT_FOLDER)
    logging.info("Found %d DOCX files", len(docx_files))
    
    if TEST_MODE:
        docx_files = docx_files[:MAX_FILES_FOR_TEST]
        logging.info("TEST MODE: Processing only first %d files", len(docx_files))

    rows: List[Dict[str, str]] = []

    for idx, docx_path in enumerate(docx_files, start=1):
        sd_name = os.path.basename(docx_path)
        logging.info("[%d/%d] Processing: %s", idx, len(docx_files), docx_path)

        try:
            chapters = extract_chapters(docx_path)
        except Exception as exc:
            logging.error("Skipping corrupt/unreadable document: %s | %s", docx_path, exc)
            continue

        if not chapters:
            logging.info("No Heading 1/2/3 chapters found in: %s", docx_path)
            continue

        for chapter in chapters:
            title = chapter.get("title", "")
            full_text = chapter.get("text", "")

            classification = classify_with_ollama(title, full_text)

            rows.append(
                {
                    "SD Name": sd_name,
                    "Full Path": docx_path,
                    "Chapter Title": title,
                    "Chapter Text (eerste 500 chars)": full_text[:500],
                    "Classification Group": classification.get("group", "ERROR"),
                    "Reason": classification.get("reason", "Failed parsing"),
                }
            )

    df = pd.DataFrame(
        rows,
        columns=[
            "SD Name",
            "Full Path",
            "Chapter Title",
            "Chapter Text (eerste 500 chars)",
            "Classification Group",
            "Reason",
        ],
    )

    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    logging.info("Done. Wrote %d rows to %s", len(df), OUTPUT_FILE)


if __name__ == "__main__":
    main()
