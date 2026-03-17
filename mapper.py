import re
from difflib import SequenceMatcher


TEMPLATE_RULES = {
    "Product Summary": ["introduction", "intro", "overview", "summary"],
    "Value Proposition": ["added value", "value proposition", "value", "benefit"],
    "Product Description": ["standard services", "approach", "service description", "description"],
    "Requirements & Prerequisites": ["prerequisites", "requirements", "pre-requisite", "dependency"],
    "Scope / Out of Scope": ["out of scope", "scope", "included", "excluded"],
    "Key Differentiators": ["differentiator", "unique", "strength", "added value", "value"],
    "Operational Support": ["support", "operational", "operations", "maintenance", "run"],
    "SLA": ["service level", "sla", "availability", "response time"],
}
MAX_SUMMARY_LINES = 8
DEFAULT_REWRITE_PROFILE = "enterprise_strict"
REWRITE_PROFILES = {
    "enterprise_strict": {
        "lead": "Cegeka delivers this service as a standardized enterprise capability aligned to governance, consistency, and measurable outcomes.",
        "close": "The model reduces delivery variance, strengthens control over service quality, and supports predictable operational performance.",
    },
    "enterprise_balanced": {
        "lead": "Cegeka positions this service as a standardized operating model that combines business value with controlled delivery.",
        "close": "Customers gain repeatable implementation patterns, clear accountability, and scalable adoption across environments.",
    },
    "enterprise_concise": {
        "lead": "Cegeka provides a standardized service focused on control, reliability, and value realization.",
        "close": "The approach enables consistent outcomes, lower risk, and efficient scaling.",
    },
}


def summarize(text):
    """
    Placeholder for LLM-based summarization.
    Replace this implementation with an LLM call later.
    """
    return text


def _trim_to_max_lines(text, max_lines=MAX_SUMMARY_LINES):
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    if len(lines) <= max_lines:
        return "\n".join(lines)
    return "\n".join(lines[:max_lines])


def rewrite_commercial(text, profile=DEFAULT_REWRITE_PROFILE):
    """
    Rewrite Product Description in third-person enterprise language.
    Reusable profiles: enterprise_strict, enterprise_balanced, enterprise_concise.
    """
    raw = str(text or "").strip()
    if not raw:
        return ""

    selected_profile = REWRITE_PROFILES.get(profile, REWRITE_PROFILES[DEFAULT_REWRITE_PROFILE])

    normalized = re.sub(r"\s+", " ", raw)
    normalized = re.sub(r"\bwe\b", "Cegeka", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\bour\b", "Cegeka's", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\byou\b", "the customer", normalized, flags=re.IGNORECASE)
    normalized = re.sub(r"\byour\b", "the customer's", normalized, flags=re.IGNORECASE)

    return "\n".join(
        [
            selected_profile["lead"],
            normalized,
            selected_profile["close"],
        ]
    )


def _normalize(text):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", text.lower())).strip()


def _extract_titles_and_texts(sections):
    """Support both extract_sections() output and legacy flat extractors."""
    extracted = []
    for title, value in sections.items():
        if str(title).lower() == "tables":
            continue

        if isinstance(value, dict):
            section_title = value.get("section_title", title)
            section_text = value.get("section_text", "")
        else:
            section_title = title
            section_text = value if isinstance(value, str) else ""

        extracted.append((str(section_title), section_text))

    return extracted


def _build_section_lookup(sections):
    """Return a normalized lookup: title -> {section_title, section_text, tables}."""
    lookup = {}
    for title, value in sections.items():
        if str(title).lower() == "tables":
            continue

        if isinstance(value, dict):
            section_title = str(value.get("section_title", title))
            section_text = str(value.get("section_text", "") or "")
            tables = value.get("tables", []) or []
        else:
            section_title = str(title)
            section_text = str(value or "") if isinstance(value, str) else ""
            tables = []

        lookup[section_title] = {
            "section_title": section_title,
            "section_text": section_text,
            "tables": tables,
        }

    return lookup


def _tables_to_text(tables):
    """Render DOCX table payload to readable plain text."""
    blocks = []
    for table_index, table in enumerate(tables, start=1):
        blocks.append(f"Table {table_index}:")
        for row in table:
            cells = [str(cell).strip() for cell in row]
            blocks.append(" | ".join(cells))
    return "\n".join(blocks)


def _compose_section_content(section_payload, preserve_title=False, include_tables=False):
    """Return text content WITHOUT tables. Tables are handled separately by the caller."""
    if not section_payload:
        return ""

    parts = []
    if preserve_title:
        parts.append(section_payload["section_title"])

    text = section_payload.get("section_text", "").strip()
    if text:
        parts.append(text)

    return "\n\n".join(part for part in parts if part).strip()


def _get_section_tables(section_payload):
    """Extract the raw table payload from a section (not converted to text)."""
    if not section_payload:
        return []
    return section_payload.get("tables", []) or []


def _keyword_score(title, keywords, body_text=""):
    """Score a section against template-field keywords.

    Title matches score 0–1.0.
    Body-text keyword hits contribute a reduced bonus (max 0.35) so that
    a weak title but strong body text can surface the right section when
    the title match alone is inconclusive.
    """
    normalized_title = _normalize(title)
    title_words = normalized_title.split()
    best = 0.0

    for keyword in keywords:
        normalized_keyword = _normalize(keyword)
        if not normalized_keyword:
            continue

        # Exact substring in title → perfect score
        if normalized_keyword in normalized_title:
            best = max(best, 1.0)
            continue

        best = max(best, SequenceMatcher(None, normalized_title, normalized_keyword).ratio())

        for word in title_words:
            best = max(best, SequenceMatcher(None, word, normalized_keyword).ratio() * 0.9)

    # Body-text bonus: if a keyword appears literally in the section text,
    # add a bonus so content-rich sections can surface.
    # Multi-word keywords (e.g. "out of scope") carry a 0.50 bonus because
    # literal phrase presence is a strong signal; single-word keywords give 0.30.
    if body_text:
        normalized_body = _normalize(body_text)
        for keyword in keywords:
            normalized_keyword = _normalize(keyword)
            if normalized_keyword and normalized_keyword in normalized_body:
                bonus = 0.50 if " " in normalized_keyword else 0.30
                best = max(best, bonus)
                break

    return best


def get_match_diagnostics(sections):
    """
    Return per-template-field the best-matched section title and fuzzy score.
    Useful for pipeline transparency and debugging.

    Returns:
        dict: {
            "<template field>": {
                "matched_title": str | None,
                "score": float,
                "keywords": list[str]
            }, ...
        }
    """
    titles_and_texts = _extract_titles_and_texts(sections)
    diagnostics = {}

    for template_field, keywords in TEMPLATE_RULES.items():
        best_score = 0.0
        best_title = None

        for title, _text in titles_and_texts:
            score = _keyword_score(title, keywords, body_text=_text)
            if score >= best_score:
                best_score = score
                best_title = title

        diagnostics[template_field] = {
            "matched_title": best_title if best_score >= 0.45 else None,
            "score": round(best_score, 3),
            "keywords": keywords,
        }

    return diagnostics


def map_sd_to_template(
    sections,
    rewrite_profile=DEFAULT_REWRITE_PROFILE,
    full_section=False,
    preserve_titles=False,
    include_tables=False,
):
    """
    Map SD sections to template blocks using keyword + fuzzy title matching.

    Args:
        sections (dict): output from extract_sections(docx_file) or legacy extract_sd(path)
        rewrite_profile (str): commercial rewrite profile for Product Description
        full_section (bool): when True, return the full matched chapter content
        preserve_titles (bool): include the matched section title in output text
        include_tables (bool): include section tables in output text

    Returns:
        dict: template block -> mapped section text, plus _source_titles for tracing
    """
    titles_and_texts = _extract_titles_and_texts(sections)
    section_lookup = _build_section_lookup(sections)

    mapped = {template_field: "" for template_field in TEMPLATE_RULES}
    source_titles = {}  # Maps template_field -> source_title for formatting preservation
    for template_field, keywords in TEMPLATE_RULES.items():
        best_score = 0.0
        best_title = None
        best_text = ""
        best_score_with_content = 0.0
        best_title_with_content = None
        best_text_with_content = ""

        for title, text in titles_and_texts:
            score = _keyword_score(title, keywords, body_text=text)
            # Use >= so that among equal-scoring sections the *last* one in
            # document order wins. SD documents often go from generic to specific.
            if score >= best_score:
                best_score = score
                best_title = title
                best_text = text
            if text.strip() and score >= best_score_with_content:
                best_score_with_content = score
                best_title_with_content = title
                best_text_with_content = text

        source_title = best_title if best_score >= 0.45 else None
        source_text = best_text if best_score >= 0.45 else ""

        # If the best-matched section is empty (shell heading), fall back to
        # the best-scoring section that actually has content.
        if not source_text.strip() and best_score_with_content >= 0.80:
            source_title = best_title_with_content
            source_text = best_text_with_content

        # Track which source section matched this template field
        source_titles[template_field] = source_title

        if full_section:
            payload = section_lookup.get(source_title or "")
            source_text = _compose_section_content(
                payload,
                preserve_title=preserve_titles,
                include_tables=False,  # Keep tables separate; don't convert to text
            )
            if include_tables:
                # Return a dict with both text and tables when full_section + include_tables
                mapped[template_field] = {
                    "text": source_text,
                    "tables": _get_section_tables(payload) if payload else [],
                }
            else:
                mapped[template_field] = source_text
        else:
            summary = summarize(source_text)
            mapped[template_field] = _trim_to_max_lines(summary, MAX_SUMMARY_LINES)

    mapped["_source_titles"] = source_titles
    return mapped


def map_to_presales(data):
    """Backward-compatible mapping used by the current DOCX template generator."""
    mapped = map_sd_to_template(data)

    return {
        "PRODUCT_SUMMARY": mapped.get("Product Summary", ""),
        "VALUE_PROP": mapped.get("Value Proposition", ""),
        "DESCRIPTION": mapped.get("Product Description", ""),
        "REQUIREMENTS": mapped.get("Requirements & Prerequisites", ""),
        "SCOPE": mapped.get("Scope / Out of Scope", ""),
        "SLA": mapped.get("SLA", ""),
        "OPS_SUPPORT": mapped.get("Operational Support", ""),
    }