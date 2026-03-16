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


def _keyword_score(title, keywords):
    """Combine hard keyword hits with fuzzy similarity for flexible matching."""
    normalized_title = _normalize(title)
    title_words = normalized_title.split()
    best = 0.0

    for keyword in keywords:
        normalized_keyword = _normalize(keyword)
        if not normalized_keyword:
            continue

        if normalized_keyword in normalized_title:
            best = max(best, 1.0)
            continue

        best = max(best, SequenceMatcher(None, normalized_title, normalized_keyword).ratio())

        for word in title_words:
            best = max(best, SequenceMatcher(None, word, normalized_keyword).ratio() * 0.9)

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
            score = _keyword_score(title, keywords)
            if score >= best_score:
                best_score = score
                best_title = title

        diagnostics[template_field] = {
            "matched_title": best_title if best_score >= 0.45 else None,
            "score": round(best_score, 3),
            "keywords": keywords,
        }

    return diagnostics


def map_sd_to_template(sections, rewrite_profile=DEFAULT_REWRITE_PROFILE):
    """
    Map SD sections to template blocks using keyword + fuzzy title matching.

    Args:
        sections (dict): output from extract_sections(docx_file) or legacy extract_sd(path)
        rewrite_profile (str): commercial rewrite profile for Product Description

    Returns:
        dict: template block -> summarized matched section text (max 8 lines)
    """
    titles_and_texts = _extract_titles_and_texts(sections)

    mapped = {template_field: "" for template_field in TEMPLATE_RULES}
    for template_field, keywords in TEMPLATE_RULES.items():
        best_score = 0.0
        best_text = ""

        for title, text in titles_and_texts:
            score = _keyword_score(title, keywords)
            # Use >= so that among equal-scoring sections the *last* one in
            # document order wins.  SD documents go from general (top-level
            # headings) to specific (numbered sub-sections), so the deepest
            # relevant section reliably replaces a shallow one with the same score.
            if score >= best_score:
                best_score = score
                best_text = text

        # Keep weak/accidental fuzzy hits out of the final mapping.
        source_text = best_text if best_score >= 0.45 else ""

        if template_field == "Product Description":
            source_text = rewrite_commercial(source_text, profile=rewrite_profile)

        summary = summarize(source_text)
        mapped[template_field] = _trim_to_max_lines(summary, MAX_SUMMARY_LINES)

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