import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from datetime import datetime
import difflib
import re
import csv
import shutil
import hashlib
import struct
import json

from docx import Document
from docx.shared import Inches

LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)
GOLD_EXAMPLES_FILE = Path("rules") / "gold_examples.json"


def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(line)
    print(line, end="")


# =====================================================================
#  1:1 mapping met Presales Guide template-secties
# =====================================================================
MAPPING = {
    # 1. Product Summary
    "service introduction": "PRODUCT_SUMMARY",
    "service identification": "PRODUCT_SUMMARY",
    "service reporting": "PRODUCT_SUMMARY",
    "service window": "PRODUCT_SUMMARY",
    "product summary": "PRODUCT_SUMMARY",

    # 2. Understanding the Client's Needs
    "service overview": "CLIENT_NEEDS",
    "goals": "CLIENT_NEEDS",
    "service target audience": "CLIENT_NEEDS",
    "understanding clients needs": "CLIENT_NEEDS",
    "understanding client's needs": "CLIENT_NEEDS",

    # 3. Product Description
    "product description": "PRODUCT_DESCRIPTION",
    "service description": "PRODUCT_DESCRIPTION",
    "services": "PRODUCT_DESCRIPTION",
    "standard services": "PRODUCT_DESCRIPTION",
    "optional services": "PRODUCT_DESCRIPTION",

    # 3.1 Architectural Description
    "architectural description": "ARCHITECTURAL_DESCRIPTION",
    "technical architecture": "ARCHITECTURAL_DESCRIPTION",
    "technical implementation": "ARCHITECTURAL_DESCRIPTION",
    "architecture": "ARCHITECTURAL_DESCRIPTION",

    # 3.2 Key Features & Functionalities
    "key features": "KEY_FEATURES",
    "key features and functionalities": "KEY_FEATURES",
    "operational readiness": "KEY_FEATURES",
    "run services": "KEY_FEATURES",
    "management services": "KEY_FEATURES",
    "governance and reporting": "KEY_FEATURES",
    "process": "KEY_FEATURES",

    # 3.3 Scope / Out-of-Scope
    "scope": "SCOPE",
    "out_of_scope": "SCOPE",
    "out of scope": "SCOPE",
    "scope and out of scope": "SCOPE",

    # 3.4 Requirements & Prerequisites
    "requirements": "REQUIREMENTS",
    "eligibility and prerequisites": "REQUIREMENTS",
    "eligibility & prerequisites": "REQUIREMENTS",
    "prerequisites": "REQUIREMENTS",

    # 4. Value Proposition
    "value proposition": "VALUE_PROPOSITION",
    "value and benefits": "VALUE_PROPOSITION",
    "value & benefits": "VALUE_PROPOSITION",

    # 5. Key Differentiators
    "differentiators": "DIFFERENTIATORS",
    "key differentiators": "DIFFERENTIATORS",

    # 6. Transition & Transformation
    "transition services": "TRANSITION_TRANSFORMATION",
    "transition and transformation": "TRANSITION_TRANSFORMATION",
    "transition & transformation": "TRANSITION_TRANSFORMATION",

    # 7. Client Responsibilities
    "client responsibilities": "CLIENT_RESPONSIBILITIES",
    "customer responsibilities": "CLIENT_RESPONSIBILITIES",
    "cegeka responsibilities": "CLIENT_RESPONSIBILITIES",

    # 8. Operational Support
    "operational support": "OPERATIONAL_SUPPORT",
    "support model": "OPERATIONAL_SUPPORT",
    "incident management": "OPERATIONAL_SUPPORT",
    "incident response": "OPERATIONAL_SUPPORT",

    # 9. Terms & Conditions
    "terms": "TERMS_CONDITIONS",
    "terms and conditions": "TERMS_CONDITIONS",
    "conditions": "TERMS_CONDITIONS",

    # 10. SLA & KPI Management
    "service level": "SLA_KPI",
    "service levels": "SLA_KPI",
    "sla": "SLA_KPI",
    "kpi": "SLA_KPI",
    "availability": "SLA_KPI",

    # 11. Pricing Elements
    "pricing": "PRICING_ELEMENTS",
    "pricing elements": "PRICING_ELEMENTS",
    "service billing": "PRICING_ELEMENTS",
    "optional service billing": "PRICING_ELEMENTS",
    "change request billing": "PRICING_ELEMENTS",
    "service request billing": "PRICING_ELEMENTS",
    "delivery model": "PRICING_ELEMENTS",
}


CATEGORY_TO_SDT = {
    "PRODUCT_SUMMARY": "PRODUCT_SUMMARY",
    "CLIENT_NEEDS": "CLIENT_NEEDS",
    "PRODUCT_DESCRIPTION": "PRODUCT_DESCRIPTION",
    "ARCHITECTURAL_DESCRIPTION": "ARCHITECTURAL_DESCRIPTION",
    "KEY_FEATURES": "KEY_FEATURES",
    "SCOPE": "SCOPE",
    "REQUIREMENTS": "REQUIREMENTS",
    "VALUE_PROPOSITION": "VALUE_PROPOSITION",
    "DIFFERENTIATORS": "DIFFERENTIATORS",
    "TRANSITION_TRANSFORMATION": "TRANSITION_TRANSFORMATION",
    "CLIENT_RESPONSIBILITIES": "CLIENT_RESPONSIBILITIES",
    "OPERATIONAL_SUPPORT": "OPERATIONAL_SUPPORT",
    "TERMS_CONDITIONS": "TERMS_CONDITIONS",
    "SLA_KPI": "SLA_KPI",
    "PRICING_ELEMENTS": "PRICING_ELEMENTS",
}


TAG_SIGNALS = {
    "PRODUCT_SUMMARY": ["introduction", "overview", "summary", "service identification", "service model"],
    "CLIENT_NEEDS": ["need", "goal", "challenge", "target audience", "business outcome"],
    "PRODUCT_DESCRIPTION": ["description", "standard services", "optional services", "application", "platform"],
    "ARCHITECTURAL_DESCRIPTION": ["architecture", "technical", "design", "implementation"],
    "KEY_FEATURES": ["feature", "function", "capability", "management", "governance"],
    "SCOPE": ["scope", "in scope", "out of scope", "included", "excluded"],
    "REQUIREMENTS": ["requirement", "prerequisite", "dependency", "eligibility"],
    "VALUE_PROPOSITION": ["value", "benefit", "outcome", "impact"],
    "DIFFERENTIATORS": ["differentiat", "unique", "strength"],
    "TRANSITION_TRANSFORMATION": ["transition", "transformation", "onboarding", "migration"],
    "CLIENT_RESPONSIBILITIES": ["responsibil", "customer", "client", "provided by customer"],
    "OPERATIONAL_SUPPORT": ["support", "incident", "request", "problem", "operation"],
    "TERMS_CONDITIONS": ["terms", "conditions", "contract", "limitation"],
    "SLA_KPI": ["sla", "kpi", "availability", "service level", "response"],
    "PRICING_ELEMENTS": ["pricing", "billing", "price", "cost", "charge"],

    # Section 1: Presales instructions & checks
    "PRESALES_INSTRUCTIONS": ["presales", "qualification", "sales approach", "go to market", "offering guidance"],
    "CEGEKA_CONTACTS": ["contact", "owner", "product manager", "service manager", "architect", "escalation"],
    "PRESALES_CHECKS": ["check", "checklist", "must", "precondition", "verify", "validation"],
    "SKU_INFORMATION": ["sku", "part number", "license", "licensing", "catalog", "billable item"],
    "OTHER_CONDITIONAL_SOLUTIONS": ["dependency", "dependent", "requires", "prerequisite solution", "related solution"],
    "QA_CUSTOMERS": ["faq", "question", "answer", "customer ask", "common question"],

    # Section 2 and sub-sections
    "OFFER_SECTIONS": ["offer", "proposal", "scope", "solution section", "statement of work"],
    "TRANSITION_PROJECT_DESC": ["transition", "project", "onboarding", "implementation approach", "rollout"],
    "TRANSITION_SCOPE": ["in scope", "out of scope", "project scope", "deliverable", "not included"],
    "TRANSITION_ASSUMPTIONS": ["assumption", "dependency", "constraint", "premise"],
    "TRANSITION_MILESTONES": ["milestone", "timeline", "planning", "phase", "deliverable", "target date"],
    "ROLES_CEGEKA": ["cegeka", "responsible", "role", "team", "raci", "accountable"],
    "ROLES_CUSTOMER": ["customer", "client", "responsible", "role", "raci", "accountable"],
    "CLIENT_RESPONSIBILITIES_2": ["customer responsibility", "client responsibility", "customer provides", "provided by customer"],
    "ASSUMPTIONS_RISKS": ["assumption", "risk", "mitigation", "dependency", "constraint"],
    "ACCEPTANCE_CRITERIA": ["acceptance criteria", "acceptance", "entry criteria", "exit criteria", "definition of done"],

    # Pricing sub-sections and other docs
    "COST_ONE_TIME": ["one time", "setup", "implementation fee", "project cost", "initial cost"],
    "COST_RECURRING": ["recurring", "monthly", "yearly", "subscription", "run cost", "opex"],
    "CHARGING_MECHANISM": ["charging", "billing model", "consumption", "unit", "invoice", "chargeback"],
    "OTHER_DOCUMENTS": ["appendix", "attachment", "annex", "reference document", "supporting document"],
    "COMMERCIAL_SHEET": ["commercial", "price sheet", "rate card", "commercial sheet", "quote"],
    "SERVICE_DESCRIPTION_LINK": ["service description", "link", "reference", "document location", "repository"],
}

TAG_TABLE_FACT_TYPES = {
    "SLA_KPI": {"service_level"},
    "PRICING_ELEMENTS": {"pricing"},
    "OPERATIONAL_SUPPORT": {"operations"},
    "SCOPE": {"scope"},
}


IGNORED_SECTION_HINTS = [
    "table of contents",
    "document history",
    "document control",
    "version history",
    "approval",
    "distribution",
    "appendix",
    "glossary",
]

HITL_PREFIX = "AI generated, teverifieren door HITL"
LOW_INFO_TEXT = "AI agent heeft te weinig info om dit zelf op te stellen"
HITL_VALIDATION_NOTICE = "Deze tekst is AI generated en moet inhoudelijk gevalideerd worden door HITL."
AI_DECORATED_FILL_TYPES = {"ai_related_documents", "ai_missing_chapter", "open_too_little_info"}

TEMPLATE_TAG_ORDER = [
    "PRODUCT_SUMMARY",
    "CLIENT_NEEDS",
    "PRODUCT_DESCRIPTION",
    "ARCHITECTURAL_DESCRIPTION",
    "KEY_FEATURES",
    "SCOPE",
    "REQUIREMENTS",
    "VALUE_PROPOSITION",
    "DIFFERENTIATORS",
    "TRANSITION_TRANSFORMATION",
    "CLIENT_RESPONSIBILITIES",
    "OPERATIONAL_SUPPORT",
    "TERMS_CONDITIONS",
    "SLA_KPI",
    "PRICING_ELEMENTS",
]

LOW_CONFIDENCE_TAGS = {
    "SLA_KPI",
    "PRICING_ELEMENTS",
}

STRICT_EVIDENCE_TAGS = {
    "SLA_KPI",
    "PRICING_ELEMENTS",
    "CLIENT_RESPONSIBILITIES",
    "TERMS_CONDITIONS",
}

RELATED_TAGS = {
    "PRODUCT_SUMMARY": ["PRODUCT_DESCRIPTION", "KEY_FEATURES", "VALUE_PROPOSITION"],
    "CLIENT_NEEDS": ["PRODUCT_SUMMARY", "PRODUCT_DESCRIPTION", "KEY_FEATURES"],
    "PRODUCT_DESCRIPTION": ["PRODUCT_SUMMARY", "KEY_FEATURES", "ARCHITECTURAL_DESCRIPTION"],
    "ARCHITECTURAL_DESCRIPTION": ["PRODUCT_DESCRIPTION", "KEY_FEATURES"],
    "KEY_FEATURES": ["PRODUCT_DESCRIPTION", "VALUE_PROPOSITION"],
    "SCOPE": ["PRODUCT_DESCRIPTION", "KEY_FEATURES", "REQUIREMENTS"],
    "REQUIREMENTS": ["PRODUCT_DESCRIPTION", "ARCHITECTURAL_DESCRIPTION", "SCOPE"],
    "VALUE_PROPOSITION": ["PRODUCT_SUMMARY", "KEY_FEATURES", "PRODUCT_DESCRIPTION"],
    "DIFFERENTIATORS": ["KEY_FEATURES", "VALUE_PROPOSITION", "PRODUCT_DESCRIPTION"],
    "TRANSITION_TRANSFORMATION": ["PRODUCT_DESCRIPTION", "CLIENT_RESPONSIBILITIES", "OPERATIONAL_SUPPORT"],
    "CLIENT_RESPONSIBILITIES": ["PRODUCT_DESCRIPTION", "SCOPE", "REQUIREMENTS"],
    "OPERATIONAL_SUPPORT": ["KEY_FEATURES", "SLA_KPI", "CLIENT_RESPONSIBILITIES"],
    "TERMS_CONDITIONS": ["SCOPE", "CLIENT_RESPONSIBILITIES", "OPERATIONAL_SUPPORT"],
}

TAG_LABEL_NL = {
    "PRODUCT_SUMMARY": "Productsamenvatting",
    "CLIENT_NEEDS": "Klantbehoeften",
    "PRODUCT_DESCRIPTION": "Productbeschrijving",
    "ARCHITECTURAL_DESCRIPTION": "Architecturale beschrijving",
    "KEY_FEATURES": "Belangrijkste functionaliteiten",
    "SCOPE": "Scope",
    "REQUIREMENTS": "Vereisten en randvoorwaarden",
    "VALUE_PROPOSITION": "Waardepropositie",
    "DIFFERENTIATORS": "Differentiatoren",
    "TRANSITION_TRANSFORMATION": "Transitie en transformatie",
    "CLIENT_RESPONSIBILITIES": "Klantverantwoordelijkheden",
    "OPERATIONAL_SUPPORT": "Operationele ondersteuning",
    "TERMS_CONDITIONS": "Voorwaarden",
    "SLA_KPI": "SLA en KPI",
    "PRICING_ELEMENTS": "Prijscomponenten",
}

NON_CHAPTER_TAGS = {
    "Customer",
}

SUPPORTED_RELATED_EXTENSIONS = {
    ".docx",
    ".txt",
    ".md",
    ".html",
    ".htm",
    ".csv",
    ".pptx",
    ".xlsx",
}

MAX_RELATED_FILES = 120
MAX_CHARS_PER_RELATED_FILE = 24000
MAX_IMAGES_PER_FILE = 3
MAX_IMAGES_PER_TAG = 1
MAX_SOURCE_SECTIONS_PER_TAG = 3
MAX_SOURCE_CONTENT_CHARS = 2200

SEVERITY_INFO = "info"
SEVERITY_WARNING = "warning"
SEVERITY_BLOCKING = "blocking"
QUALITY_LOW_SCORE_THRESHOLD = 60

INTENT_MUST_HAVE_SIGNALS = {
    "PRODUCT_SUMMARY": ["service", "summary", "overview", "identification"],
    "CLIENT_NEEDS": ["need", "goal", "challenge", "business"],
    "PRODUCT_DESCRIPTION": ["description", "application", "service", "platform"],
    "ARCHITECTURAL_DESCRIPTION": ["architecture", "technical", "design", "platform"],
    "KEY_FEATURES": ["feature", "function", "capability", "service"],
    "SCOPE": ["scope", "included", "excluded", "out of scope", "in scope"],
    "REQUIREMENTS": ["requirement", "prerequisite", "dependency", "eligibility"],
    "VALUE_PROPOSITION": ["value", "benefit", "outcome", "impact"],
    "DIFFERENTIATORS": ["unique", "differentiat", "advantage", "strength"],
    "TRANSITION_TRANSFORMATION": ["transition", "transformation", "onboarding", "migration"],
    "CLIENT_RESPONSIBILITIES": ["responsibil", "customer", "client", "raci"],
    "OPERATIONAL_SUPPORT": ["support", "incident", "request", "operation", "service window"],
    "TERMS_CONDITIONS": ["terms", "conditions", "contract", "limitation"],
    "SLA_KPI": ["sla", "kpi", "availability", "service level", "response"],
    "PRICING_ELEMENTS": ["pricing", "billing", "cost", "price", "charge"],
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}


# =====================================================================
#  SEMANTISCHE MATCHING ENGINE
# =====================================================================

def normalize(s):
    return s.strip().lower().replace("_", " ").replace("&", "and")


def clean_section_numbering(s):
    s = s.strip()
    s = re.sub(r"^\d+(?:\.\d+)*\s*[-.)]?\s*", "", s)
    return s


def resolve_sdt_tag(section_name):
    section_norm = normalize(clean_section_numbering(section_name))

    # Deterministic exact match first (strict 1:1 behavior)
    if section_norm in MAPPING:
        return MAPPING[section_norm]

    # Fallback: only use high-confidence fuzzy match as compatibility bridge.
    best_key = None
    best_ratio = 0.0
    for key in MAPPING.keys():
        ratio = difflib.SequenceMatcher(None, section_norm, normalize(key)).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_key = key

    if best_key and best_ratio >= 0.90:
        return MAPPING[best_key]

    return None


def looks_irrelevant_section(section_name):
    name_norm = normalize(clean_section_numbering(section_name))
    return any(hint in name_norm for hint in IGNORED_SECTION_HINTS)


def _looks_toc_noise_content(content):
    text = str(content or "")
    if not text:
        return False

    low = normalize(text)
    if "table of contents" not in low:
        return False

    numbered_heads = len(re.findall(r"\b\d+(?:\.\d+)+\b", text))
    return numbered_heads >= 5


def _strip_toc_noise(content):
    text = str(content or "").strip()
    if not text:
        return ""

    # Remove common TOC line patterns when they appear in extracted body text.
    text = re.sub(r"(?i)table\s+of\s+contents", "", text)
    text = re.sub(r"\b\d+(?:\.\d+)+\s+[A-Za-z][^\n\r]{0,120}?\s+\d{1,3}\b", "", text)
    text = re.sub(r"\b\d+\s+[A-Za-z][^\n\r]{0,120}?\s+\d{1,3}\b", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def is_relevant_for_tag(tag_name, section_name, section_content):
    if not section_content or len(section_content.strip()) < 30:
        return False
    if _looks_toc_noise_content(section_content):
        return False

    combined = normalize(f"{section_name} {section_content[:1200]}")
    signals = TAG_SIGNALS.get(tag_name, [])

    if not signals:
        return True

    return any(sig in combined for sig in signals)


def resolve_tag_for_section(section_name, section_category):
    category_norm = normalize(section_category).upper().replace(" ", "_") if section_category else ""
    if category_norm in CATEGORY_TO_SDT:
        return CATEGORY_TO_SDT[category_norm], "category"

    tag = resolve_sdt_tag(section_name)
    if tag:
        return tag, "header"

    return None, "none"


def sanitize_xml_text(text):
    # Remove XML 1.0 illegal control chars that can make Word parts unreadable.
    if text is None:
        return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def _clip(text, max_len=260):
    cleaned = " ".join(str(text or "").split())
    if len(cleaned) <= max_len:
        return cleaned
    return cleaned[: max_len - 3].rstrip() + "..."


def _build_low_info_text():
    return f"{HITL_PREFIX}\n{LOW_INFO_TEXT}"


def _is_generated_fallback_text(text):
    t = str(text or "")
    return (HITL_PREFIX in t) or (LOW_INFO_TEXT in t)


def _load_gold_examples_by_tag(sd_name):
    # This library is strictly for writing style anchoring and is never used as factual evidence.
    if not GOLD_EXAMPLES_FILE.exists():
        return {}

    try:
        payload = json.loads(GOLD_EXAMPLES_FILE.read_text(encoding="utf-8", errors="ignore"))
    except Exception as exc:
        log(f"Gold examples file unreadable: {exc}", sd_name)
        return {}

    examples = payload.get("examples", {}) if isinstance(payload, dict) else {}
    if not isinstance(examples, dict):
        return {}

    out = {}
    for tag, entries in examples.items():
        if not isinstance(entries, list):
            continue
        clean_entries = []
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            if str(entry.get("status", "approved")).strip().lower() != "approved":
                continue
            text = str(entry.get("sample_text", "") or "").strip()
            if not text:
                continue
            if _is_generated_fallback_text(text):
                continue
            clean_entries.append(text)
        if clean_entries:
            out[str(tag).strip()] = clean_entries

    if out:
        total = sum(len(v) for v in out.values())
        log(f"Gold examples loaded: {total} approved sample(s) across {len(out)} tag(s)", sd_name)
    return out


def _style_profile_from_examples(example_texts):
    texts = [str(t or "").strip() for t in (example_texts or []) if str(t or "").strip()]
    if not texts:
        return {}

    bullet_docs = 0
    sentence_counts = []
    word_counts = []

    for txt in texts:
        lines = [line.strip() for line in txt.splitlines() if line.strip()]
        if lines:
            bullet_lines = sum(1 for ln in lines if ln.startswith("-") or ln.startswith("*"))
            if bullet_lines >= max(1, len(lines) // 2):
                bullet_docs += 1

        sentences = [p.strip() for p in re.split(r"(?<=[.!?])\s+", txt) if p.strip()]
        sentence_counts.append(max(1, len(sentences)))
        word_counts.append(max(1, len(re.findall(r"\b\w+\b", txt))))

    avg_sentences = sum(sentence_counts) / len(sentence_counts)
    avg_words = sum(word_counts) / len(word_counts)
    avg_words_per_sentence = avg_words / max(1.0, avg_sentences)
    prefer_bullets = bullet_docs >= max(1, len(texts) // 2)

    return {
        "prefer_bullets": prefer_bullets,
        "target_sentences": max(2, min(6, int(round(avg_sentences)))),
        "target_words_per_sentence": max(10, min(26, int(round(avg_words_per_sentence)))),
        "examples_count": len(texts),
    }


def _build_style_profiles(gold_examples_by_tag):
    profiles = {}
    for tag, examples in (gold_examples_by_tag or {}).items():
        profile = _style_profile_from_examples(examples)
        if profile:
            profiles[tag] = profile
    return profiles


def _apply_style_anchor_to_lines(lines, style_profile):
    clean_lines = [str(x or "").strip() for x in (lines or []) if str(x or "").strip()]
    if not clean_lines:
        return ""

    if not style_profile:
        return "\n".join(clean_lines)

    prefer_bullets = bool(style_profile.get("prefer_bullets", True))
    target_sentences = int(style_profile.get("target_sentences", 4) or 4)
    target_words = int(style_profile.get("target_words_per_sentence", 18) or 18)

    trimmed = clean_lines[: max(2, target_sentences)]
    out = []
    for line in trimmed:
        words = line.split()
        if len(words) > target_words:
            line = " ".join(words[:target_words]).rstrip(".,;:") + "..."
        if prefer_bullets and not line.startswith("-"):
            line = f"- {line}"
        out.append(line)

    if not prefer_bullets:
        para = " ".join(x.lstrip("-*").strip() for x in out)
        return para.strip()
    return "\n".join(out)


def _contains_number_like(text):
    t = str(text or "")
    return bool(re.search(r"\b\d+(?:[\.,]\d+)?\b", t)) or ("%" in t)


def _contains_any(text, needles):
    t = normalize(text or "")
    return any(n in t for n in (needles or []))


def _safe_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return default


def _clamp_score(value):
    return max(0, min(100, _safe_int(round(value), 0)))


def _score_coverage(tag_name, content, fill_type, sources):
    txt = str(content or "").strip()
    if not txt or LOW_INFO_TEXT in txt:
        return 5

    if fill_type == "direct_from_sd_chapter":
        base = 68
    elif fill_type == "ai_related_documents":
        base = 45
    elif fill_type == "ai_missing_chapter":
        base = 25
    else:
        base = 10

    len_score = min(1.0, len(txt) / 320.0)
    source_bonus = min(20, len(sources or []) * 6)
    return _clamp_score(base + (20 * len_score) + source_bonus)


def _score_specificity(tag_name, content):
    txt = str(content or "")
    if not txt:
        return 0

    num_hits = len(re.findall(r"\b\d+(?:[\.,]\d+)?\b", txt))
    percent_hits = txt.count("%")
    signals = TAG_SIGNALS.get(tag_name, [])
    sig_hits = sum(1 for s in signals if s in normalize(txt))

    score = 20
    score += min(35, (num_hits + percent_hits) * 6)
    score += min(35, sig_hits * 7)
    score += 10 if len(txt) >= 200 else 0
    return _clamp_score(score)


def _score_evidence_count(sources):
    sources = sources or []
    if not sources:
        return 0

    src_count = len(sources)
    table_fact_count = 0
    for s in sources:
        table_fact_count += len(s.get("table_facts", []) or [])

    score = (src_count * 20) + min(40, table_fact_count * 4)
    return _clamp_score(score)


def _score_policy_compliance(tag_name, has_exact_evidence, conflicts, fill_type):
    conflicts = conflicts or []

    if tag_name in STRICT_EVIDENCE_TAGS and not has_exact_evidence:
        return 0

    score = 100
    if fill_type in {"ai_missing_chapter", "open_too_little_info"}:
        score -= 40
    elif fill_type == "ai_related_documents":
        score -= 20

    for c in conflicts:
        sev = str(c.get("severity", "")).strip().lower()
        if sev == SEVERITY_BLOCKING:
            score -= 70
        elif sev == SEVERITY_WARNING:
            score -= 20

    return _clamp_score(score)


def _compute_quality_for_tag(tag_name, content, sources, conflicts, has_exact_evidence, fill_type):
    coverage = _score_coverage(tag_name, content, fill_type, sources)
    specificity = _score_specificity(tag_name, content)
    evidence = _score_evidence_count(sources)
    policy = _score_policy_compliance(tag_name, has_exact_evidence, conflicts, fill_type)

    overall = _clamp_score((0.30 * coverage) + (0.25 * specificity) + (0.20 * evidence) + (0.25 * policy))
    is_low = overall < QUALITY_LOW_SCORE_THRESHOLD

    return {
        "overall": overall,
        "coverage": coverage,
        "specificity": specificity,
        "evidence_count": evidence,
        "policy_compliance": policy,
        "is_low": is_low,
        "fill_type": fill_type,
    }


def _log_quality_for_tag(tag_name, quality, sd_name):
    if not quality:
        return
    low_text = "yes" if quality.get("is_low") else "no"
    log(
        f"Quality SDT '{tag_name}': overall={quality.get('overall', 0)}; "
        f"coverage={quality.get('coverage', 0)}; specificity={quality.get('specificity', 0)}; "
        f"evidence_count={quality.get('evidence_count', 0)}; policy_compliance={quality.get('policy_compliance', 0)}; "
        f"fill_type={quality.get('fill_type', 'unknown')}; low_score={low_text}",
        sd_name,
    )


def _strip_ai_generated_markers(text):
    lines = [line.strip() for line in str(text or "").splitlines() if line and line.strip()]
    cleaned = []
    for line in lines:
        lower = line.lower()
        if lower.startswith("gemiddelde kwaliteitsscore ("):
            continue
        if line == HITL_PREFIX:
            continue
        if line == HITL_VALIDATION_NOTICE:
            continue
        cleaned.append(line)
    return "\n".join(cleaned).strip()


def _chapter_hitl_questions(tag_name, fill_type, body_lines=None, target_label=None, quality_score=0):
    chapter_label = str(target_label or TAG_LABEL_NL.get(tag_name, tag_name)).strip()
    raw_lines = [str(x or "").strip() for x in (body_lines or []) if str(x or "").strip()]

    stop = {
        "voor", "met", "van", "en", "de", "het", "een", "op", "in", "tot", "door", "te", "om", "is", "zijn",
        "hoofdstuk", "inhoud", "voorlopige", "placeholder", "ai", "generated", "hitl", "bron", "bronnen",
    }
    term_candidates = []
    for line in raw_lines:
        text = re.sub(r"^[-*\u2022]\s*", "", line)
        text = re.sub(r"\(bron:[^)]+\)", "", text, flags=re.IGNORECASE)
        for token in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-_/]{2,}", text):
            lower = token.lower()
            if lower in stop:
                continue
            if lower.startswith("http"):
                continue
            term_candidates.append(token)

    focus_terms = []
    seen = set()
    for token in term_candidates:
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        focus_terms.append(token)
        if len(focus_terms) >= 3:
            break

    primary = focus_terms[0] if len(focus_terms) >= 1 else chapter_label
    secondary = focus_terms[1] if len(focus_terms) >= 2 else "operationele invulling"
    tertiary = focus_terms[2] if len(focus_terms) >= 3 else "governance"

    tag_specific = {
        "SLA_KPI": [
            "Welke SLA/KPI-metrics ontbreken nog (targetwaarde, meetmethode, meetfrequentie)?",
            "Zijn service windows, responstijden en oplostijden expliciet en consistent beschreven?",
            "Welke uitzonderingen, afhankelijkheden of uitsluitingen bij de SLA moeten nog toegevoegd worden?",
            "Zijn de KPI-definities meetbaar en contractueel ondubbelzinnig geformuleerd?",
            "Welke bewijsbron bevestigt de genoemde SLA/KPI-waarden?",
        ],
        "PRICING_ELEMENTS": [
            "Welke prijscomponenten ontbreken nog (eenmalig, recurrent, usage-based)?",
            "Zijn prijsaannames, volumes en indexeringsregels expliciet gemaakt?",
            "Welke kosten vallen buiten scope en moeten als uitsluiting worden benoemd?",
            "Zijn de prijsdrivers en verrekenlogica voldoende transparant voor sales en klant?",
            "Welke brondata is nodig om dit hoofdstuk financieel valide te maken?",
        ],
        "CLIENT_RESPONSIBILITIES": [
            "Zijn rollen en verantwoordelijkheden van klant en Cegeka expliciet gescheiden?",
            "Welke verantwoordelijkheden missen nog in onboarding, run en change context?",
            "Zijn afhankelijkheden op klantinput, tooling of approvals duidelijk benoemd?",
            "Waar kunnen interpretatieconflicten ontstaan in eigenaarschap of besluitvorming?",
            "Welke RACI-onderdelen moeten nog worden aangevuld of gecorrigeerd?",
        ],
        "CLIENT_RESPONSIBILITIES_2": [
            "Zijn rollen en verantwoordelijkheden van klant en Cegeka expliciet gescheiden?",
            "Welke verantwoordelijkheden missen nog in onboarding, run en change context?",
            "Zijn afhankelijkheden op klantinput, tooling of approvals duidelijk benoemd?",
            "Waar kunnen interpretatieconflicten ontstaan in eigenaarschap of besluitvorming?",
            "Welke RACI-onderdelen moeten nog worden aangevuld of gecorrigeerd?",
        ],
        "SCOPE": [
            "Welke in-scope onderdelen missen nog en moeten expliciet worden toegevoegd?",
            "Welke out-of-scope elementen ontbreken voor heldere afbakening?",
            "Zijn scopegrenzen meetbaar en toetsbaar geformuleerd?",
            "Welke afhankelijkheden of precondities beïnvloeden de scope?",
            "Welke concrete voorbeelden verduidelijken de scope-afbakening?",
        ],
        "REQUIREMENTS": [
            "Welke technische of organisatorische prerequisites ontbreken nog?",
            "Zijn afhankelijkheden op klantsystemen, toegang en data expliciet beschreven?",
            "Welke minimale readiness-criteria moeten nog worden vastgelegd?",
            "Zijn verplichtingen per partij helder toegewezen?",
            "Welke requirement-bronnen moeten nog gevalideerd worden?",
        ],
        "TRANSITION_TRANSFORMATION": [
            "Welke transitieactiviteiten ontbreken nog in planning en uitvoering?",
            "Zijn milestones, afhankelijkheden en kritieke paden expliciet uitgewerkt?",
            "Welke risico's en mitigerende acties moeten nog benoemd worden?",
            "Zijn handovercriteria en acceptatiecriteria volledig en toetsbaar?",
            "Welke rollen moeten per transitiefase nog concreet worden toegewezen?",
        ],
        "ARCHITECTURAL_DESCRIPTION": [
            "Welke architectuurcomponenten of integraties ontbreken nog?",
            "Zijn security-, availability- en schaalbaarheidskeuzes expliciet onderbouwd?",
            "Welke technische aannames of constraints moeten toegevoegd worden?",
            "Zijn interfaces, datastromen en verantwoordelijkheden voldoende concreet?",
            "Welke referentie-architectuur of bron bevestigt deze beschrijving?",
        ],
        "TERMS_CONDITIONS": [
            "Welke contractuele voorwaarden of uitzonderingen ontbreken nog?",
            "Zijn aansprakelijkheid, beperkingen en afhankelijkheden expliciet benoemd?",
            "Welke compliance- of governancevoorwaarden moeten nog toegevoegd worden?",
            "Zijn termen consistent met de rest van het document en contracttemplates?",
            "Welke juridische reviewpunten moeten nog door HITL worden afgestemd?",
        ],
    }

    transition_questions = [
        "Welke transitieactiviteiten ontbreken nog in planning en uitvoering?",
        "Zijn milestones, afhankelijkheden en kritieke paden expliciet uitgewerkt?",
        "Welke risico's en mitigerende acties moeten nog benoemd worden?",
        "Zijn handovercriteria en acceptatiecriteria volledig en toetsbaar?",
        "Welke rollen moeten per transitiefase nog concreet worden toegewezen?",
    ]
    pricing_questions = [
        "Welke prijscomponenten ontbreken nog (eenmalig, recurrent, usage-based)?",
        "Zijn prijsaannames, volumes en indexeringsregels expliciet gemaakt?",
        "Welke kosten vallen buiten scope en moeten als uitsluiting worden benoemd?",
        "Zijn de prijsdrivers en verrekenlogica voldoende transparant voor sales en klant?",
        "Welke brondata is nodig om dit hoofdstuk financieel valide te maken?",
    ]
    responsibilities_questions = [
        "Zijn rollen en verantwoordelijkheden van klant en Cegeka expliciet gescheiden?",
        "Welke verantwoordelijkheden missen nog in onboarding, run en change context?",
        "Zijn afhankelijkheden op klantinput, tooling of approvals duidelijk benoemd?",
        "Waar kunnen interpretatieconflicten ontstaan in eigenaarschap of besluitvorming?",
        "Welke RACI-onderdelen moeten nog worden aangevuld of gecorrigeerd?",
    ]
    scope_questions = [
        "Welke in-scope onderdelen missen nog en moeten expliciet worden toegevoegd?",
        "Welke out-of-scope elementen ontbreken voor heldere afbakening?",
        "Zijn scopegrenzen meetbaar en toetsbaar geformuleerd?",
        "Welke afhankelijkheden of precondities beïnvloeden de scope?",
        "Welke concrete voorbeelden verduidelijken de scope-afbakening?",
    ]

    if tag_name in tag_specific:
        questions = list(tag_specific[tag_name])
    elif tag_name.startswith("TRANSITION_"):
        questions = list(transition_questions)
    elif tag_name.startswith("COST_") or tag_name in {"CHARGING_MECHANISM", "COMMERCIAL_SHEET", "PRICING_ELEMENTS"}:
        questions = list(pricing_questions)
    elif tag_name.startswith("ROLES_") or "RESPONSIBILITIES" in tag_name:
        questions = list(responsibilities_questions)
    elif "SCOPE" in tag_name:
        questions = list(scope_questions)
    else:
        questions = [
            f"Voor '{chapter_label}': welke concrete input ontbreekt nog rond {primary} om dit hoofdstuk toetsbaar te maken?",
            f"Welke expliciete bronreferenties ontbreken voor de statements over {primary} en {secondary}?",
            f"Welke meetbare criteria of waarden moeten toegevoegd worden voor '{chapter_label}'?",
            f"Welke risico's, aannames of afhankelijkheden rond {tertiary} moeten expliciet worden gemaakt?",
            f"Welke 2-3 productspecifieke voorbeelden ontbreken nog om '{chapter_label}' kwalitatief sterk te finaliseren?",
        ]

    # Force chapter specificity even for predefined packs by anchoring the chapter label.
    chapter_specific_questions = []
    for q in questions[:5]:
        if chapter_label.lower() in q.lower():
            chapter_specific_questions.append(q)
        else:
            chapter_specific_questions.append(f"Voor '{chapter_label}': {q[0].lower() + q[1:] if len(q) > 1 else q.lower()}")
    questions = chapter_specific_questions

    if fill_type == "open_too_little_info":
        questions[0] = (
            f"Voor '{chapter_label}': welke ontbrekende broninformatie moet eerst worden aangeleverd om dit hoofdstuk inhoudelijk te kunnen invullen?"
        )

    score = _safe_int(quality_score, 0)
    if score < 35:
        score_questions = [
            f"Voor '{chapter_label}' (score {score}/100): welke 3 ontbrekende feiten leveren de snelste kwaliteitswinst op?",
            f"Voor '{chapter_label}' (score {score}/100): welke bron (owner + document) moet HITL eerst valideren om hallucinatie-risico te verlagen?",
        ]
    elif score < 70:
        score_questions = [
            f"Voor '{chapter_label}' (score {score}/100): welke alinea's moeten inhoudelijk worden aangescherpt om van concept naar klantklare tekst te gaan?",
            f"Voor '{chapter_label}' (score {score}/100): welke 2 meetbare waarden ontbreken nog om claims hard te maken?",
        ]
    else:
        score_questions = [
            f"Voor '{chapter_label}' (score {score}/100): welke formuleringen moeten juridisch/commercieel worden geharmoniseerd met de offerteksten?",
            f"Voor '{chapter_label}' (score {score}/100): welke optimalisaties verhogen nog de klantduidelijkheid zonder scope uit te breiden?",
        ]

    questions = (questions[:3] + score_questions)[:5]
    return questions


def _decorate_ai_generated_text_with_quality(tag_name, generated_text, quality, fill_type=None):
    if fill_type not in AI_DECORATED_FILL_TYPES:
        return _strip_ai_generated_markers(generated_text)

    score = 0
    if isinstance(quality, dict):
        score = _safe_int(quality.get("overall", 0), 0)

    score_line = f"Gemiddelde kwaliteitsscore ({tag_name}): {score}/100"
    raw_text = _strip_ai_generated_markers(generated_text)
    lines = [line.strip() for line in str(raw_text or "").splitlines() if line and line.strip()]

    if not lines:
        lines = [HITL_PREFIX, LOW_INFO_TEXT]

    if not any(HITL_PREFIX in line for line in lines):
        lines.insert(0, HITL_PREFIX)

    has_explicit_hitl_notice = any(
        ("hitl" in line.lower()) and ("valide" in line.lower() or "verifier" in line.lower() or "validate" in line.lower())
        for line in lines
    )
    if not has_explicit_hitl_notice:
        insert_idx = 1 if lines and HITL_PREFIX in lines[0] else 0
        lines.insert(insert_idx, HITL_VALIDATION_NOTICE)

    body_lines = [
        line for line in lines
        if line not in {HITL_PREFIX, HITL_VALIDATION_NOTICE}
        and not line.lower().startswith("gemiddelde kwaliteitsscore (")
    ]

    fill_meta = {
        "ai_related_documents": {
            "bronstatus": "AI-voorinvulling op basis van gerelateerde documenten",
            "inhoud_label": "Inhoud uit gerelateerde bronnen:",
        },
        "ai_missing_chapter": {
            "bronstatus": "AI-voorinvulling wegens ontbrekend hoofdstuk in de bron-SD",
            "inhoud_label": "Voorlopige hoofdstukinhoud:",
        },
        "open_too_little_info": {
            "bronstatus": "Open placeholder wegens onvoldoende broninformatie",
            "inhoud_label": "Ontbrekende informatie / placeholder:",
        },
    }
    meta = fill_meta.get(fill_type, {
        "bronstatus": "AI-voorinvulling",
        "inhoud_label": "Inhoud:",
    })

    structured_body = [
        f"Hoofdstuk: {TAG_LABEL_NL.get(tag_name, tag_name)}",
        f"Bronstatus: {meta['bronstatus']}",
        meta["inhoud_label"],
    ]

    if not body_lines:
        structured_body.append(f"- {LOW_INFO_TEXT}")
    else:
        for line in body_lines:
            clean = line.strip()
            if not clean:
                continue
            if clean.endswith(":") and len(clean.split()) <= 10:
                structured_body.append(clean)
            elif clean.startswith("- "):
                structured_body.append(clean)
            else:
                structured_body.append(f"- {clean}")

    tag_questions = _chapter_hitl_questions(
        tag_name=tag_name,
        fill_type=fill_type,
        body_lines=body_lines,
        target_label=TAG_LABEL_NL.get(tag_name, tag_name),
        quality_score=score,
    )

    structured_body.extend([
        "HITL-actie:",
        "- Verifieer de inhoud tegen de SD-bron en/of gevalideerde input.",
        "- Werk ontbrekende details bij en finaliseer de tekst.",
        "HITL-vragen voor kwaliteitsverbetering:",
    ])
    structured_body.extend([f"- {question}" for question in tag_questions])

    return "\n".join([score_line, HITL_PREFIX, HITL_VALIDATION_NOTICE, *structured_body]).strip()


def _parse_table_facts_from_section(section_node):
    raw = section_node.findtext("TableFactsJson", default="").strip()
    if not raw:
        return []

    try:
        parsed = json.loads(raw)
    except Exception:
        return []

    if not isinstance(parsed, list):
        return []

    normalized_facts = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        pairs = item.get("facts", {})
        if not isinstance(pairs, dict):
            pairs = {}

        clean_pairs = {}
        for k, v in pairs.items():
            kk = str(k or "").strip()
            vv = str(v or "").strip()
            if kk and vv:
                clean_pairs[kk] = vv

        row_text = str(item.get("row_text", "") or "").strip()
        fact_type = str(item.get("fact_type", "general") or "general").strip().lower()
        if clean_pairs or row_text:
            normalized_facts.append(
                {
                    "fact_type": fact_type,
                    "row_text": row_text,
                    "facts": clean_pairs,
                }
            )

    return normalized_facts


def _table_facts_to_text(table_facts, max_items=6):
    lines = []
    for fact in table_facts[:max_items]:
        if not isinstance(fact, dict):
            continue
        pairs = fact.get("facts", {}) or {}
        pair_text = "; ".join(f"{k}={v}" for k, v in pairs.items() if str(k).strip() and str(v).strip())
        row_text = str(fact.get("row_text", "") or "").strip()
        fact_type = str(fact.get("fact_type", "general") or "general").strip()

        content = pair_text or row_text
        if not content:
            continue
        if fact_type:
            lines.append(f"[{fact_type}] {content}")
        else:
            lines.append(content)

    return "\n".join(lines)


def _has_exact_evidence_for_tag(tag_name, sources):
    if tag_name not in STRICT_EVIDENCE_TAGS:
        return True, "not a strict-evidence chapter"
    if not sources:
        return False, "no selected source evidence"

    full_parts = []
    for s in sources:
        full_parts.append(f"{s.get('section_name', '')}\n{s.get('content', '')}")
        table_text = _table_facts_to_text(s.get("table_facts", []), max_items=8)
        if table_text:
            full_parts.append(table_text)
    full = "\n".join(full_parts)

    if tag_name == "SLA_KPI":
        has_sla_term = _contains_any(full, ["sla", "service level", "availability", "response time", "resolution time"])
        has_kpi_term = _contains_any(full, ["kpi", "metric", "metrics", "target", "threshold"])
        has_metric_value = _contains_number_like(full) or _contains_any(full, ["uptime", "hours", "minutes", "seconds", "percent"])
        if has_sla_term and has_kpi_term and has_metric_value:
            return True, "contains explicit SLA and KPI terms with measurable indicators"
        return False, "missing explicit SLA+KPI metric evidence"

    if tag_name == "PRICING_ELEMENTS":
        has_pricing_term = _contains_any(full, ["pricing", "billing", "cost", "price", "charge", "sku", "invoice", "recurring", "monthly", "one-time"])
        has_pricing_value = _contains_number_like(full) or _contains_any(full, ["per month", "per user", "per app", "eur", "usd"])
        if has_pricing_term and has_pricing_value:
            return True, "contains pricing/billing terms with billable indicators"
        return False, "missing explicit pricing evidence"

    if tag_name == "CLIENT_RESPONSIBILITIES":
        has_resp_term = _contains_any(full, ["responsibility", "responsibilities", "customer", "client", "provided by customer", "raci", "accountable", "responsible"])
        if has_resp_term:
            return True, "contains explicit responsibilities language"
        return False, "missing explicit customer/client responsibilities evidence"

    if tag_name == "TERMS_CONDITIONS":
        has_legal_term = _contains_any(full, ["terms", "conditions", "contract", "liability", "compliance", "obligation", "limitation", "governing law", "agreement"])
        if has_legal_term:
            return True, "contains explicit legal/contract terms"
        return False, "missing explicit legal terms evidence"

    return False, "strict-evidence tag without matching validator"


def _tokenize_semantic(text):
    tokens = re.findall(r"[a-z0-9]+", normalize(text or ""))
    return [t for t in tokens if len(t) >= 3]


def _tag_mapping_keys(tag_name):
    out = []
    for key, value in MAPPING.items():
        if value == tag_name:
            out.append(key)
    return out


def _semantic_score_section_for_tag(tag_name, section_name, section_category, content, table_facts=None):
    signals = TAG_SIGNALS.get(tag_name, [])
    table_facts = table_facts or []
    table_text = _table_facts_to_text(table_facts, max_items=8)
    if not content:
        if not table_text:
            return None

    heading_norm = normalize(section_name)
    combined = normalize(f"{section_name} {content[:2800]} {table_text[:1800]}")
    signal_hits = [sig for sig in signals if sig in combined]
    table_signal_hits = [sig for sig in signals if sig in normalize(table_text)]

    expected_fact_types = TAG_TABLE_FACT_TYPES.get(tag_name, set())
    fact_type_hits = sum(
        1 for f in table_facts if str(f.get("fact_type", "")).strip().lower() in expected_fact_types
    )

    profile_text = " ".join(
        [
            TAG_LABEL_NL.get(tag_name, tag_name),
            " ".join(signals),
            " ".join(_tag_mapping_keys(tag_name)),
        ]
    )
    profile_tokens = set(_tokenize_semantic(profile_text))
    section_tokens = set(_tokenize_semantic(f"{section_name} {content[:1600]}"))
    common = len(profile_tokens & section_tokens)
    union = len(profile_tokens | section_tokens)
    jaccard = (common / union) if union else 0.0

    mapping_ratio = 0.0
    for key in _tag_mapping_keys(tag_name):
        mapping_ratio = max(mapping_ratio, difflib.SequenceMatcher(None, heading_norm, normalize(key)).ratio())

    mapped_tag, match_source = resolve_tag_for_section(section_name, section_category)
    mapping_boost = 3.0 if mapped_tag == tag_name else 0.0
    match_bonus = 0.8 if (mapped_tag == tag_name and match_source == "category") else 0.0
    table_bonus = (1.2 * len(table_signal_hits)) + (1.5 * fact_type_hits)

    score = (1.6 * len(signal_hits)) + (4.0 * jaccard) + (3.0 * mapping_ratio) + mapping_boost + match_bonus + table_bonus

    if score < 2.2 and len(signal_hits) == 0 and len(table_signal_hits) == 0 and jaccard < 0.08 and mapping_ratio < 0.60:
        return None

    reason_parts = []
    if mapped_tag == tag_name:
        reason_parts.append(f"{match_source} match")
    if signal_hits:
        reason_parts.append("signals: " + ", ".join(signal_hits[:2]))
    if jaccard >= 0.10:
        reason_parts.append(f"semantic overlap {jaccard:.2f}")
    if mapping_ratio >= 0.60:
        reason_parts.append(f"header similarity {mapping_ratio:.2f}")
    if table_signal_hits:
        reason_parts.append("table signals: " + ", ".join(table_signal_hits[:2]))
    if fact_type_hits > 0:
        reason_parts.append(f"table fact type match x{fact_type_hits}")

    return {
        "score": score,
        "section_name": section_name,
        "content": content,
        "table_facts": table_facts,
        "reason": "; ".join(reason_parts) if reason_parts else "best semantic match",
    }


def _select_best_sources_per_tag(xml_sections):
    selected = {}

    for tag in TEMPLATE_TAG_ORDER:
        candidates = []
        seen_sections = set()

        for section_node in xml_sections:
            section_name = section_node.get("name", "").strip()
            section_category = section_node.findtext("Category", default="")
            content = section_node.findtext("Content", default="").strip()

            if not section_name or not content:
                continue
            if looks_irrelevant_section(section_name):
                continue
            if _looks_toc_noise_content(content):
                continue

            content = _strip_toc_noise(content)
            if len(content) < 30:
                continue

            table_facts = _parse_table_facts_from_section(section_node)
            scored = _semantic_score_section_for_tag(tag, section_name, section_category, content, table_facts)
            if not scored:
                continue

            sec_key = section_name.lower()
            if sec_key in seen_sections:
                continue
            seen_sections.add(sec_key)
            candidates.append(scored)

        if not candidates:
            continue

        candidates.sort(key=lambda x: (-x["score"], -len(x["content"])))
        selected[tag] = candidates[:MAX_SOURCE_SECTIONS_PER_TAG]

    return selected


def _merge_selected_source_content(tag_name, sources):
    if not sources:
        return ""

    table_evidence = []
    for src in sources:
        table_facts = src.get("table_facts", []) or []
        for fact in table_facts:
            if not isinstance(fact, dict):
                continue
            fact_type = str(fact.get("fact_type", "general") or "general").strip().lower()
            expected_fact_types = TAG_TABLE_FACT_TYPES.get(tag_name, set())
            if expected_fact_types and fact_type not in expected_fact_types:
                continue
            pairs = fact.get("facts", {}) or {}
            pair_text = "; ".join(f"{k}={v}" for k, v in pairs.items() if str(k).strip() and str(v).strip())
            row_text = str(fact.get("row_text", "") or "").strip()
            rendered = pair_text or row_text
            if rendered:
                table_evidence.append(f"- {rendered} (bron: {src.get('section_name', '')})")

    if table_evidence:
        lines = [
            f"Genormaliseerde tabel-feiten voor {TAG_LABEL_NL.get(tag_name, tag_name)}:",
            *table_evidence[:10],
        ]
        return "\n".join(lines).strip()

    cleaned_blocks = []
    for src in sources:
        cleaned = _strip_toc_noise(src.get("content", ""))
        if not cleaned:
            continue
        if _looks_toc_noise_content(cleaned):
            continue
        cleaned_blocks.append(cleaned[:MAX_SOURCE_CONTENT_CHARS].strip())

    if not cleaned_blocks:
        return ""

    if len(cleaned_blocks) == 1:
        return cleaned_blocks[0]

    # Keep final chapter text clean; traceability remains available in logs/XLSX.
    return "\n\n".join(cleaned_blocks).strip()


def _validate_source_relevance_for_tag(tag_name, source_item):
    section_name = source_item.get("section_name", "")
    content = source_item.get("content", "")
    combined = normalize(f"{section_name} {content[:2200]}")

    must_signals = INTENT_MUST_HAVE_SIGNALS.get(tag_name, [])
    hits = [sig for sig in must_signals if sig in combined]
    if hits:
        return True, f"intent ok via: {', '.join(hits[:2])}", SEVERITY_INFO

    map_ratio = 0.0
    for key in _tag_mapping_keys(tag_name):
        map_ratio = max(map_ratio, difflib.SequenceMatcher(None, normalize(section_name), normalize(key)).ratio())

    if map_ratio >= 0.82:
        return True, f"intent ok via strong header similarity {map_ratio:.2f}", SEVERITY_INFO

    return False, "intent mismatch: insufficient chapter-intent signals", SEVERITY_WARNING


def _validate_selected_sources(selected_sources_by_tag):
    validated = {}
    rejected = {}

    for tag, sources in selected_sources_by_tag.items():
        ok_items = []
        bad_items = []

        for item in sources:
            is_ok, reason, severity = _validate_source_relevance_for_tag(tag, item)
            enriched = dict(item)
            enriched["intent_reason"] = reason
            enriched["intent_severity"] = severity
            if is_ok:
                ok_items.append(enriched)
            else:
                bad_items.append(enriched)

        if ok_items:
            validated[tag] = ok_items
        if bad_items:
            rejected[tag] = bad_items

    return validated, rejected


def _find_conflicts(validated_sources_by_tag):
    conflicts = {tag: [] for tag in validated_sources_by_tag.keys()}

    # Conflict 1: scope statements that mix inclusive and exclusive language in selected sources.
    scope_sources = validated_sources_by_tag.get("SCOPE", [])
    if scope_sources:
        scope_text = " ".join((s.get("content", "") + " " + s.get("section_name", "")) for s in scope_sources).lower()
        has_in = any(k in scope_text for k in ["in scope", "included", "part of service", "covered"])
        has_out = any(k in scope_text for k in ["out of scope", "excluded", "not included", "outside scope"])
        if has_in and has_out:
            conflicts["SCOPE"].append({
                "severity": SEVERITY_WARNING,
                "code": "scope_mixed_statements",
                "message": "potential conflict: mixed in-scope and out-of-scope statements",
            })

    # Conflict 2: SLA promises vs support window constraints.
    sla_sources = validated_sources_by_tag.get("SLA_KPI", [])
    support_sources = validated_sources_by_tag.get("OPERATIONAL_SUPPORT", [])
    if sla_sources and support_sources:
        sla_text = " ".join(s.get("content", "") for s in sla_sources).lower()
        support_text = " ".join(s.get("content", "") for s in support_sources).lower()

        sla_247 = any(k in sla_text for k in ["24/7", "24x7", "24 x 7", "always available"])
        support_limited = any(
            k in support_text
            for k in [
                "business hours",
                "office hours",
                "service window",
                "working days",
                "08:00",
                "09:00",
                "17:00",
                "18:00",
            ]
        )

        if sla_247 and support_limited:
            msg = {
                "severity": SEVERITY_BLOCKING,
                "code": "sla_support_window_mismatch",
                "message": "potential conflict: SLA suggests 24/7 while support window appears limited",
            }
            conflicts.setdefault("SLA_KPI", []).append(msg)
            conflicts.setdefault("OPERATIONAL_SUPPORT", []).append(msg)

    return conflicts


def _split_sentences(text):
    cleaned = " ".join(str(text or "").split())
    if not cleaned:
        return []
    parts = re.split(r"(?<=[.!?])\s+", cleaned)
    return [p.strip() for p in parts if p and len(p.strip()) >= 40]


def _section_evidence_for_tag(tag_name, xml_sections, max_items=3):
    signals = TAG_SIGNALS.get(tag_name, [])
    if not xml_sections or not signals:
        return []

    scored = []
    seen = set()

    for section_node in xml_sections:
        section_name = section_node.get("name", "").strip()
        section_category = section_node.findtext("Category", default="")
        content = section_node.findtext("Content", default="").strip()
        if not section_name or not content:
            continue
        if looks_irrelevant_section(section_name):
            continue
        if _looks_toc_noise_content(content):
            continue

        content = _strip_toc_noise(content)
        if len(content) < 30:
            continue

        combined = normalize(f"{section_name} {content[:2400]}")
        signal_hits = sum(1 for sig in signals if sig in combined)
        if signal_hits == 0:
            continue

        mapped_tag, _ = resolve_tag_for_section(section_name, section_category)
        heading_norm = normalize(section_name)
        heading_hits = sum(1 for sig in signals if sig in heading_norm)

        if mapped_tag != tag_name and heading_hits == 0 and signal_hits < 2:
            continue

        base_score = signal_hits + (4 if mapped_tag == tag_name else 0)

        for sentence in _split_sentences(content)[:12]:
            snorm = normalize(sentence)
            if len(sentence) > 320:
                continue
            if len(sentence.split()) > 50:
                continue
            if sentence.count("|") >= 2:
                continue
            sentence_hits = sum(1 for sig in signals if sig in snorm)
            if sentence_hits == 0:
                continue
            key = (section_name.lower(), sentence.lower())
            if key in seen:
                continue
            seen.add(key)
            score = base_score + (2 * sentence_hits)
            scored.append((score, section_name, sentence))

    scored.sort(key=lambda x: (-x[0], len(x[2])))
    return scored[:max_items]


def _build_generated_text(tag_name, collected_content, xml_sections=None, style_profile=None):
    if tag_name in STRICT_EVIDENCE_TAGS:
        return _build_low_info_text()

    if tag_name in LOW_CONFIDENCE_TAGS:
        return _build_low_info_text()

    target_label = TAG_LABEL_NL.get(tag_name, tag_name)
    section_evidence = _section_evidence_for_tag(tag_name, xml_sections or [])
    if section_evidence:
        style_lines = [f"{_clip(sentence, 280)} (bron: {section_name})" for _, section_name, sentence in section_evidence]
        anchored_block = _apply_style_anchor_to_lines(style_lines, style_profile)
        return "\n".join([
            HITL_PREFIX,
            f"Voorlopige invulling voor {target_label} op basis van SD-broninhoud:",
            anchored_block,
            "Gelieve dit hoofdstuk inhoudelijk te verifieren en te finaliseren via HITL.",
        ])

    related = RELATED_TAGS.get(tag_name, [])
    evidence_lines = []

    for rel_tag in related:
        rel_text = collected_content.get(rel_tag, "").strip()
        if not rel_text:
            continue
        if _is_generated_fallback_text(rel_text):
            continue
        evidence_lines.append(f"- Afgeleid uit {TAG_LABEL_NL.get(rel_tag, rel_tag)}: {_clip(rel_text)}")
        if len(evidence_lines) >= 2:
            break

    if not evidence_lines:
        return _build_low_info_text()

    anchored_block = _apply_style_anchor_to_lines(evidence_lines, style_profile)
    return "\n".join([
        HITL_PREFIX,
        f"Voor {target_label} is onderstaande voorlopige invulling afgeleid uit beschikbare SD-informatie:",
        anchored_block,
        "Gelieve dit hoofdstuk inhoudelijk te verifieren en te finaliseren via HITL.",
    ])


def _read_docx_text(path):
    try:
        with ZipFile(path, "r") as z:
            xml_bytes = z.read("word/document.xml")
        root = etree.fromstring(xml_bytes)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        texts = root.xpath(".//w:t/text()", namespaces=ns)
        return " ".join(t.strip() for t in texts if t and t.strip())
    except Exception:
        return ""


def _extract_docx_tables(path, max_tables=4, max_rows=4, max_cells=6):
    snippets = []
    try:
        with ZipFile(path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        tbl_nodes = root.xpath(".//*[local-name()='tbl']")
        for tbl in tbl_nodes[:max_tables]:
            row_parts = []
            rows = tbl.xpath(".//*[local-name()='tr']")
            for tr in rows[:max_rows]:
                cells = []
                tc_nodes = tr.xpath(".//*[local-name()='tc']")
                for tc in tc_nodes[:max_cells]:
                    cell_text = " ".join(
                        t.strip()
                        for t in tc.xpath(".//*[local-name()='t']/text()")
                        if t and t.strip()
                    )
                    if cell_text:
                        cells.append(cell_text)
                if cells:
                    row_parts.append(" | ".join(cells))

            if row_parts:
                snippets.append(" ; ".join(row_parts))
    except Exception:
        return []
    return snippets


def _extract_docx_image_snippets(path, max_images=5):
    snippets = []
    try:
        with ZipFile(path, "r") as z:
            root = etree.fromstring(z.read("word/document.xml"))

        # Prefer alt/caption-style metadata when available.
        docpr_nodes = root.xpath(".//*[local-name()='docPr']")
        for node in docpr_nodes[:max_images]:
            title = (node.get("title") or "").strip()
            descr = (node.get("descr") or "").strip()
            name = (node.get("name") or "").strip()
            meta = descr or title or name
            if meta:
                snippets.append(meta)

        if not snippets:
            drawing_count = len(root.xpath(".//*[local-name()='drawing']"))
            if drawing_count > 0:
                snippets.append(f"{drawing_count} afbeelding(en) aanwezig")
    except Exception:
        return []
    return snippets


def _read_pptx_text(path):
    chunks = []
    try:
        with ZipFile(path, "r") as z:
            slide_names = [n for n in z.namelist() if n.startswith("ppt/slides/") and n.endswith(".xml")]
            for name in slide_names:
                try:
                    root = etree.fromstring(z.read(name))
                except Exception:
                    continue
                texts = root.xpath(".//*[local-name()='t']/text()")
                if texts:
                    chunks.append(" ".join(t.strip() for t in texts if t and t.strip()))
    except Exception:
        return ""
    return "\n".join(chunks)


def _extract_pptx_tables(path, max_tables=4, max_rows=4, max_cells=6):
    snippets = []
    try:
        with ZipFile(path, "r") as z:
            slide_names = [n for n in z.namelist() if n.startswith("ppt/slides/") and n.endswith(".xml")]
            for name in slide_names:
                try:
                    root = etree.fromstring(z.read(name))
                except Exception:
                    continue

                tbl_nodes = root.xpath(".//*[local-name()='tbl']")
                for tbl in tbl_nodes[:max_tables]:
                    row_parts = []
                    rows = tbl.xpath(".//*[local-name()='tr']")
                    for tr in rows[:max_rows]:
                        cells = []
                        tc_nodes = tr.xpath(".//*[local-name()='tc']")
                        for tc in tc_nodes[:max_cells]:
                            cell_text = " ".join(
                                t.strip()
                                for t in tc.xpath(".//*[local-name()='t']/text()")
                                if t and t.strip()
                            )
                            if cell_text:
                                cells.append(cell_text)
                        if cells:
                            row_parts.append(" | ".join(cells))
                    if row_parts:
                        snippets.append(" ; ".join(row_parts))
                    if len(snippets) >= max_tables:
                        return snippets
    except Exception:
        return []
    return snippets


def _extract_pptx_image_snippets(path, max_images=5):
    snippets = []
    try:
        with ZipFile(path, "r") as z:
            slide_names = [n for n in z.namelist() if n.startswith("ppt/slides/") and n.endswith(".xml")]
            total_pics = 0
            for name in slide_names:
                try:
                    root = etree.fromstring(z.read(name))
                except Exception:
                    continue

                c_nv_pr_nodes = root.xpath(".//*[local-name()='pic']//*[local-name()='cNvPr']")
                for node in c_nv_pr_nodes:
                    pic_name = (node.get("name") or "").strip()
                    descr = (node.get("descr") or "").strip()
                    meta = descr or pic_name
                    if meta:
                        snippets.append(meta)
                        if len(snippets) >= max_images:
                            return snippets

                total_pics += len(root.xpath(".//*[local-name()='pic']"))

            if not snippets and total_pics > 0:
                snippets.append(f"{total_pics} afbeelding(en) aanwezig")
    except Exception:
        return []
    return snippets


def _read_xlsx_text(path):
    chunks = []
    try:
        with ZipFile(path, "r") as z:
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                try:
                    shared_root = etree.fromstring(z.read("xl/sharedStrings.xml"))
                    shared = shared_root.xpath(".//*[local-name()='t']/text()")
                except Exception:
                    shared = []

            sheet_names = [n for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
            for name in sheet_names:
                try:
                    root = etree.fromstring(z.read(name))
                except Exception:
                    continue

                # Values from inline strings and plain cell values.
                inline = root.xpath(".//*[local-name()='is']//*[local-name()='t']/text()")
                direct_vals = root.xpath(".//*[local-name()='c']/*[local-name()='v']/text()")
                resolved = []
                for v in direct_vals:
                    vv = (v or "").strip()
                    if vv.isdigit():
                        idx = int(vv)
                        if 0 <= idx < len(shared):
                            resolved.append(shared[idx])
                            continue
                    resolved.append(vv)

                merged = [x.strip() for x in (inline + resolved) if x and str(x).strip()]
                if merged:
                    chunks.append(" ".join(merged))
    except Exception:
        return ""
    return "\n".join(chunks)


def _extract_xlsx_table_snippets(path, max_sheets=3, max_rows=5, max_cells=6):
    snippets = []
    try:
        with ZipFile(path, "r") as z:
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                try:
                    shared_root = etree.fromstring(z.read("xl/sharedStrings.xml"))
                    shared = [
                        " ".join(
                            t.strip() for t in si.xpath(".//*[local-name()='t']/text()") if t and t.strip()
                        )
                        for si in shared_root.xpath(".//*[local-name()='si']")
                    ]
                except Exception:
                    shared = []

            sheet_names = [n for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
            for name in sheet_names[:max_sheets]:
                try:
                    root = etree.fromstring(z.read(name))
                except Exception:
                    continue

                row_parts = []
                rows = root.xpath(".//*[local-name()='row']")
                for row in rows[:max_rows]:
                    cells = []
                    for c in row.xpath("./*[local-name()='c']")[:max_cells]:
                        ctype = (c.get("t") or "").strip()
                        v = c.xpath("./*[local-name()='v']/text()")
                        inline = c.xpath(".//*[local-name()='is']//*[local-name()='t']/text()")
                        val = ""
                        if inline:
                            val = " ".join(x.strip() for x in inline if x and x.strip())
                        elif v:
                            vv = (v[0] or "").strip()
                            if ctype == "s" and vv.isdigit():
                                idx = int(vv)
                                if 0 <= idx < len(shared):
                                    val = shared[idx]
                                else:
                                    val = vv
                            else:
                                val = vv
                        if val:
                            cells.append(val)
                    if cells:
                        row_parts.append(" | ".join(cells))

                if row_parts:
                    snippets.append(" ; ".join(row_parts))
    except Exception:
        return []
    return snippets


def _read_plain_text(path):
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""


def _read_csv_text(path):
    rows = []
    try:
        with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                line = " ".join((cell or "").strip() for cell in row if (cell or "").strip())
                if line:
                    rows.append(line)
    except Exception:
        return ""
    return "\n".join(rows)


def _extract_csv_table_snippets(path, max_rows=6):
    rows = []
    try:
        with open(path, "r", encoding="utf-8", errors="ignore", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                line = " | ".join((cell or "").strip() for cell in row if (cell or "").strip())
                if line:
                    rows.append(line)
                if len(rows) >= max_rows:
                    break
    except Exception:
        return []
    return [" ; ".join(rows)] if rows else []


def _read_related_file_text(path):
    ext = path.suffix.lower()
    if ext == ".docx":
        return _read_docx_text(path)
    if ext in {".txt", ".md", ".html", ".htm"}:
        return _read_plain_text(path)
    if ext == ".csv":
        return _read_csv_text(path)
    if ext == ".pptx":
        return _read_pptx_text(path)
    if ext == ".xlsx":
        return _read_xlsx_text(path)
    return ""


def _extract_related_artifacts(path):
    ext = path.suffix.lower()
    tables = []
    images = []

    if ext == ".docx":
        tables = _extract_docx_tables(path)
        images = _extract_docx_image_snippets(path)
    elif ext == ".pptx":
        tables = _extract_pptx_tables(path)
        images = _extract_pptx_image_snippets(path)
    elif ext == ".xlsx":
        tables = _extract_xlsx_table_snippets(path)
    elif ext == ".csv":
        tables = _extract_csv_table_snippets(path)

    return {
        "tables": [t for t in tables if t and t.strip()],
        "images": [i for i in images if i and i.strip()],
    }


def _safe_name(value):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "")).strip("_") or "item"


def _short_id(value, length=12):
    raw = str(value or "").encode("utf-8", errors="ignore")
    return hashlib.md5(raw).hexdigest()[:length]


def _extract_media_files(path, dest_root):
    ext = path.suffix.lower()
    prefixes = []
    if ext == ".docx":
        prefixes = ["word/media/"]
    elif ext == ".pptx":
        prefixes = ["ppt/media/"]
    elif ext == ".xlsx":
        prefixes = ["xl/media/"]
    else:
        return []

    dest_root.mkdir(parents=True, exist_ok=True)
    out = []

    try:
        with ZipFile(path, "r") as z:
            members = [
                n
                for n in z.namelist()
                if any(n.startswith(pref) for pref in prefixes)
                and Path(n).suffix.lower() in IMAGE_EXTENSIONS
            ]

            for idx, member in enumerate(members[:MAX_IMAGES_PER_FILE], start=1):
                suffix = Path(member).suffix.lower()
                out_name = f"img_{idx}_{_short_id(member, 8)}{suffix}"
                out_path = dest_root / out_name
                try:
                    out_path.write_bytes(z.read(member))
                except Exception:
                    continue
                out.append(out_path)
    except Exception:
        return []

    return out


def _read_img_dimensions_emu(path, max_width_emu=4114800):
    """Read image (w,h) from file header; return (cx,cy) in EMUs scaled to max_width_emu."""
    MAX_W = max_width_emu
    try:
        raw = Path(path).read_bytes()
        w_px, h_px = 0, 0

        if len(raw) >= 24 and raw[:8] == b"\x89PNG\r\n\x1a\n":
            w_px, h_px = struct.unpack(">II", raw[16:24])
        elif len(raw) >= 4 and raw[:2] == b"\xff\xd8":
            i = 2
            while i < len(raw) - 8:
                if raw[i] != 0xFF:
                    break
                marker = raw[i + 1]
                if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                    i += 2
                    continue
                if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                    h_px, w_px = struct.unpack(">HH", raw[i + 5 : i + 9])
                    break
                if i + 4 > len(raw):
                    break
                seg_len = struct.unpack(">H", raw[i + 2 : i + 4])[0]
                i += 2 + seg_len
        elif len(raw) >= 10 and raw[:3] == b"GIF":
            w_px, h_px = struct.unpack("<HH", raw[6:10])
        elif len(raw) >= 26 and raw[:2] == b"BM":
            w_px = struct.unpack("<I", raw[18:22])[0]
            h_px = abs(struct.unpack("<i", raw[22:26])[0])

        if w_px > 0 and h_px > 0:
            px_to_emu = 914400 / 96  # 96 DPI default
            nat_cx = int(w_px * px_to_emu)
            nat_cy = int(h_px * px_to_emu)
            if nat_cx > MAX_W:
                cx = MAX_W
                cy = int(nat_cy * MAX_W / nat_cx)
            else:
                cx, cy = nat_cx, nat_cy
        else:
            cx = MAX_W
            cy = MAX_W * 3 // 4

        return int(cx), int(cy)
    except Exception:
        return int(max_width_emu), int(max_width_emu * 3 // 4)


def _make_inline_drawing_para(rid, cx, cy, img_name, pic_id):
    """Return a <w:p> lxml element containing an inline image drawing."""
    NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    p = etree.Element(f"{{{NS_W}}}p")
    r = etree.SubElement(p, f"{{{NS_W}}}r")
    drawing = etree.SubElement(r, f"{{{NS_W}}}drawing")

    inline = etree.SubElement(
        drawing, f"{{{NS_WP}}}inline",
        attrib={"distT": "0", "distB": "114300", "distL": "0", "distR": "0"},
    )
    etree.SubElement(inline, f"{{{NS_WP}}}extent", attrib={"cx": str(cx), "cy": str(cy)})
    etree.SubElement(inline, f"{{{NS_WP}}}effectExtent", attrib={"l": "0", "t": "0", "r": "0", "b": "0"})
    etree.SubElement(inline, f"{{{NS_WP}}}docPr", attrib={"id": str(pic_id), "name": img_name})
    cnv = etree.SubElement(inline, f"{{{NS_WP}}}cNvGraphicFramePr")
    etree.SubElement(cnv, f"{{{NS_A}}}graphicFrameLocks", attrib={"noChangeAspect": "1"})

    graphic = etree.SubElement(inline, f"{{{NS_A}}}graphic")
    graphicData = etree.SubElement(
        graphic, f"{{{NS_A}}}graphicData",
        attrib={"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"},
    )
    pic_el = etree.SubElement(graphicData, f"{{{NS_PIC}}}pic")
    nvPicPr = etree.SubElement(pic_el, f"{{{NS_PIC}}}nvPicPr")
    etree.SubElement(nvPicPr, f"{{{NS_PIC}}}cNvPr", attrib={"id": str(pic_id), "name": img_name})
    etree.SubElement(nvPicPr, f"{{{NS_PIC}}}cNvPicPr")
    blipFill = etree.SubElement(pic_el, f"{{{NS_PIC}}}blipFill")
    etree.SubElement(blipFill, f"{{{NS_A}}}blip", attrib={f"{{{NS_R}}}embed": rid})
    stretch = etree.SubElement(blipFill, f"{{{NS_A}}}stretch")
    etree.SubElement(stretch, f"{{{NS_A}}}fillRect")
    spPr = etree.SubElement(pic_el, f"{{{NS_PIC}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
    etree.SubElement(xfrm, f"{{{NS_A}}}off", attrib={"x": "0", "y": "0"})
    etree.SubElement(xfrm, f"{{{NS_A}}}ext", attrib={"cx": str(cx), "cy": str(cy)})
    prstGeom = etree.SubElement(spPr, f"{{{NS_A}}}prstGeom", attrib={"prst": "rect"})
    etree.SubElement(prstGeom, f"{{{NS_A}}}avLst")

    return p


def _ensure_media_content_types(buffer, exts):
    """Ensure [Content_Types].xml declares Default content types for image extensions."""
    ct_key = "[Content_Types].xml"
    if ct_key not in buffer:
        return

    content_types = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tif": "image/tiff",
        "tiff": "image/tiff",
    }

    needed = [e.lower().lstrip(".") for e in (exts or []) if e]
    needed = [e for e in needed if e in content_types]
    if not needed:
        return

    try:
        ct_root = etree.fromstring(buffer[ct_key])
    except Exception:
        return

    NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
    ns_ct = {"ct": NS_CT}
    existing = {
        (d.get("Extension", "") or "").lower()
        for d in ct_root.xpath("/ct:Types/ct:Default", namespaces=ns_ct)
    }

    added = False
    for ext in needed:
        if ext in existing:
            continue
        ct_root.append(
            etree.Element(
                f"{{{NS_CT}}}Default",
                attrib={
                    "Extension": ext,
                    "ContentType": content_types[ext],
                },
            )
        )
        added = True

    if added:
        buffer[ct_key] = etree.tostring(ct_root, encoding="UTF-8", xml_declaration=True)


def _extract_snippet_for_signals(text, signals):
    lowered = text.lower()
    first_idx = -1
    used_signal = ""
    for signal in signals:
        idx = lowered.find(signal)
        if idx >= 0 and (first_idx < 0 or idx < first_idx):
            first_idx = idx
            used_signal = signal

    if first_idx < 0:
        return ""

    start = max(0, first_idx - 180)
    end = min(len(text), first_idx + 420)
    snippet = text[start:end].strip()
    if start > 0:
        snippet = "..." + snippet
    if end < len(text):
        snippet = snippet + "..."
    return f"({used_signal}) {snippet}"


def _collect_related_evidence(source_dir, source_docx, sd_name):
    if not source_dir or not source_dir.exists() or not source_dir.is_dir():
        return {}, {}

    candidates = []
    for file_path in source_dir.rglob("*"):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~$"):
            continue
        if source_docx and file_path.resolve() == source_docx.resolve():
            continue
        if file_path.suffix.lower() not in SUPPORTED_RELATED_EXTENSIONS:
            continue
        candidates.append(file_path)
        if len(candidates) >= MAX_RELATED_FILES:
            break

    evidence_pool = {tag: [] for tag in TAG_SIGNALS.keys()}

    for file_path in candidates:
        raw = _read_related_file_text(file_path)
        raw = " ".join((raw or "").split())
        artifacts = _extract_related_artifacts(file_path)
        table_snippets = artifacts.get("tables", [])

        artifact_text = " ".join(table_snippets).strip()
        combined_raw = " ".join(x for x in [raw, artifact_text] if x).strip()
        if not combined_raw:
            continue

        limited = combined_raw[:MAX_CHARS_PER_RELATED_FILE]
        lowered = limited.lower()

        for tag, signals in TAG_SIGNALS.items():
            matched = [s for s in signals if s in lowered]
            if not matched:
                continue

            snippet = _extract_snippet_for_signals(limited, matched)
            if not snippet:
                continue

            score = len(matched)
            table_line = f"- Tabel uit {file_path.name}: {_clip(table_snippets[0], 320)}" if table_snippets else ""
            evidence_pool[tag].append((score, file_path.name, snippet, table_line))

    evidence_by_tag = {}
    for tag, items in evidence_pool.items():
        if not items:
            continue

        items.sort(key=lambda x: (-x[0], len(x[2])))
        top_items = items[:2]
        lines = []
        for _, name, snippet, table_line in top_items:
            lines.append(f"- Bron {name}: {_clip(snippet, 320)}")
            if table_line:
                lines.append(table_line)
        evidence_by_tag[tag] = "\n".join(lines)

    if evidence_by_tag:
        log(f"Gerelateerde documenten geanalyseerd: {len(candidates)} bestand(en), evidence voor {len(evidence_by_tag)} tag(s)", sd_name)
    else:
        log(f"Gerelateerde documenten geanalyseerd: {len(candidates)} bestand(en), geen bruikbare evidence gevonden", sd_name)

    return evidence_by_tag, {}


def _extract_sd_images_by_section(source_docx, sd_media_root):
    """Extract images from the SD source DOCX, grouped by the SDT tag of the
    section heading they appear under.  Only images larger than 4 KB (to skip
    tiny icons/decorations) are kept.  Each image file is assigned to at most
    one tag (document order wins).
    Returns: {sdt_tag: [(Path, source_filename), ...]}
    """
    if not source_docx:
        return {}
    source_docx = Path(source_docx)
    if not source_docx.exists():
        return {}

    sd_media_root = Path(sd_media_root)
    sd_media_root.mkdir(parents=True, exist_ok=True)

    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_w = {"w": NS_W}

    images_by_tag = {}
    globally_used = set()

    try:
        with ZipFile(source_docx, "r") as z:
            all_entries = set(z.namelist())

            rels_key = "word/_rels/document.xml.rels"
            if rels_key not in all_entries:
                return {}

            rels_root = etree.fromstring(z.read(rels_key))
            rId_to_media = {
                rel.get("Id", ""): rel.get("Target", "")
                for rel in rels_root
                if "image" in rel.get("Type", "")
            }

            doc_root = etree.fromstring(z.read("word/document.xml"))
            body = doc_root.find(f"{{{NS_W}}}body")
            if body is None:
                return {}

            current_tag = None

            for elem in body:
                tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

                # Update current section tag when we encounter a heading paragraph
                if tag_local == "p":
                    pStyle = elem.xpath(".//w:pStyle/@w:val", namespaces=ns_w)
                    sval = (pStyle[0] if pStyle else "").lower()
                    if "heading" in sval or sval.startswith("kop"):
                        texts = elem.xpath(".//w:t/text()", namespaces=ns_w)
                        heading = " ".join(t.strip() for t in texts if t and t.strip())
                        current_tag = resolve_sdt_tag(heading) if heading else None

                if current_tag is None:
                    continue

                # Collect images (blip elements) within this element subtree
                for blip in elem.xpath(".//*[local-name()='blip']"):
                    rId = blip.get(f"{{{NS_R}}}embed", "")
                    media_rel = rId_to_media.get(rId, "")
                    if not media_rel or media_rel in globally_used:
                        continue
                    suffix = Path(media_rel).suffix.lower()
                    if suffix not in IMAGE_EXTENSIONS:
                        continue
                    media_key = f"word/{media_rel}"
                    if media_key not in all_entries:
                        continue
                    raw = z.read(media_key)
                    if len(raw) < 4096:  # skip tiny icons / decorations
                        continue
                    out_name = f"sd_{_short_id(media_rel, 8)}{suffix}"
                    out_path = sd_media_root / out_name
                    try:
                        out_path.write_bytes(raw)
                    except Exception:
                        continue
                    globally_used.add(media_rel)
                    tag_list = images_by_tag.setdefault(current_tag, [])
                    if len(tag_list) < MAX_IMAGES_PER_TAG:
                        tag_list.append((out_path, source_docx.name))

    except Exception:
        return {}

    return images_by_tag


def _inject_images_into_docx_buffer(buffer, related_images_by_tag, sd_name):
    """Inject relevant images inline into the matching SDT content blocks in the OOXML buffer.
    Adds media files and relationship entries to the buffer; no new chapters are created."""
    if not related_images_by_tag:
        return 0

    doc_key = "word/document.xml"
    rels_key = "word/_rels/document.xml.rels"
    if doc_key not in buffer or rels_key not in buffer:
        return 0

    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    IMAGE_REL_TYPE = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    ns = {"w": NS_W, "wp": NS_WP, "pic": NS_PIC}

    try:
        xml_root = etree.fromstring(buffer[doc_key])
    except Exception as exc:
        log(f"Kon document.xml niet parsen voor image-injectie: {exc}", sd_name)
        return 0

    try:
        rels_root = etree.fromstring(buffer[rels_key])
    except Exception as exc:
        log(f"Kon document.xml.rels niet parsen voor image-injectie: {exc}", sd_name)
        return 0

    existing_rids = []
    for rel in rels_root:
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            existing_rids.append(int(rid[3:]))
    next_rid_num = max(existing_rids, default=0) + 1

    existing_docpr_ids = []
    for raw_id in xml_root.xpath(".//wp:docPr/@id", namespaces=ns):
        try:
            existing_docpr_ids.append(int(raw_id))
        except Exception:
            pass
    for raw_id in xml_root.xpath(".//pic:cNvPr/@id", namespaces=ns):
        try:
            existing_docpr_ids.append(int(raw_id))
        except Exception:
            pass
    next_pic_id = max(existing_docpr_ids, default=0) + 1
    embedded = 0
    injected_exts = set()

    for tag, img_entries in related_images_by_tag.items():
        sdts = xml_root.xpath(
            f".//w:sdt[w:sdtPr/w:tag[@w:val='{tag}']]", namespaces=ns
        )
        if not sdts:
            sdts = xml_root.xpath(
                f".//w:sdt[.//w:tag[@w:val='{tag}']]", namespaces=ns
            )
        if not sdts:
            continue

        sdt_content = sdts[0].find("w:sdtContent", ns)
        if sdt_content is None:
            continue

        for img_path, source_name in img_entries[:MAX_IMAGES_PER_TAG]:
            img_path = Path(img_path)
            if not img_path.exists():
                continue

            img_suffix = img_path.suffix.lower()
            injected_exts.add(img_suffix)
            media_name = f"inj_{next_rid_num}_{_short_id(str(img_path), 6)}{img_suffix}"
            media_key = f"word/media/{media_name}"
            buffer[media_key] = img_path.read_bytes()

            rid = f"rId{next_rid_num}"
            rels_root.append(
                etree.Element(
                    f"{{{NS_PKG_REL}}}Relationship",
                    attrib={
                        "Id": rid,
                        "Type": IMAGE_REL_TYPE,
                        "Target": f"media/{media_name}",
                    },
                )
            )

            # Caption paragraph
            W = f"{{{NS_W}}}"
            caption_p = etree.SubElement(sdt_content, f"{W}p")
            caption_r = etree.SubElement(caption_p, f"{W}r")
            caption_t = etree.SubElement(caption_r, f"{W}t")
            caption_t.text = f"[Bron: {source_name}]"

            # Inline drawing paragraph
            cx, cy = _read_img_dimensions_emu(img_path)
            draw_para = _make_inline_drawing_para(rid, cx, cy, media_name, next_pic_id)
            sdt_content.append(draw_para)

            next_rid_num += 1
            next_pic_id += 1
            embedded += 1

    if embedded:
        _ensure_media_content_types(buffer, injected_exts)
        buffer[doc_key] = etree.tostring(xml_root, encoding="UTF-8", xml_declaration=True)
        buffer[rels_key] = etree.tostring(rels_root, encoding="UTF-8", xml_declaration=True)

    return embedded


def _parse_product_fields_from_sd_name(sd_name):
    base = sd_name
    if base.endswith("_mapped"):
        base = base[:-7]

    base = re.sub(r"^SD\s*-\s*", "", base, flags=re.IGNORECASE).strip()
    code_blocks = re.findall(r"\[[^\]]+\]", base)
    product_code = "".join(code_blocks).strip()
    product_name = re.sub(r"\s*\[[^\]]+\]", "", base).strip()
    product_name = re.sub(r"\s+", " ", product_name).strip()
    return product_name, product_code


def _set_cover_fields(buffer, sd_name):
    product_name, product_code = _parse_product_fields_from_sd_name(sd_name)
    if not product_name:
        return

    ns_word = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ns_core = {
        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dc": "http://purl.org/dc/elements/1.1/",
    }

    # 1) Update Word core properties used by data-bound footer fields.
    core_name = "docProps/core.xml"
    if core_name in buffer:
        try:
            core_root = etree.fromstring(buffer[core_name])
            title_nodes = core_root.xpath("/cp:coreProperties/dc:title", namespaces=ns_core)
            if title_nodes:
                title_nodes[0].text = sanitize_xml_text(product_name)

            description_nodes = core_root.xpath("/cp:coreProperties/dc:description", namespaces=ns_core)
            if description_nodes:
                description_nodes[0].text = sanitize_xml_text(product_code)

            buffer[core_name] = etree.tostring(core_root, encoding="UTF-8", xml_declaration=True)
        except Exception:
            pass

    # 2) Replace cover-page filename field with product name and bind to Title property.
    doc_name = "word/document.xml"
    if doc_name in buffer:
        try:
            doc_root = etree.fromstring(buffer[doc_name])
            filename_paras = doc_root.xpath(
                ".//w:p[.//w:instrText[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'FILENAME')]]",
                namespaces=ns_word,
            )

            for p in filename_paras:
                instr_nodes = p.xpath(".//w:instrText", namespaces=ns_word)
                for instr in instr_nodes:
                    text = (instr.text or "")
                    if "FILENAME" in text.upper():
                        instr.text = " DOCPROPERTY  Title  \\* MERGEFORMAT "

                text_nodes = p.xpath(".//w:t", namespaces=ns_word)
                if text_nodes:
                    text_nodes[0].text = sanitize_xml_text(product_name)
                    for node in text_nodes[1:]:
                        node.text = ""

            buffer[doc_name] = etree.tostring(doc_root, encoding="UTF-8", xml_declaration=True)
        except Exception:
            pass

    # 3) Ensure footer data-bound visible values are updated immediately.
    footer_parts = [name for name in buffer if name.startswith("word/footer") and name.endswith(".xml")]
    for footer_name in footer_parts:
        try:
            footer_root = etree.fromstring(buffer[footer_name])

            title_sdts = footer_root.xpath(
                ".//w:sdt[w:sdtPr/w:alias[@w:val='Title']]",
                namespaces=ns_word,
            )
            for sdt in title_sdts:
                content = sdt.find("w:sdtContent", ns_word)
                if content is not None:
                    _set_sdt_text_preserving_structure(content, product_name, ns_word)

            comments_sdts = footer_root.xpath(
                ".//w:sdt[w:sdtPr/w:alias[@w:val='Comments']]",
                namespaces=ns_word,
            )
            for sdt in comments_sdts:
                content = sdt.find("w:sdtContent", ns_word)
                if content is not None:
                    _set_sdt_text_preserving_structure(content, product_code, ns_word)

            buffer[footer_name] = etree.tostring(footer_root, encoding="UTF-8", xml_declaration=True)
        except Exception:
            continue


def _set_sdt_text_preserving_structure(content, new_text, ns):
    def _append_paragraph(parent, text, w_ns, bold=False):
        p = etree.SubElement(parent, f"{{{w_ns}}}p")
        r = etree.SubElement(p, f"{{{w_ns}}}r")
        if bold:
            r_pr = etree.SubElement(r, f"{{{w_ns}}}rPr")
            etree.SubElement(r_pr, f"{{{w_ns}}}b")
        t = etree.SubElement(r, f"{{{w_ns}}}t")
        t.text = text

    def _append_table(parent, rows, w_ns):
        if not rows or len(rows) < 2:
            return

        max_cols = max(len(r) for r in rows)
        if max_cols < 2:
            return

        normalized_rows = []
        for row in rows:
            current = [str(c or "").strip() for c in row]
            if len(current) < max_cols:
                current.extend([""] * (max_cols - len(current)))
            normalized_rows.append(current)

        tbl = etree.SubElement(parent, f"{{{w_ns}}}tbl")
        tbl_pr = etree.SubElement(tbl, f"{{{w_ns}}}tblPr")
        etree.SubElement(tbl_pr, f"{{{w_ns}}}tblW", attrib={f"{{{w_ns}}}type": "auto", f"{{{w_ns}}}w": "0"})

        # Basic visible borders so table structure is clear in generated guides.
        borders = etree.SubElement(tbl_pr, f"{{{w_ns}}}tblBorders")
        for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            etree.SubElement(
                borders,
                f"{{{w_ns}}}{side}",
                attrib={f"{{{w_ns}}}val": "single", f"{{{w_ns}}}sz": "4", f"{{{w_ns}}}space": "0", f"{{{w_ns}}}color": "auto"},
            )

        for ridx, row in enumerate(normalized_rows):
            tr = etree.SubElement(tbl, f"{{{w_ns}}}tr")
            for cell_value in row:
                tc = etree.SubElement(tr, f"{{{w_ns}}}tc")
                tc_pr = etree.SubElement(tc, f"{{{w_ns}}}tcPr")
                etree.SubElement(tc_pr, f"{{{w_ns}}}tcW", attrib={f"{{{w_ns}}}type": "auto", f"{{{w_ns}}}w": "0"})
                p = etree.SubElement(tc, f"{{{w_ns}}}p")
                r = etree.SubElement(p, f"{{{w_ns}}}r")
                t = etree.SubElement(r, f"{{{w_ns}}}t")
                t.text = cell_value

                if ridx == 0:
                    r_pr = etree.SubElement(r, f"{{{w_ns}}}rPr")
                    etree.SubElement(r_pr, f"{{{w_ns}}}b")

    def _parse_pipe_row(line):
        pieces = [p.strip() for p in str(line or "").split("|") if p and p.strip()]
        return pieces if len(pieces) >= 2 else None

    def _parse_key_value_row(line):
        raw = str(line or "").strip()
        raw = re.sub(r"^[-*]\s*", "", raw)
        if "=" not in raw and ":" not in raw:
            return None

        out = []
        for seg in [x.strip() for x in raw.split(";") if x.strip()]:
            if "=" in seg:
                k, v = seg.split("=", 1)
            elif ":" in seg:
                k, v = seg.split(":", 1)
            else:
                continue
            k = k.strip()
            v = v.strip()
            if k and v:
                out.append((k, v))

        return out if len(out) >= 2 else None

    def _extract_table_candidate(lines):
        pipe_rows = []
        pipe_line_idxs = []
        for idx, line in enumerate(lines):
            chunks = [c.strip() for c in str(line or "").split(";") if c.strip()]
            if not chunks:
                chunks = [line]
            for chunk in chunks:
                parsed = _parse_pipe_row(chunk)
                if parsed:
                    pipe_rows.append(parsed)
                    pipe_line_idxs.append(idx)

        if len(pipe_rows) >= 2:
            return {
                "rows": pipe_rows,
                "start_idx": min(pipe_line_idxs),
                "end_idx": max(pipe_line_idxs),
            }

        kv_rows = []
        kv_line_idxs = []
        for idx, line in enumerate(lines):
            parsed = _parse_key_value_row(line)
            if parsed:
                kv_rows.append(parsed)
                kv_line_idxs.append(idx)

        if len(kv_rows) >= 2:
            headers = []
            seen = set()
            for pairs in kv_rows:
                for k, _ in pairs:
                    lk = k.lower()
                    if lk in seen:
                        continue
                    seen.add(lk)
                    headers.append(k)

            matrix = [headers]
            for pairs in kv_rows:
                lookup = {k.lower(): v for k, v in pairs}
                matrix.append([lookup.get(h.lower(), "") for h in headers])

            return {
                "rows": matrix,
                "start_idx": min(kv_line_idxs),
                "end_idx": max(kv_line_idxs),
            }

        return None

    text_nodes = content.xpath(".//w:t", namespaces=ns)
    safe_text = sanitize_xml_text(new_text).strip()
    raw_lines = [line.rstrip() for line in sanitize_xml_text(new_text).splitlines()]

    lines = [line.strip() for line in safe_text.splitlines() if line and line.strip()]
    table_candidate = _extract_table_candidate(lines) if lines else None

    if table_candidate:
        wp = ns["w"]

        # Rebuild SDT content to keep paragraph context and render detected rows as a true Word table.
        for child in list(content):
            content.remove(child)

        start_idx = table_candidate["start_idx"]
        end_idx = table_candidate["end_idx"]

        for pre in lines[:start_idx]:
            _append_paragraph(content, pre, wp)

        _append_table(content, table_candidate["rows"], wp)

        for post in lines[end_idx + 1:]:
            _append_paragraph(content, post, wp)
        return

    # Keep multi-line structured content readable in Word by writing one paragraph per line.
    if len(raw_lines) > 1:
        wp = ns["w"]
        for child in list(content):
            content.remove(child)

        for raw in raw_lines:
            line = str(raw or "").strip()
            if not line:
                _append_paragraph(content, "", wp)
                continue

            is_heading = line.endswith(":") and len(line.split()) <= 8
            if line.startswith("- "):
                _append_paragraph(content, f"• {line[2:].strip()}", wp)
            else:
                _append_paragraph(content, line, wp, bold=is_heading)
        return

    if text_nodes:
        text_nodes[0].text = safe_text
        for node in text_nodes[1:]:
            node.text = ""
        return

    wp = ns["w"]
    first_child = next(iter(content), None)
    first_tag = etree.QName(first_child.tag).localname if first_child is not None else None

    if first_tag == "r":
        r = etree.Element(f"{{{wp}}}r")
        t = etree.SubElement(r, f"{{{wp}}}t")
        t.text = safe_text
        content.append(r)
        return

    p = etree.Element(f"{{{wp}}}p")
    r = etree.SubElement(p, f"{{{wp}}}r")
    t = etree.SubElement(r, f"{{{wp}}}t")
    t.text = safe_text
    content.append(p)


def replace_sdt(xml_root, tag_name, new_text):
    NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdts = xml_root.xpath(f".//w:sdt[w:sdtPr/w:tag[@w:val='{tag_name}']]", namespaces=NS)

    count = 0
    for sdt in sdts:
        content = sdt.find("w:sdtContent", NS)
        if content is None:
            continue

        _set_sdt_text_preserving_structure(content, new_text, NS)
        count += 1

    return count


def process_docx(buffer, xml_root_data, sd_name, source_dir=None, source_docx=None):
    total = 0
    xml_sections = xml_root_data.xpath("//Section")
    filled_tags = set()
    quality_logged_tags = set()
    force_open_tags = set()
    collected_content = {}
    selected_sources_by_tag = _select_best_sources_per_tag(xml_sections)
    validated_sources_by_tag, rejected_sources_by_tag = _validate_selected_sources(selected_sources_by_tag)
    conflict_by_tag = _find_conflicts(validated_sources_by_tag)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    related_evidence, _ = _collect_related_evidence(source_dir, source_docx, sd_name)
    gold_examples_by_tag = _load_gold_examples_by_tag(sd_name)
    style_profiles_by_tag = _build_style_profiles(gold_examples_by_tag)
    sd_media_root = Path("output") / "sd_media" / _short_id(sd_name, 10)
    related_images_by_tag = _extract_sd_images_by_section(source_docx, sd_media_root) if source_docx else {}

    if selected_sources_by_tag:
        for tag in TEMPLATE_TAG_ORDER:
            sources = validated_sources_by_tag.get(tag, [])
            if not sources:
                rejected = rejected_sources_by_tag.get(tag, [])
                if rejected:
                    rejected_trace = " | ".join(
                        f"{r['section_name']} ({r.get('intent_severity', SEVERITY_WARNING).upper()}: {_clip(r.get('intent_reason', ''), 120)})" for r in rejected
                    )
                    log(f"Rejected sources for SDT '{tag}': {rejected_trace}", sd_name)
                continue

            trace = " | ".join(
                f"{s['section_name']} ({_clip(s['reason'], 90)}; {s.get('intent_severity', SEVERITY_INFO).upper()}: {_clip(s.get('intent_reason', ''), 90)})" for s in sources
            )
            log(f"Selected sources for SDT '{tag}': {trace}", sd_name)
            tag_conflicts = conflict_by_tag.get(tag, [])
            for c in tag_conflicts:
                log(
                    f"Conflict check for SDT '{tag}' [severity={c.get('severity', SEVERITY_WARNING)}] "
                    f"[{c.get('code', 'n/a')}]: {c.get('message', '')}",
                    sd_name,
                )

    if related_images_by_tag:
        log(f"SD-afbeeldingen per sectie gevonden: {sum(len(v) for v in related_images_by_tag.values())} afbeelding(en) in {len(related_images_by_tag)} sectie(s)", sd_name)

    for part_name, xml_bytes in list(buffer.items()):

        # Only Word XML parts
        if not (part_name.startswith("word/") and part_name.endswith(".xml")):
            continue

        try:
            xml_root = etree.fromstring(xml_bytes)
        except:
            continue

        # Fill chapters directly from best semantic source sections (top 1..3) per SDT.
        for tag in TEMPLATE_TAG_ORDER:
            if tag in filled_tags:
                continue

            sources = validated_sources_by_tag.get(tag, [])
            if not sources:
                if tag in STRICT_EVIDENCE_TAGS:
                    force_open_tags.add(tag)
                    log(
                        f"Strict evidence check for SDT '{tag}' failed: no selected source evidence. Forcing open placeholder.",
                        sd_name,
                    )
                continue

            has_exact, exact_reason = _has_exact_evidence_for_tag(tag, sources)
            if not has_exact:
                force_open_tags.add(tag)
                log(
                    f"Strict evidence check for SDT '{tag}' failed: {exact_reason}. Forcing open placeholder.",
                    sd_name,
                )
                continue

            tag_conflicts = conflict_by_tag.get(tag, [])
            has_blocking = any(c.get("severity") == SEVERITY_BLOCKING for c in tag_conflicts)
            if has_blocking:
                if tag in STRICT_EVIDENCE_TAGS:
                    force_open_tags.add(tag)
                for c in tag_conflicts:
                    if c.get("severity") == SEVERITY_BLOCKING:
                        log(
                            f"Skipped direct fill for SDT '{tag}' due to blocking conflict [{c.get('code', 'n/a')}]: {c.get('message', '')}",
                            sd_name,
                        )
                continue

            merged_content = _merge_selected_source_content(tag, sources)
            merged_content = _strip_ai_generated_markers(merged_content)
            if not merged_content:
                continue

            replaced = replace_sdt(xml_root, tag, merged_content)
            if replaced > 0:
                total += replaced
                filled_tags.add(tag)
                collected_content[tag] = "\n\n".join(s["content"] for s in sources)
                source_headers = " | ".join(s["section_name"] for s in sources)
                trace_reason = " | ".join(
                    _clip(f"{s['reason']}; {s.get('intent_reason', '')}", 180) for s in sources
                )
                log(f"Filled SDT '{tag}' with XML section '{source_headers}'", sd_name)
                log(f"Trace SDT '{tag}': {trace_reason}", sd_name)
                for c in conflict_by_tag.get(tag, []):
                    log(
                        f"Trace SDT '{tag}' conflict [severity={c.get('severity', SEVERITY_WARNING)}] "
                        f"[{c.get('code', 'n/a')}]: {c.get('message', '')}",
                        sd_name,
                    )
                if tag not in quality_logged_tags:
                    quality = _compute_quality_for_tag(
                        tag_name=tag,
                        content=merged_content,
                        sources=sources,
                        conflicts=conflict_by_tag.get(tag, []),
                        has_exact_evidence=has_exact,
                        fill_type="direct_from_sd_chapter",
                    )
                    _log_quality_for_tag(tag, quality, sd_name)
                    quality_logged_tags.add(tag)

        # Fill missing template chapters with controlled AI fallback text.
        for tag in TEMPLATE_TAG_ORDER:
            if tag in filled_tags:
                continue

            if tag in STRICT_EVIDENCE_TAGS or tag in force_open_tags:
                generated = _build_low_info_text()
                quality = _compute_quality_for_tag(
                    tag_name=tag,
                    content=generated,
                    sources=[],
                    conflicts=conflict_by_tag.get(tag, []),
                    has_exact_evidence=False,
                    fill_type="open_too_little_info",
                )
                generated_for_output = _decorate_ai_generated_text_with_quality(
                    tag,
                    generated,
                    quality,
                    fill_type="open_too_little_info",
                )
                replaced = replace_sdt(xml_root, tag, generated_for_output)
                if replaced > 0:
                    total += replaced
                    filled_tags.add(tag)
                    log(f"Forced open placeholder for hoofdstuk '{tag}' wegens ontbrekende exacte bron-evidence", sd_name)
                    if tag not in quality_logged_tags:
                        _log_quality_for_tag(tag, quality, sd_name)
                        quality_logged_tags.add(tag)
                continue

            direct_related = related_evidence.get(tag, "").strip()
            style_profile = style_profiles_by_tag.get(tag, {})
            if direct_related:
                anchored_related = _apply_style_anchor_to_lines(
                    [x for x in str(direct_related).splitlines() if str(x or "").strip()],
                    style_profile,
                )
                generated = "\n".join([
                    HITL_PREFIX,
                    "Voorlopige inhoud afgeleid uit gerelateerde productdocumenten:",
                    anchored_related,
                    "Gelieve dit hoofdstuk inhoudelijk te verifieren en aan te vullen via HITL.",
                ])
            else:
                generated = _build_generated_text(
                    tag,
                    collected_content,
                    xml_sections,
                    style_profile=style_profile,
                )

            if LOW_INFO_TEXT in generated:
                fill_type = "open_too_little_info"
            elif direct_related:
                fill_type = "ai_related_documents"
            else:
                fill_type = "ai_missing_chapter"

            quality = _compute_quality_for_tag(
                tag_name=tag,
                content=generated,
                sources=[],
                conflicts=conflict_by_tag.get(tag, []),
                has_exact_evidence=(tag not in STRICT_EVIDENCE_TAGS),
                fill_type=fill_type,
            )
            generated_for_output = _decorate_ai_generated_text_with_quality(
                tag,
                generated,
                quality,
                fill_type=fill_type,
            )

            replaced = replace_sdt(xml_root, tag, generated_for_output)
            if replaced > 0:
                total += replaced
                filled_tags.add(tag)
                if LOW_INFO_TEXT in generated:
                    log(f"AI open wegens te weinig info voor hoofdstuk '{tag}'", sd_name)
                elif direct_related:
                    log(f"AI aangevuld op basis van gerelateerde documenten voor hoofdstuk '{tag}'", sd_name)
                else:
                    log(f"AI aangevuld voor ontbrekend hoofdstuk '{tag}'", sd_name)

                if tag not in quality_logged_tags:
                    _log_quality_for_tag(tag, quality, sd_name)
                    quality_logged_tags.add(tag)

        # Ensure no tagged template placeholders remain (including subchapters like 1.X fields).
        tagged_sdts = xml_root.xpath(".//w:sdt", namespaces=ns)
        for sdt in tagged_sdts:
            tag_vals = sdt.xpath("./w:sdtPr/w:tag/@w:val", namespaces=ns)
            tag = tag_vals[0].strip() if tag_vals else ""
            if not tag or tag in NON_CHAPTER_TAGS:
                continue

            content = sdt.find("w:sdtContent", ns)
            if content is None:
                continue

            texts = content.xpath(".//w:t/text()", namespaces=ns)
            current_text = " ".join(t.strip() for t in texts if t and t.strip()).strip()
            current_norm = current_text.lower()

            is_placeholder = (
                (not current_text)
                or ("vul hier de inhoud in voor:" in current_norm)
                or ("[to be completed]" in current_norm)
            )
            if not is_placeholder:
                continue

            direct_related = related_evidence.get(tag, "").strip()
            style_profile = style_profiles_by_tag.get(tag, {})
            if tag in STRICT_EVIDENCE_TAGS or tag in force_open_tags:
                generated = _build_low_info_text()
            elif direct_related:
                anchored_related = _apply_style_anchor_to_lines(
                    [x for x in str(direct_related).splitlines() if str(x or "").strip()],
                    style_profile,
                )
                generated = "\n".join([
                    HITL_PREFIX,
                    "Voorlopige inhoud afgeleid uit gerelateerde productdocumenten:",
                    anchored_related,
                    "Gelieve dit hoofdstuk inhoudelijk te verifieren en aan te vullen via HITL.",
                ])
            else:
                generated = _build_generated_text(
                    tag,
                    collected_content,
                    xml_sections,
                    style_profile=style_profile,
                )

            if LOW_INFO_TEXT in generated:
                fill_type = "open_too_little_info"
            else:
                fill_type = "ai_related_documents"

            quality = _compute_quality_for_tag(
                tag_name=tag,
                content=generated,
                sources=[],
                conflicts=conflict_by_tag.get(tag, []),
                has_exact_evidence=(tag not in STRICT_EVIDENCE_TAGS),
                fill_type=fill_type,
            )
            generated_for_output = _decorate_ai_generated_text_with_quality(
                tag,
                generated,
                quality,
                fill_type=fill_type,
            )

            _set_sdt_text_preserving_structure(content, generated_for_output, ns)
            total += 1
            filled_tags.add(tag)
            if LOW_INFO_TEXT in generated:
                log(f"AI open wegens te weinig info voor hoofdstuk '{tag}'", sd_name)
            else:
                log(f"AI aangevuld op basis van gerelateerde documenten voor hoofdstuk '{tag}'", sd_name)

            if tag not in quality_logged_tags:
                _log_quality_for_tag(tag, quality, sd_name)
                quality_logged_tags.add(tag)

        # Replace part in buffer
        buffer[part_name] = etree.tostring(xml_root, encoding="UTF-8", xml_declaration=True)

    return total, related_images_by_tag


# =====================================================================
#  MAIN — met correcte ZIP-merging fix (Word compatible!)
# =====================================================================
if __name__ == "__main__":
    xml_file = Path(sys.argv[1])
    template = Path("templates/presales_template_sdt_v2.docx")
    sd_name = xml_file.stem
    source_docx = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else None
    source_dir = source_docx.parent if source_docx else None

    log(f"START xml_to_docx_v3_fixed for: {xml_file}", sd_name)

    if not template.exists():
        log(f"ERROR: Template not found: {template}", sd_name)
        sys.exit(1)

    xml_tree = etree.parse(str(xml_file))
    xml_root = xml_tree.getroot()

    with ZipFile(template, "r") as zin:
        original = {n: zin.read(n) for n in zin.namelist()}

    # make a buffer we can modify
    buffer = dict(original)

    # process all xml parts
    total_filled, related_images_by_tag = process_docx(
        buffer,
        xml_root,
        sd_name,
        source_dir=source_dir,
        source_docx=source_docx,
    )
    _set_cover_fields(buffer, sd_name)
    log("Covervelden ingevuld op basis van bestandsnaam", sd_name)
    log(f"TOTAL SDT FIELDS FILLED: {total_filled}", sd_name)

    # Inject relevant images inline into matching SDT content blocks (before ZIP is written).
    added_images = _inject_images_into_docx_buffer(buffer, related_images_by_tag, sd_name)
    if added_images > 0:
        log(f"Afbeeldingen inline ingevoegd in bijbehorende tekst: {added_images}", sd_name)

    # SAFE DOCX REBUILD FIX
    out_folder = Path("output/docx")
    out_folder.mkdir(exist_ok=True)
    out_file = out_folder / f"{sd_name}_FINAL.docx"

    try:
        with ZipFile(template, "r") as zin:
            with ZipFile(out_file, "w") as zout:
                infos = list(zin.infolist())
                template_names = {i.filename for i in infos}
                for info in infos:
                    zout.writestr(info, buffer.get(info.filename, zin.read(info.filename)))
                for name, data in buffer.items():
                    if name not in template_names:
                        zout.writestr(name, data)
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = out_folder / f"{sd_name}_FINAL_{ts}.docx"
        log(f"Output file locked, using fallback: {out_file.name}", sd_name)
        with ZipFile(template, "r") as zin:
            with ZipFile(out_file, "w") as zout:
                infos = list(zin.infolist())
                template_names = {i.filename for i in infos}
                for info in infos:
                    zout.writestr(info, buffer.get(info.filename, zin.read(info.filename)))
                for name, data in buffer.items():
                    if name not in template_names:
                        zout.writestr(name, data)

    log(f"XML to DOCX OK → {out_file}", sd_name)
    sys.exit(0)