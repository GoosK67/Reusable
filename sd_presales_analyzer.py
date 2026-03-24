#!/usr/bin/env python3
"""
SD Presales Usefulness Analyzer
Evaluates how useful each SD content is for presales guide generation
"""

import os
import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from docx import Document
import pandas as pd
import yaml

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

ROOT_FOLDER = r"C:\Users\koengo\Cegeka\Product Management - Product Management Library"
TEMPLATE_PATH = r"templates\presales_template.md"
MAPPING_PATH = r"rules\field_mapping.yaml"
OUTPUT_FILE = "sd_presales_usefulness.xlsx"

# Template sections that must be filled
REQUIRED_TEMPLATE_SECTIONS = [
    "Product Summary",
    "Understanding the Client's Needs",
    "Product Description",
    "Architectural Description",
    "Key Features & Functionalities",
    "Scope / Out-of-Scope",
    "Requirements & Prerequisites",
    "Value Proposition",
    "Key Differentiators",
    "Transition & Transformation",
    "Client Responsibilities",
    "Operational Support",
    "Terms & Conditions",
    "SLA & KPI Management",
    "Pricing Elements",
]


SECTION_KEYWORDS = {
    "Product Summary": [
        "service description", "service overview", "service introduction", "summary", "introduction"
    ],
    "Understanding the Client's Needs": [
        "client need", "customer need", "business need", "stakeholder", "pain point", "challenge"
    ],
    "Product Description": [
        "description", "offering", "service offering", "solution description", "product description"
    ],
    "Architectural Description": [
        "architecture", "architectural", "solution design", "technical design", "platform architecture"
    ],
    "Key Features & Functionalities": [
        "feature", "functionality", "capability", "technical implementation", "management service"
    ],
    "Scope / Out-of-Scope": [
        "scope", "out of scope", "in scope", "exclusion", "included", "not included"
    ],
    "Requirements & Prerequisites": [
        "requirement", "prerequisite", "assumption", "eligibility", "dependency"
    ],
    "Value Proposition": [
        "value proposition", "business value", "added business value", "benefit", "business outcome", "roi"
    ],
    "Key Differentiators": [
        "differentiator", "unique", "advantage", "competitive", "why cegeka"
    ],
    "Transition & Transformation": [
        "transition", "transformation", "migration", "onboarding", "roll out", "rollout", "adoption"
    ],
    "Client Responsibilities": [
        "client responsibility", "customer responsibility", "raci", "responsibility matrix", "customer role"
    ],
    "Operational Support": [
        "operational", "support", "run service", "incident", "service desk", "operations"
    ],
    "Terms & Conditions": [
        "term", "condition", "limiting condition", "contract", "legal", "compliance"
    ],
    "SLA & KPI Management": [
        "sla", "kpi", "service level", "availability", "target", "performance metric", "metric"
    ],
    "Pricing Elements": [
        "pricing", "price", "cost", "billing", "financial", "financial information", "per user", "per month"
    ],
}


RULE_TO_TEMPLATE_SECTIONS = {
    "title": ["Product Summary"],
    "service_summary": ["Product Summary", "Product Description"],
    "key_features": ["Key Features & Functionalities"],
    "standard_services": ["Operational Support"],
    "optional_services": ["Scope / Out-of-Scope"],
    "operational_services": ["Operational Support"],
    "prerequisites": ["Requirements & Prerequisites"],
    "out_of_scope": ["Scope / Out-of-Scope"],
    "conditions": ["Terms & Conditions"],
    "sla": ["SLA & KPI Management"],
    "pricing": ["Pricing Elements"],
    "risks": ["Requirements & Prerequisites"],
    "assumptions": ["Requirements & Prerequisites"],
    "differentiators": ["Key Differentiators"],
}


def normalize_text(text: str) -> str:
    """Lowercase and collapse punctuation/whitespace for robust matching."""
    lowered = (text or "").lower()
    cleaned = re.sub(r"[^a-z0-9]+", " ", lowered)
    return re.sub(r"\s+", " ", cleaned).strip()


def normalize_windows_path(path_str: str) -> str:
    """Return a Windows path usable beyond MAX_PATH when needed."""
    if os.name != "nt":
        return path_str

    # Already normalized long path.
    if path_str.startswith("\\\\?\\"):
        return path_str

    abs_path = os.path.abspath(path_str)

    # UNC path: \\server\share -> \\?\UNC\server\share
    if abs_path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_path.lstrip("\\")

    # Local drive path: C:\... -> \\?\C:\...
    return "\\\\?\\" + abs_path


def path_exists(path_str: str) -> bool:
    """Check existence using standard and long-path aware checks."""
    return os.path.exists(path_str) or os.path.exists(normalize_windows_path(path_str))


def build_workspace_docx_index() -> Dict[str, str]:
    """Build filename->path index for DOCX files available in current workspace."""
    index: Dict[str, str] = {}
    workspace_root = Path.cwd()

    for path in workspace_root.rglob("*.docx"):
        key = path.name.lower()
        # Keep the first match deterministically; enough for fallback by basename.
        if key not in index:
            index[key] = str(path)

    return index


def resolve_accessible_docx_path(sd_path: str, workspace_docx_index: Dict[str, str]) -> Tuple[str, str]:
    """Resolve an accessible path for the SD. Falls back to workspace copy by basename."""
    if path_exists(sd_path):
        return sd_path, "source"

    fallback = workspace_docx_index.get(Path(sd_path).name.lower())
    if fallback and path_exists(fallback):
        return fallback, "workspace-fallback"

    return sd_path, "missing"


def load_mapping() -> Dict:
    """Load field mapping rules"""
    with open(MAPPING_PATH, 'r') as f:
        return yaml.safe_load(f)


def find_sd_files(root: str) -> List[str]:
    """Find all DOCX SD files"""
    files = []
    for dirpath, _, filenames in os.walk(root):
        for filename in filenames:
            if filename.lower().startswith("sd") and filename.lower().endswith(".docx"):
                files.append(os.path.join(dirpath, filename))
    return sorted(files)


def extract_chapters_from_docx(docx_path: str) -> Tuple[List[Dict[str, str]], Optional[str]]:
    """Extract chapters from DOCX SD file."""
    try:
        doc = Document(normalize_windows_path(docx_path))
        chapters = []
        current_title = None
        current_lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            style_name = para.style.name if para.style else None

            is_heading = style_name and style_name.lower() in {"heading 1", "heading 2", "heading 3"}

            if is_heading:
                if current_title:
                    chapters.append({
                        "title": current_title,
                        "text": "\n".join(current_lines).strip()
                    })
                current_title = text if text else "[UNTITLED]"
                current_lines = []
            else:
                if current_title and text:
                    current_lines.append(text)

        if current_title:
            chapters.append({
                "title": current_title,
                "text": "\n".join(current_lines).strip()
            })

        return chapters, None
    except Exception as e:
        logging.warning(f"Failed to extract chapters from {docx_path}: {e}")
        return [], str(e)


def map_chapters_to_template(chapters: List[Dict], mapping: Dict) -> Dict[str, bool]:
    """Map SD chapters to template sections"""
    section_coverage = {section: False for section in REQUIRED_TEMPLATE_SECTIONS}
    mapping_rules = mapping.get('mappings', {})

    chapter_blobs = [
        normalize_text(f"{c.get('title', '')} {c.get('text', '')[:1200]}")
        for c in chapters
    ]

    # 1) Detect coverage using explicit section keyword phrases.
    for template_section, keywords in SECTION_KEYWORDS.items():
        normalized_keywords = [normalize_text(k) for k in keywords]
        for blob in chapter_blobs:
            if any(k and k in blob for k in normalized_keywords):
                section_coverage[template_section] = True
                break

    # 2) Apply field_mapping.yaml rules onto mapped template sections.
    for rule_section, keywords in mapping_rules.items():
        target_sections = RULE_TO_TEMPLATE_SECTIONS.get(rule_section, [])
        if not target_sections:
            continue

        normalized_keywords = [normalize_text(k) for k in keywords]
        matched = any(any(k and k in blob for k in normalized_keywords) for blob in chapter_blobs)
        if matched:
            for section in target_sections:
                section_coverage[section] = True

    return section_coverage


def generate_reasoning(section_coverage: Dict[str, bool], chapters: List[Dict]) -> str:
    """Generate explanation for why usefulness is low"""
    missing_sections = [s for s, covered in section_coverage.items() if not covered]
    covered_sections = [s for s, covered in section_coverage.items() if covered]
    
    chapter_titles = [c.get('title', '').lower()[:50] for c in chapters]
    
    reasoning = []
    
    if len(covered_sections) == 0:
        reasoning.append("No presales template sections detected in SD content.")
    else:
        reasoning.append(f"Covers {len(covered_sections)}/15 sections: {', '.join(covered_sections[:3])}{'...' if len(covered_sections) > 3 else ''}")
    
    # Identify specific gaps
    gap_reasons = []
    
    # Check for pricing/commercial content
    has_pricing = any('pricing' in ct or 'billing' in ct or 'cost' in ct for ct in chapter_titles)
    if 'Pricing Elements' in missing_sections and not has_pricing:
        gap_reasons.append("No pricing/billing information found")
    
    # Check for customer value/positioning
    has_value_prop = any('value' in ct or 'benefit' in ct or 'advantage' in ct for ct in chapter_titles)
    if 'Value Proposition' in missing_sections and not has_value_prop:
        gap_reasons.append("No customer value or business benefit messaging")
    
    # Check for differentiators
    has_differentiators = any('compete' in ct or 'differenti' in ct or 'unique' in ct or 'advantage' in ct for ct in chapter_titles)
    if 'Key Differentiators' in missing_sections and not has_differentiators:
        gap_reasons.append("No competitive differentiators or unique selling points")
    
    # Check for customer/client perspective
    has_client_focus = any('client' in ct or 'customer' in ct or 'stakeholder' in ct or 'need' in ct for ct in chapter_titles)
    if 'Understanding the Client\'s Needs' in missing_sections and not has_client_focus:
        gap_reasons.append("SD focused on supplier perspective, not customer needs")
    
    # Check for SLA/KPI content
    has_sla = any('sla' in ct or 'kpi' in ct or 'metric' in ct or 'target' in ct for ct in chapter_titles)
    if 'SLA & KPI Management' in missing_sections and not has_sla:
        gap_reasons.append("Lacks SLA/KPI targets and performance metrics")
    
    # Check for architectural content
    has_architecture = any('architecture' in ct or 'design' in ct or 'technical' in ct or 'platform' in ct for ct in chapter_titles)
    if 'Architectural Description' in missing_sections and not has_architecture:
        gap_reasons.append("No technical architecture or system design details")
    
    # Check for terms/conditions
    has_terms = any('condition' in ct or 'term' in ct or 'responsibil' in ct or 'raci' in ct for ct in chapter_titles)
    if 'Terms & Conditions' in missing_sections and not has_terms:
        gap_reasons.append("Missing terms, conditions, and responsibilities")
    
    if gap_reasons:
        reasoning.extend(gap_reasons)
    else:
        reasoning.append(f"SD content does not align with presales template structure.")
    
    return " | ".join(reasoning)


def calculate_usefulness(section_coverage: Dict[str, bool]) -> Tuple[float, int]:
    """Calculate usefulness percentage and missing chapters"""
    covered = sum(1 for v in section_coverage.values() if v)
    total = len(section_coverage)
    usefulness_percent = (covered / total * 100) if total > 0 else 0
    missing_chapters = total - covered
    return usefulness_percent, missing_chapters


def get_missing_sections_list(section_coverage: Dict[str, bool]) -> str:
    """Get comma-separated list of missing sections"""
    missing = [s for s, covered in section_coverage.items() if not covered]
    if len(missing) > 8:
        return ", ".join(missing[:8]) + f", +{len(missing)-8} more"
    return ", ".join(missing) if missing else "All covered"


def analyze_all_sds() -> List[Dict]:
    """Analyze all SD files"""
    logging.info("Starting usefulness analysis...")
    
    sd_files = find_sd_files(ROOT_FOLDER)
    logging.info(f"Found {len(sd_files)} SD files")
    workspace_docx_index = build_workspace_docx_index()
    logging.info(f"Workspace fallback DOCX files indexed: {len(workspace_docx_index)}")

    mapping = load_mapping()
    results = []

    for idx, sd_path in enumerate(sd_files, start=1):
        sd_name = os.path.basename(sd_path)
        logging.info(f"[{idx}/{len(sd_files)}] Analyzing: {sd_name}")

        resolved_path, source_mode = resolve_accessible_docx_path(sd_path, workspace_docx_index)

        if source_mode == "missing":
            chapters = []
            extract_error = "Source file is not locally accessible (likely cloud placeholder or missing file)."
        else:
            chapters, extract_error = extract_chapters_from_docx(resolved_path)

        if extract_error:
            section_coverage = {section: False for section in REQUIRED_TEMPLATE_SECTIONS}
            usefulness_percent = None
            missing_chapters = None
            missing_list = "N/A"
            reasoning = f"Not analyzed: {extract_error}"
            source_status = f"Unavailable ({source_mode})"
        else:
            section_coverage = map_chapters_to_template(chapters, mapping)
            usefulness_percent, missing_chapters = calculate_usefulness(section_coverage)
            reasoning = generate_reasoning(section_coverage, chapters)
            missing_list = get_missing_sections_list(section_coverage)
            source_status = "Analyzed from source" if source_mode == "source" else "Analyzed from workspace fallback"

        results.append({
            "SD Name": sd_name,
            "Full Path": sd_path,
            "Analyzed Path": resolved_path,
            "Source Status": source_status,
            "Chapter Count": len(chapters),
            "Usefulness %": round(usefulness_percent, 2) if usefulness_percent is not None else None,
            "Missing Chapters": missing_chapters,
            "Covered Sections": sum(1 for v in section_coverage.values() if v),
            "Total Sections": len(section_coverage),
            "Missing Sections List": missing_list,
            "Reasoning": reasoning,
        })

    return results


def main():
    logging.info("=" * 80)
    logging.info("SD Presales Usefulness Analyzer (Enhanced with Reasoning)")
    logging.info("=" * 80)

    results = analyze_all_sds()

    # Create DataFrame
    df = pd.DataFrame(results)

    # Sort by usefulness descending; unavailable files are kept at bottom.
    df = df.sort_values("Usefulness %", ascending=False, na_position="last")

    # Save to Excel
    output_target = OUTPUT_FILE
    try:
        df.to_excel(output_target, index=False, engine="openpyxl")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_target = f"sd_presales_usefulness_{timestamp}.xlsx"
        logging.warning(
            f"Output file '{OUTPUT_FILE}' is locked. Writing to '{output_target}' instead."
        )
        df.to_excel(output_target, index=False, engine="openpyxl")

    logging.info(f"✓ Saved {len(df)} results to {output_target}")

    # Summary
    analyzed_df = df[df["Usefulness %"].notna()]
    unavailable_count = len(df) - len(analyzed_df)

    if len(analyzed_df) > 0:
        avg_usefulness = analyzed_df["Usefulness %"].mean()
        max_usefulness = analyzed_df["Usefulness %"].max()
        min_usefulness = analyzed_df["Usefulness %"].min()

        logging.info(f"\n--- Summary ---")
        logging.info(f"Analyzed SDs: {len(analyzed_df)}")
        logging.info(f"Unavailable SDs: {unavailable_count}")
        logging.info(f"Average usefulness: {avg_usefulness:.2f}%")
        logging.info(f"Max usefulness: {max_usefulness:.2f}%")
        logging.info(f"Min usefulness: {min_usefulness:.2f}%")
        logging.info(f"Average missing chapters: {analyzed_df['Missing Chapters'].mean():.1f}")

        # Show top performers with reasoning
        logging.info(f"\n--- Top 5 Most Useful SDs ---")
        for _, row in analyzed_df.head(5).iterrows():
            logging.info(f"{row['Usefulness %']:.1f}% | {row['SD Name'][:60]}")
            logging.info(f"  └─ {row['Reasoning'][:140]}")

        logging.info(f"\n--- Bottom 5 Least Useful SDs ---")
        for _, row in analyzed_df.tail(5).iterrows():
            logging.info(f"{row['Usefulness %']:.1f}% | {row['SD Name'][:60]}")
            logging.info(f"  └─ {row['Reasoning'][:140]}")
    else:
        logging.info(f"\n--- Summary ---")
        logging.info("No SD files could be analyzed (all unavailable).")
        logging.info(f"Unavailable SDs: {unavailable_count}")


if __name__ == "__main__":
    main()
