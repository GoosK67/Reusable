"""
End-to-end AI mapping pipeline for Cegeka Service Descriptions.

Entry point: run_pipeline(sd_input)

Accepts:
    - A raw SD text block (str with newlines)
    - A path to a DOCX file (str | Path)

Steps:
    1. EXTRACT   — parse numbered sections from raw text or DOCX
    2. DETECT    — keyword + fuzzy match every title against template fields
    3. MAP       — assign the best-matching section text to each template field
    4. REWRITE   — apply enterprise commercial rewrite profile to Product Description
    5. SUMMARIZE — trim each field to max 8 lines via summarize() placeholder
    6. OUTPUT    — return a fully populated dict (+ optional DOCX generation)

Available rewrite profiles:
    enterprise_strict | enterprise_balanced | enterprise_concise
"""

from __future__ import annotations

from pathlib import Path

from extractor import extract_sections, extract_sections_from_text
from generator import fill_template
from mapper import DEFAULT_REWRITE_PROFILE, get_match_diagnostics, map_sd_to_template

# Maps human-readable template field names to DOCX placeholder tags used by fill_template.
TEMPLATE_FIELD_TO_TAG: dict[str, str] = {
    "Product Summary":              "ProductSummary",
    "Value Proposition":            "ValueProposition",
    "Product Description":          "ProductDescription",
    "Requirements & Prerequisites": "Requirements",
    "Scope / Out of Scope":         "Scope",
    "Key Differentiators":          "KeyDifferentiators",
    "Operational Support":          "OperationalSupport",
    "SLA":                          "SLA",
}


def run_pipeline(
    sd_input,
    template_path=None,
    output_path=None,
    rewrite_profile: str = DEFAULT_REWRITE_PROFILE,
    verbose: bool = True,
) -> dict:
    """
    End-to-end SD → Presales Guide mapping pipeline.

    Args:
        sd_input:         Raw SD text (str) or path to a DOCX file (str | Path).
        template_path:    Optional DOCX template path for DOCX generation.
        output_path:      Optional output DOCX path.
        rewrite_profile:  Tone profile for Product Description rewrite.
        verbose:          Print step-by-step diagnostics to stdout.

    Returns:
        dict with one key per template field (filled text) plus '_diagnostics'.
    """
    # ── 1. EXTRACT ────────────────────────────────────────────────────────────
    is_file = isinstance(sd_input, Path) or (
        isinstance(sd_input, str) and "\n" not in sd_input and Path(sd_input).exists()
    )

    if is_file:
        sections = extract_sections(Path(sd_input))
        source_label = str(sd_input)
    else:
        sections = extract_sections_from_text(str(sd_input))
        source_label = "raw text input"

    if verbose:
        print(f"[1/5] EXTRACT  — {len(sections)} section(s) from {source_label}")
        for title in sections:
            print(f"        • {title}")

    # ── 2. DETECT ─────────────────────────────────────────────────────────────
    diagnostics = get_match_diagnostics(sections)

    if verbose:
        col = max(len(f) for f in diagnostics) + 2
        print(f"\n[2/5] DETECT   — keyword + fuzzy match per template field")
        for field, info in diagnostics.items():
            if info["matched_title"]:
                print(
                    f"        {field:<{col}} <- '{info['matched_title']}'  "
                    f"(score {info['score']:.2f})"
                )
            else:
                print(
                    f"        {field:<{col}} <- NO MATCH  "
                    f"(best score {info['score']:.2f})"
                )

    # ── 3 + 4 + 5. MAP → REWRITE → SUMMARIZE ─────────────────────────────────
    mapped = map_sd_to_template(sections, rewrite_profile=rewrite_profile)

    if verbose:
        col = max(len(f) for f in mapped) + 2
        print(f"\n[3-5/5] MAP + REWRITE + SUMMARIZE  (profile: {rewrite_profile})")
        for field, text in mapped.items():
            flat = text.replace("\n", " ")
            preview = flat[:90].rstrip()
            ellipsis = "…" if len(flat) > 90 else ""
            print(f"        {field:<{col}}: {preview}{ellipsis}")

    # ── 6. BUILD OUTPUT DICT ──────────────────────────────────────────────────
    output: dict = {field: mapped.get(field, "") for field in TEMPLATE_FIELD_TO_TAG}
    output["_diagnostics"] = diagnostics

    # ── OPTIONAL: GENERATE DOCX ───────────────────────────────────────────────
    if template_path and output_path:
        fill_fields = {
            tag: output.get(field, "")
            for field, tag in TEMPLATE_FIELD_TO_TAG.items()
        }
        fill_template(Path(template_path), Path(output_path), fill_fields)
        if verbose:
            print(f"\n[OUTPUT] Presales Guide saved to: {output_path}")

    return output


# ── SD Demo fragment (from a real Cegeka Digital Wellbeing Service Description) ──────────
SD_DEMO = """\
1 Standard Service Description

1.1 Introduction
Our Digital Wellbeing service is designed to enhance the employee experience within
your organization by leveraging data-driven insights and actionable recommendations.
By implementing this service, you can drive business success, increase employee
engagement, enhance productivity, and improve overall digital wellbeing.
Through the use of Insights, we provide managers and leaders with visibility into team
dynamics, enabling them to foster a sense of belonging, encourage effective
communication, and create a positive work environment.
With our service we analyse work patterns and offer personalized recommendations for
time management, focused work, and learning opportunities.
In addition to the overall benefits, we focus on three key pillars:
Individual Digital Wellbeing: leveraging data-driven insights to provide personalized
recommendations that help employees manage workloads and maintain work-life balance.
Hybrid Teamwork and Remote Habits: providing managers with visibility into team
dynamics to cultivate belonging and effective communication regardless of location.
Leadership in a Digital Age: equipping leaders with tools and insights to become
digitally savvy and empathetic, driving optimal performance and wellbeing.

1.2 What is the added business value?
Our Digital Wellbeing service offers significant added business value by aligning with
our three key focus areas, contributing to the holistic improvement of the employee
experience within your organization.

1.2.1 Enhanced employee productivity
By nurturing digital wellbeing, employees experience an increase in productivity,
a more harmonious work-life balance, and a decrease in burnout. Insights provide
valuable data on work patterns and collaboration trends, enabling targeted guidance.

1.2.2 Increased Employee Wellbeing and Engagement
Prioritizing digital wellbeing demonstrates commitment to employee mental health.
When employees feel supported, they are more likely to be engaged and motivated.
Personalized recommendations help establish healthy work habits and reduce turnover.

2.3 Product Description (MVP)
Commercial description of a standardized Microsoft Viva Insights deployment.

2.3.1 Architectural description
A standardized deployment of Microsoft Viva Insights including configuration,
governance setup, and report templates. Integration with Microsoft 365 tenant
services and Active Directory for org hierarchy and manager role resolution.

2.3.2 Key features and functionalities
Data-driven insights, personalized recommendations, team dynamics visibility,
leadership dashboards, and wellbeing trend analysis across the organization.

2.3.3 Scope out-of-scope
In scope: Viva Insights configuration, report templates, onboarding sessions,
governance documentation, and manager enablement.
Out of scope: custom development, HR system integration, psychosocial risk
assessments, and continuous managed service beyond initial rollout.

2.3.4 Requirements and Prerequisites
Microsoft 365 E3 or higher license. Entra ID tenant in good standing.
Manager roles and org hierarchy correctly configured in Active Directory.
Privacy review and works council approval completed prior to deployment.

2.4 Value Proposition
Digital Wellbeing as a Service enables organizations to systematically improve
employee engagement, reduce digital overload, and build healthier work habits at scale.
The service provides measurable impact through structured onboarding, continuous
improvement cycles, and leadership enablement anchored to real usage data.
"""


if __name__ == "__main__":
    result = run_pipeline(SD_DEMO, verbose=True)

    width = 62
    print("\n" + "=" * width)
    print("  FINAL OUTPUT DICT")
    print("=" * width)
    for field, value in result.items():
        if field.startswith("_"):
            continue
        print(f"\n  ┌─ {field}")
        for line in (value or "(empty)").splitlines():
            print(f"  │  {line}")
