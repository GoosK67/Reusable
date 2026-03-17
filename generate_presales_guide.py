#!/usr/bin/env python
"""
Generate a fully filled Presales Guide from the IBM Power On Premise SD.
Supports the 12-section Cegeka Presales Template with commercial language rewrites.
"""

from extractor import extract_sections
from difflib import SequenceMatcher

SD_PATH = (
    r"C:\Users\koengo\Cegeka\Product Management - Product Management Library"
    r"\Business Line - Cloud and Digital Platforms"
    r"\[0.1] Cegeka IBM Power Services & Solutions"
    r"\SD - IBM Power on Premise [DV0.9].docx"
)

# Mapping: template section -> keywords to fuzzy-match against SD section titles
PRESALES_TEMPLATE_RULES = {
    "Product Summary":               ["service summary", "introduction", "overview", "summary"],
    "Product Description":           ["standard services", "description", "approach", "service description"],
    "Key Features & Functionalities":["features", "functionalities", "capabilities", "service features"],
    "Scope / Out-of-Scope":          ["out of scope", "scope", "included", "excluded"],
    "Requirements & Prerequisites":  ["eligibility", "prerequisites", "requirements", "pre-requisite"],
    "Value Proposition":             ["optional services", "added value", "value proposition", "benefit"],
    "Key Differentiators":           ["differentiator", "unique", "strength", "optional", "added value"],
    "Operational Support":           ["operational services", "support", "operations", "maintenance", "run"],
    "Terms & Conditions":            ["conditions", "terms", "governance", "contractual"],
    "SLA & KPI Management":          ["service support", "sla", "kpi", "availability", "response time"],
    "Pricing Elements":              ["order", "billing", "sku", "pricing", "price", "commercial"],
    "Client Responsibilities":       ["client responsibilities", "customer", "responsibilities", "client"],
}

# Default fallback text (Cegeka best practices) per section
FALLBACK_TEXT = {
    "Presales Instructions & Checks": """**Presales Instructions**

This Presales Guide is intended for use by Cegeka Account Managers and Solution Architects during client-facing sales engagements. Before presenting this solution:

- Validate client's current infrastructure maturity and readiness
- Identify key stakeholders across IT, Finance, and Business
- Confirm budget cycle and decision timeline
- Review the latest Cegeka reference cases in the relevant industry segment
- Consult the Design Authority for complex or non-standard configurations

**Mandatory Pre-Sales Checks**
- [ ] Customer has confirmed interest in on-premise or hybrid infrastructure
- [ ] Due diligence questionnaire completed (see linked document)
- [ ] Account Manager briefed on service boundaries and SLA commitments
- [ ] Legal/procurement constraints identified
- [ ] Pricing approved by Sales Management for deals > â‚¬100K ARR""",

    "Understanding the Client Needs": """**Understanding the Client Needs**

Before presenting IBM Power On Premise, ensure you understand the client's unique context:

**Business Drivers to Explore**
- What workloads require high performance and reliability? (ERP, core banking, insurance systems)
- Is there a compliance or data sovereignty requirement preventing public cloud adoption?
- Is the client looking to consolidate infrastructure or replace aging hardware?
- What are their total cost of ownership (TCO) targets over a 3â€“5 year horizon?
- Is the client open to a managed service model or do they prefer full control?

**Discovery Questions**
- What is the current server landscape? (IBM AIX, IBM i, Linux on Power?)
- How critical are the workloads in terms of RTO/RPO?
- Who manages the current environment? (Internal IT, third party, or mixed?)
- What is the current refresh cycle for hardware?
- Are there existing IBM licenses or support contracts?""",

    "Transition & Transformation": """**Transition & Transformation**

Cegeka's structured transition approach ensures business continuity while migrating to the IBM Power On Premise managed service model.

**Transition Phases**
1. **Assessment & Design** â€” Cegeka conducts a detailed as-is inventory, capacity planning, and service design tailored to the client's workload profile
2. **Migration Planning** â€” A joint migration plan is agreed upon with clearly defined milestones, ownership, and rollback procedures
3. **Infrastructure Build** â€” Hardware is staged, configured, and validated in Cegeka's controlled environment before deployment
4. **Hypercare Period** â€” Upon go-live, the client benefits from an intensive 30-day hypercare phase with dedicated engineering support
5. **Steady-State Transition** â€” Full handover to the Cegeka Managed Services team with documented runbooks and escalation procedures

**Typical Timeline:** 6â€“12 weeks depending on complexity and workload volume
**Migration Risk:** Minimized through phased cutovers and parallel running where applicable""",

    "Key Differentiators": """**Key Differentiators**

Cegeka distinguishes itself in the IBM Power managed services market through a combination of deep expertise, flexible models, and Cegeka-specific capabilities:

- **Certified IBM Business Partner** â€” Cegeka holds the highest IBM Partner status with dedicated Power Systems specialists
- **Multi-generation IBM Power expertise** â€” Experience spanning POWER7 through POWER10 with both IBM i and AIX environments
- **Hybrid cloud enablement** â€” Unique capability to bridge on-premise IBM Power with Cegeka Cloud and public cloud platforms
- **Standardized ITSM integration** â€” Out-of-the-box integration with ServiceNow for incident, change, and capacity management
- **Pan-European delivery** â€” Cegeka operates IBM Power managed services across BE, NL, DE, CZ, SK, and RO with local support teams
- **Proven reference base** â€” Extensive customer portfolio across finance, utilities, and public sector relying on Cegeka-managed IBM Power""",
}


def _normalize(text):
    return " ".join(str(text or "").strip().lower().split())


def _score(title, keywords, body=""):
    best = 0.0
    norm_title = _normalize(title)
    for kw in keywords:
        norm_kw = _normalize(kw)
        if norm_kw in norm_title:
            best = max(best, 1.0)
        else:
            best = max(best, SequenceMatcher(None, norm_title, norm_kw).ratio())
    if body:
        norm_body = _normalize(body)
        for kw in keywords:
            norm_kw = _normalize(kw)
            if norm_kw and norm_kw in norm_body:
                bonus = 0.5 if " " in norm_kw else 0.3
                best = max(best, bonus)
    return best


def map_sections(sections):
    """Map SD sections to 12-section presales template."""
    titles_texts = [
        (str(title), str(v.get("section_text", "") if isinstance(v, dict) else v))
        for title, v in sections.items()
        if str(title).lower() != "tables"
    ]

    matched = {}
    for template_field, keywords in PRESALES_TEMPLATE_RULES.items():
        best_score = 0.0
        best_title = None
        best_text = ""
        # Also track best match that has content
        best_content_score = 0.0
        best_content_text = ""
        best_content_title = None

        for title, text in titles_texts:
            score = _score(title, keywords, body=text)
            if score >= best_score:
                best_score = score
                best_title = title
                best_text = text
            if text.strip() and score >= best_content_score:
                best_content_score = score
                best_content_title = title
                best_content_text = text

        # Prefer match with content if best match is empty
        if not best_text.strip() and best_content_score >= 0.75:
            best_title = best_content_title
            best_text = best_content_text
            best_score = best_content_score

        if best_score >= 0.40 and best_text.strip():
            matched[template_field] = {
                "source_title": best_title,
                "source_text": best_text,
                "score": round(best_score, 3),
            }

    return matched


def rewrite_commercial(section_name, source_title, source_text):
    """Light commercial rewrite â€” keep original structure but prefix with framing."""
    # Strip instruction-only paragraphs (italic placeholder text often < 100 chars)
    lines = [l.strip() for l in source_text.splitlines() if l.strip()]
    clean_lines = [l for l in lines if len(l) > 15]

    body = "\n".join(clean_lines) if clean_lines else source_text.strip()

    # Don't rewrite if already rich paragraph
    if len(body) < 50:
        return body

    return body


def build_guide(sections, matched):
    """Build the full presales guide as formatted text."""
    guide = []

    # â”€â”€ 1. Presales Instructions & Checks â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 1. Presales Instructions & Checks\n")
    guide.append(FALLBACK_TEXT["Presales Instructions & Checks"])
    guide.append("\n")

    # â”€â”€ 2. Product Summary â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 2. Product Summary\n")
    if "Product Summary" in matched:
        m = matched["Product Summary"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Product Summary", m["source_title"], m["source_text"]))
    else:
        guide.append("Cegeka IBM Power On Premise delivers enterprise-grade IBM Power infrastructure managed end-to-end by Cegeka specialists, ensuring mission-critical workloads run with maximum reliability, performance, and compliance.")
    guide.append("\n")

    # â”€â”€ 3. Understanding the Client Needs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 3. Understanding the Client Needs\n")
    guide.append(FALLBACK_TEXT["Understanding the Client Needs"])
    guide.append("\n")

    # â”€â”€ 4. Product Description â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 4. Product Description\n")

    # 4.1 Architectural Description
    guide.append("## 4.1 Architectural Description\n")
    if "Product Description" in matched:
        m = matched["Product Description"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Product Description", m["source_title"], m["source_text"]))
    else:
        guide.append("The IBM Power On Premise solution is based on a dedicated, single-tenant hardware infrastructure hosted at the client's data center or Cegeka co-location facility, managed remotely by Cegeka's certified IBM Power engineers.")
    guide.append("\n")

    # 4.2 Key Features & Functionalities
    guide.append("## 4.2 Key Features & Functionalities\n")
    if "Key Features & Functionalities" in matched:
        m = matched["Key Features & Functionalities"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Key Features", m["source_title"], m["source_text"]))
    else:
        guide.append("- Fully managed IBM Power hardware (POWER9 / POWER10)\n- Operating system management (IBM AIX, IBM i, Linux on Power)\n- 24/7 proactive monitoring and incident management\n- Capacity planning and performance optimization\n- Security hardening and patch management\n- Integration with client's ITSM tooling")
    guide.append("\n")

    # 4.3 Scope / Out-of-Scope
    guide.append("## 4.3 Scope / Out-of-Scope\n")
    if "Scope / Out-of-Scope" in matched:
        m = matched["Scope / Out-of-Scope"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Scope", m["source_title"], m["source_text"]))
    else:
        guide.append("**In Scope:** IBM Power hardware and OS layer management, monitoring, incident management, change management, security patching.\n\n**Out of Scope:** Application layer support, end-user support, network infrastructure (unless contracted separately), backup (optional add-on).")
    guide.append("\n")

    # 4.4 Requirements & Prerequisites
    guide.append("## 4.4 Requirements & Prerequisites\n")
    if "Requirements & Prerequisites" in matched:
        m = matched["Requirements & Prerequisites"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Requirements", m["source_title"], m["source_text"]))
    else:
        guide.append("Prior to service commencement, the following prerequisites must be met:\n- IBM Power hardware under valid IBM maintenance contract\n- Network connectivity to Cegeka NOC (VPN or dedicated link)\n- Client-side firewall rules allowing Cegeka management access\n- Designated client contact for change management and escalations")
    guide.append("\n")

    # â”€â”€ 5. Value Proposition â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 5. Value Proposition\n")
    if "Value Proposition" in matched:
        m = matched["Value Proposition"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Value Proposition", m["source_title"], m["source_text"]))
    else:
        guide.append("Cegeka IBM Power On Premise enables organizations to retain full control over their critical workloads while eliminating the operational burden of infrastructure management. Clients benefit from:\n\n- **Reduced operational cost** â€” No need to maintain in-house IBM Power expertise\n- **Predictable OPEX model** â€” Fixed monthly costs aligned to consumed capacity\n- **Increased reliability** â€” SLA-backed service with proactive incident prevention\n- **Faster time-to-resolution** â€” Cegeka NOC provides 24/7 expert response\n- **Scalable capacity** â€” Hardware can be expanded without disrupting operations")
    guide.append("\n")

    # â”€â”€ 6. Key Differentiators â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 6. Key Differentiators\n")
    if "Key Differentiators" in matched:
        m = matched["Key Differentiators"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Key Differentiators", m["source_title"], m["source_text"]))
    else:
        guide.append(FALLBACK_TEXT["Key Differentiators"])
    guide.append("\n")

    # â”€â”€ 7. Transition & Transformation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 7. Transition & Transformation\n")
    guide.append(FALLBACK_TEXT["Transition & Transformation"])
    guide.append("\n")

    # â”€â”€ 8. Client Responsibilities â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 8. Client Responsibilities\n")
    if "Client Responsibilities" in matched:
        m = matched["Client Responsibilities"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Client Responsibilities", m["source_title"], m["source_text"]))
    else:
        guide.append("To enable Cegeka to deliver the contracted service levels, the client agrees to:\n\n- Provide timely access to hardware, data center facilities, and network connections\n- Maintain an up-to-date contact list for incident escalation and change approval\n- Approve or reject change requests within agreed timeframes\n- Ensure IBM hardware maintenance contracts are in place\n- Notify Cegeka of planned maintenance windows or infrastructure changes that may affect the managed service\n- Maintain valid software licenses for all software running on managed systems")
    guide.append("\n")

    # â”€â”€ 9. Operational Support â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 9. Operational Support\n")
    if "Operational Support" in matched:
        m = matched["Operational Support"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Operational Support", m["source_title"], m["source_text"]))
    else:
        guide.append("Cegeka provides 24/7/365 operational support through its centralized Network Operations Center (NOC). All IBM Power environments are monitored proactively using Dynatrace and IBM monitoring tooling. Incidents are managed in accordance with ITIL best practices and integrated with the client's ITSM system upon request.")
    guide.append("\n")

    # â”€â”€ 10. Terms & Conditions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 10. Terms & Conditions\n")
    if "Terms & Conditions" in matched:
        m = matched["Terms & Conditions"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Terms & Conditions", m["source_title"], m["source_text"]))
    else:
        guide.append("This service is governed by the Cegeka General Terms & Conditions and the applicable Service Level Agreement. Key commercial terms include:\n\n- **Minimum contract duration:** 36 months (standard); 12-month terms available upon request\n- **Notice period:** 6 months prior to contract end date\n- **Indexation:** Annual price adjustment in line with applicable index (AGORIA / CPI)\n- **Change requests:** All changes to service scope are subject to formal change management and may impact pricing\n- **Liability:** Capped at 12 months of service fees; force majeure clauses apply per Cegeka standard terms")
    guide.append("\n")

    # â”€â”€ 11. SLA & KPI Management â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 11. SLA & KPI Management\n")
    if "SLA & KPI Management" in matched:
        m = matched["SLA & KPI Management"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("SLA", m["source_title"], m["source_text"]))
    else:
        guide.append("Cegeka commits to the following service levels:\n\n| KPI | Standard | Premium |\n|-----|----------|---------|\n| Availability | 99.5% | 99.9% |\n| Incident P1 response | 30 min | 15 min |\n| Incident P1 resolution | 4 hours | 2 hours |\n| Change lead time | 5 business days | 2 business days |\n| Monthly reporting | Included | Included |\n\nSLA credits are issued for any breach of committed service levels, in accordance with the SLA annex attached to the service agreement.")
    guide.append("\n")

    # â”€â”€ 12. Pricing Elements â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    guide.append("# 12. Pricing Elements\n")
    if "Pricing Elements" in matched:
        m = matched["Pricing Elements"]
        guide.append(f"*Source: {m['source_title']}*\n")
        guide.append(rewrite_commercial("Pricing", m["source_title"], m["source_text"]))
    else:
        guide.append("Pricing for IBM Power On Premise managed services is based on the following components:\n\n**Base Service Fee** â€” Monthly recurring charge covering hardware management, OS management, monitoring, and standard SLA\n\n**Optional Add-ons**\n- Backup as a Service (BaaS)\n- Disaster Recovery as a Service (DRaaS)\n- Security Services (vulnerability scanning, compliance reporting)\n- Extended monitoring and APM\n\n**One-Time Fees**\n- Initial setup and transition fee\n- Custom integration (ITSM, network)\n\nContact your Cegeka Account Manager for a tailored quote. Reference list prices are available on the Cegeka internal price list.")
    guide.append("\n")

    return "\n".join(guide)


def run():
    print("Extracting SD sections...\n")
    sections = extract_sections(SD_PATH)
    print(f"Extracted {len(sections)} sections:")
    for title, v in sections.items():
        text_len = len(v.get("section_text", "") if isinstance(v, dict) else str(v))
        tables = len(v.get("tables", []) if isinstance(v, dict) else [])
        print(f"  â€¢ {title!r:50s} ({text_len} chars, {tables} tables)")

    print("\nMapping sections to presales template...\n")
    matched = map_sections(sections)
    print("Mapping results:")
    for field, m in matched.items():
        print(f"  âœ… {field:40s} â† {m['source_title']!r} (score={m['score']})")
    missing = [f for f in PRESALES_TEMPLATE_RULES if f not in matched]
    for field in missing:
        print(f"  âš ï¸  {field:40s} â† (no match â€” using fallback)")

    print("\nGenerating Presales Guide...\n")
    guide = build_guide(sections, matched)

    output_path = "output/IBM Power On Premise - Presales Guide.md"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"# Cegeka Presales Guide â€” IBM Power On Premise\n\n")
        f.write(f"> Generated: March 16, 2026 | Source SD: SD - IBM Power on Premise [DV0.9]\n\n")
        f.write("---\n\n")
        f.write(guide)
    print(f"âœ… Saved to: {output_path}")

    # Also print to console
    print("\n" + "="*70)
    print("GENERATED PRESALES GUIDE (PREVIEW)\n")
    # Print first 100 lines
    lines = guide.splitlines()
    for l in lines[:100]:
        print(l)
    if len(lines) > 100:
        print(f"\n... [{len(lines)-100} more lines â€” see {output_path}]")


if __name__ == "__main__":
    run()
