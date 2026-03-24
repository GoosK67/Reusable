from pathlib import Path
import sys
import json
from datetime import datetime

# ---------------------------------------------------------
# LOGGING (always append, same pattern as extract + parse)
# ---------------------------------------------------------
LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)

def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"

    with open(logfile, "a", encoding="utf-8", errors="ignore") as f:
        f.write(line)

    print(line, end="")

# ---------------------------------------------------------
# Deterministic classifier aligned 1:1 with presales template
# ---------------------------------------------------------
def classify_section(header: str, content: str, table_facts=None) -> dict:
    """
    Very simple deterministic classifier.
    Replace later with Azure OpenAI / local model.
    """
    header_lower = header.lower().strip()
    table_facts = table_facts or []
    table_text = " ".join(
        " ".join([
            str(item.get("fact_type", "")),
            str(item.get("row_text", "")),
            " ".join(f"{k} {v}" for k, v in (item.get("facts", {}) or {}).items()),
        ])
        for item in table_facts
        if isinstance(item, dict)
    )
    content_lower = f"{content} {table_text}".lower().strip()

    def has_any(terms):
        return any(t in header_lower or t in content_lower for t in terms)

    # 1. Product Summary
    if has_any(["service introduction", "service identification", "service reporting", "service window", "product summary"]):
        category = "PRODUCT_SUMMARY"

    # 2. Understanding the Client's Needs
    elif has_any(["service overview", "goals", "target audience", "client needs", "customer needs"]):
        category = "CLIENT_NEEDS"

    # 3. Product Description
    elif has_any(["product description", "service description", "standard services", "optional services"]):
        category = "PRODUCT_DESCRIPTION"

    # 3.1 Architectural Description
    elif has_any(["architecture", "technical implementation", "technical architecture", "architectural description"]):
        category = "ARCHITECTURAL_DESCRIPTION"

    # 3.2 Key Features & Functionalities
    elif has_any(["key features", "functionalities", "operational readiness", "run services", "management services", "governance", "process"]):
        category = "KEY_FEATURES"

    # 3.3 Scope / Out-of-Scope
    elif has_any(["scope", "out of scope", "in scope"]):
        category = "SCOPE"

    # 3.4 Requirements & Prerequisites
    elif has_any(["requirement", "prerequisite", "eligibility", "dependency"]):
        category = "REQUIREMENTS"

    # 4. Value Proposition
    elif has_any(["value proposition", "value and benefits", "value & benefits", "benefit", "business value"]):
        category = "VALUE_PROPOSITION"

    # 5. Key Differentiators
    elif has_any(["differentiator", "unique", "strength"]):
        category = "DIFFERENTIATORS"

    # 6. Transition & Transformation
    elif has_any(["transition", "transformation", "onboarding", "migration"]):
        category = "TRANSITION_TRANSFORMATION"

    # 7. Client Responsibilities
    elif has_any(["responsibilit", "provided by customer", "customer provides", "client provides"]):
        category = "CLIENT_RESPONSIBILITIES"

    # 8. Operational Support
    elif has_any(["support", "incident", "problem", "service request", "operational"]):
        category = "OPERATIONAL_SUPPORT"

    # 9. Terms & Conditions
    elif has_any(["terms", "conditions", "limitations", "contract"]):
        category = "TERMS_CONDITIONS"

    # 10. SLA & KPI
    elif has_any(["sla", "kpi", "service level", "availability", "response time"]):
        category = "SLA_KPI"

    # 11. Pricing Elements
    elif has_any(["pricing", "billing", "price", "delivery model", "cost", "charge"]):
        category = "PRICING_ELEMENTS"

    else:
        category = "UNMAPPED"

    return {
        "header": header,
        "category": category,
        "content": content.strip(),
        "table_facts": table_facts,
    }

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    json_file = Path(sys.argv[1])
    sd_name = json_file.stem

    log(f"START auto_map for: {json_file}", sd_name)

    try:
        data = json.loads(
            json_file.read_text(encoding="utf-8", errors="ignore")
        )

        mapped = {}

        for header, payload in data.items():
            if isinstance(payload, dict):
                content = str(payload.get("content", "") or "")
                table_facts = payload.get("table_facts", []) or []
            else:
                content = str(payload or "")
                table_facts = []

            mapped_entry = classify_section(header, content, table_facts)
            mapped[header] = mapped_entry

        # Save result
        out_folder = Path("output/mapped")
        out_folder.mkdir(parents=True, exist_ok=True)

        out_file = out_folder / f"{sd_name}_mapped.json"
        out_file.write_text(
            json.dumps(mapped, indent=2, ensure_ascii=False),
            encoding="utf-8",
            errors="ignore"
        )

        log(f"auto_map OK → {out_file}", sd_name)
        sys.exit(0)

    except Exception as e:
        log(f"auto_map ERROR: {e}", sd_name)
        sys.exit(1)