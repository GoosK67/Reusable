import json
import re
from pathlib import Path

STRUCTURED = Path("structured")
OUTPUT = Path("presales_json")
OUTPUT.mkdir(exist_ok=True)

PRESALES = {
    "title": "",
    "service_summary": "",
    "key_features": [],
    "standard_services": [],
    "optional_services": [],
    "operational_services": [],
    "prerequisites": "",
    "out_of_scope": "",
    "conditions": "",
    "sla": "",
    "pricing": "",
    "risks": [],
    "assumptions": [],
    "differentiators": []
}

# CLASSIFICATION RULES
RULES = {
    "service_summary": [
        "introduction", "overview", "service description"
    ],
    "key_features": [
        "technical implementation", "operational readiness",
        "service features", "service overview"
    ],
    "standard_services": [
        "run services", "management services", "security management",
        "service management", "platform management"
    ],
    "optional_services": [
        "optional", "standard changes"
    ],
    "operational_services": [
        "governance", "process", "responsibilities", "raci"
    ],
    "prerequisites": [
        "eligibility", "prerequisite"
    ],
    "out_of_scope": [
        "out of scope"
    ],
    "conditions": [
        "conditions", "limiting conditions"
    ],
    "sla": [
        "service level", "availability", "sla", "kpi"
    ],
    "pricing": [
        "service billing", "optional service billing",
        "change request billing"
    ],
}


def classify(heading: str) -> str:
    h = heading.lower()

    for field, words in RULES.items():
        for w in words:
            if w in h:
                return field
    return "differentiators"


def main():
    for file in STRUCTURED.glob("*.sections.json"):
        sections = json.loads(file.read_text(encoding="utf-8"))
        
        result = {k: ([] if isinstance(v, list) else "") for k, v in PRESALES.items()}

        for heading, content in sections.items():
            target = classify(heading)

            if isinstance(result[target], list):
                result[target].append(content.strip())
            else:
                result[target] = content.strip()

        out = OUTPUT / file.name.replace(".sections.json", ".json")
        out.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")

        print(f"✓ Auto-mapped → {out}")


if __name__ == "__main__":
    main()