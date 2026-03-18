import json
from pathlib import Path

JSON_FOLDER = Path("presales_json")
TEMPLATE_FILE = Path("templates/presales_template.md")
OUTPUT_FOLDER = Path("output")
OUTPUT_FOLDER.mkdir(exist_ok=True)

def make_bullet_list(items):
    if not items:
        return "[TO BE COMPLETED]"
    return "\n".join(f"- {item.strip()}" for item in items if item.strip())

def safe(value):
    return value if value not in [None, "", []] else "[TO BE COMPLETED]"

def main():
    template = TEMPLATE_FILE.read_text(encoding="utf-8")

    for js in JSON_FOLDER.glob("*.json"):
        print(f"Building presales guide from {js.name}")

        data = json.loads(js.read_text(encoding="utf-8"))

        # Start from template
        guide = template

        # Replace scalar placeholders
        scalar_fields = [
            "title", "service_summary", "prerequisites",
            "out_of_scope", "conditions", "sla", "pricing"
        ]

        for field in scalar_fields:
            placeholder = f"<{field.upper()}>"
            guide = guide.replace(placeholder, safe(data.get(field)))

        # Replace list placeholders
        list_fields = [
            "key_features",
            "standard_services",
            "optional_services",
            "operational_services",
            "risks",
            "assumptions",
            "differentiators"
        ]

        for field in list_fields:
            placeholder = f"<{field.upper()}>"
            guide = guide.replace(placeholder, make_bullet_list(data.get(field, [])))

        # write result
        out = OUTPUT_FOLDER / f"{js.stem}_presales.md"
        out.write_text(guide, encoding="utf-8")

        print(f"✓ Presales guide → {out}")

if __name__ == "__main__":
    main()