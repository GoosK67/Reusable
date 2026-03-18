import json
import re
from pathlib import Path
import yaml

RAW_FOLDER = Path("extracted")
OUTPUT_FOLDER = Path("extracted/json_auto")
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

MAPPING_FILE = Path("rules/field_mapping.yaml")


BASE_JSON = {
    "title": None,
    "service_summary": None,
    "business_context": None,
    "key_features": [],
    "standard_services": [],
    "optional_services": [],
    "operational_services": [],
    "prerequisites": None,
    "out_of_scope": None,
    "conditions": None,
    "sla": None,
    "pricing": None,
    "risks": [],
    "assumptions": [],
    "differentiators": [],
    "missing_information": []
}


def load_mapping():
    with open(MAPPING_FILE, encoding="utf-8") as f:
        return yaml.safe_load(f)["mappings"]


def detect_markdown_headings(text_lines):
    headings = []
    for i, line in enumerate(text_lines):
        if re.match(r"^#{3,5}\s+", line.strip()):  # ###, ####, #####
            clean = re.sub(r"^#+\s*", "", line).strip()
            headings.append((i, clean))
    return headings


def segment_text(text):
    lines = text.split("\n")
    headings = detect_markdown_headings(lines)

    if not headings:
        return {"Uncategorized": text}

    segments = {}
    for idx, (line_no, title) in enumerate(headings):
        start = line_no + 1
        end = headings[idx + 1][0] if idx + 1 < len(headings) else len(lines)
        content = "\n".join(lines[start:end]).strip()
        segments[title] = content

    return segments


def normalize_title(t):
    return t.lower().replace(":", "").strip()


def map_segments_to_json(segments, rules):
    data = {k: ([] if isinstance(v, list) else v) for k, v in BASE_JSON.items()}

    for title, content in segments.items():
        norm = normalize_title(title)

        for field, keywords in rules.items():
            for kw in keywords:
                if kw.lower() in norm:
                    if isinstance(data[field], list):
                        data[field].append(content)
                    else:
                        data[field] = content
                    break

    return data


def main():
    mapping_rules = load_mapping()

    for raw in RAW_FOLDER.glob("*.raw.txt"):
        text = raw.read_text(encoding="utf-8", errors="ignore")

        sections = segment_text(text)
        json_data = map_segments_to_json(sections, mapping_rules)

        out = OUTPUT_FOLDER / f"{raw.stem}.json"
        out.write_text(
            json.dumps(json_data, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

        print(f"✓ Extracted JSON → {out}")


if __name__ == "__main__":
    main()