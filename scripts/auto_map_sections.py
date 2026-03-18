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
# DUMMY AI CLASSIFIER (replace later with real model)
# ---------------------------------------------------------
def classify_section(header: str, content: str) -> dict:
    """
    Very simple deterministic classifier.
    Replace later with Azure OpenAI / local model.
    """
    header_lower = header.lower()
    content_lower = content.lower()

    if "scope" in header_lower or "out of scope" in header_lower:
        category = "SCOPE"
    elif "feature" in header_lower or "service" in header_lower:
        category = "FUNCTIONAL"
    elif "incident" in header_lower or "problem" in header_lower:
        category = "SUPPORT"
    elif "license" in header_lower:
        category = "LICENSE"
    elif "sla" in header_lower or "service level" in header_lower:
        category = "SLA"
    elif "monitor" in header_lower:
        category = "MONITORING"
    elif "security" in header_lower:
        category = "SECURITY"
    else:
        category = "GENERAL"

    return {
        "header": header,
        "category": category,
        "content": content.strip()
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

        for header, content in data.items():
            mapped_entry = classify_section(header, content)
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