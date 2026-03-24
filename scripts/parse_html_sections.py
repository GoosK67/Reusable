from pathlib import Path
import sys
import json
from bs4 import BeautifulSoup
from datetime import datetime
import re

# -----------------------------------------
# LOGGING (ALTIJD APPEND)
# -----------------------------------------
LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)

def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"

    with open(logfile, "a", encoding="utf-8", errors="ignore") as f:
        f.write(line)

    print(line, end="")


def _clean_cell_text(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


def _normalize_fact_key(text: str) -> str:
    key = _clean_cell_text(text).lower()
    key = re.sub(r"[^a-z0-9]+", "_", key)
    key = re.sub(r"_+", "_", key).strip("_")
    return key[:64] if key else "field"


def _detect_fact_type(row_text: str) -> str:
    t = row_text.lower()
    if any(k in t for k in ["pricing", "billing", "cost", "charge", "invoice", "sku", "unit", "monthly", "one-time"]):
        return "pricing"
    if any(k in t for k in ["sla", "kpi", "availability", "uptime", "response", "resolution", "service level", "window"]):
        return "service_level"
    if any(k in t for k in ["support", "incident", "request", "severity", "escalation", "hours", "operation"]):
        return "operations"
    if any(k in t for k in ["in scope", "out of scope", "included", "excluded"]):
        return "scope"
    return "general"


def _table_to_structured_facts(table_el):
    raw_rows = []
    for tr in table_el.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        row = [_clean_cell_text(c.get_text(" ", strip=True)) for c in cells]
        row = [c for c in row if c]
        if row:
            raw_rows.append(row)

    if len(raw_rows) < 2:
        return []

    header = raw_rows[0]
    body = raw_rows[1:]
    facts = []

    for idx, row in enumerate(body, start=1):
        pairs = {}
        for i, value in enumerate(row):
            col_name = header[i] if i < len(header) else f"column_{i + 1}"
            key = _normalize_fact_key(col_name)
            if value:
                pairs[key] = value

        if not pairs:
            continue

        row_text = " | ".join(row)
        facts.append(
            {
                "fact_type": _detect_fact_type(" ".join([" ".join(header), row_text])),
                "row_index": idx,
                "row_text": row_text,
                "facts": pairs,
            }
        )

    return facts

# -----------------------------------------
# HTML PARSER
# -----------------------------------------
if __name__ == "__main__":
    html_file = Path(sys.argv[1])
    sd_name = html_file.stem

    log(f"START parse_html for: {html_file}", sd_name)

    try:
        html_text = html_file.read_text(encoding="utf-8", errors="ignore")
        soup = BeautifulSoup(html_text, "html.parser")

        sections = {}

        # 🟩 CRUCIALE FIX: fallback moet ALTIJD bestaan
        current_header = "UNCLASSIFIED"
        sections[current_header] = {
            "content": "",
            "table_facts": [],
        }

        for el in soup.find_all(["h1", "h2", "h3", "p", "table"]):
            if el.name in ("h1", "h2", "h3"):
                current_header = el.get_text(strip=True)
                sections[current_header] = {
                    "content": "",
                    "table_facts": [],
                }
                continue

            if el.name == "p":
                if el.find_parent("table") is not None:
                    continue
                paragraph = el.get_text(" ", strip=True)
                if paragraph:
                    sections[current_header]["content"] += paragraph + " "
                continue

            if el.name == "table":
                facts = _table_to_structured_facts(el)
                if facts:
                    sections[current_header]["table_facts"].extend(facts)

        # SAVE RESULT
        out_folder = Path("output/json")
        out_folder.mkdir(parents=True, exist_ok=True)

        out_file = out_folder / f"{sd_name}.json"
        out_file.write_text(
            json.dumps(sections, indent=2, ensure_ascii=False),
            encoding="utf-8",
            errors="ignore"
        )

        log(f"PARSE OK → {out_file}", sd_name)
        sys.exit(0)

    except Exception as e:
        log(f"PARSE ERROR: {e}", sd_name)
        sys.exit(1)