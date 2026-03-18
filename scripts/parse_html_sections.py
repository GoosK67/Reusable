from pathlib import Path
import sys
import json
from bs4 import BeautifulSoup
from datetime import datetime

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
        sections[current_header] = ""

        for el in soup.find_all(["h1", "h2", "h3", "p"]):
            if el.name in ("h1", "h2", "h3"):
                current_header = el.get_text(strip=True)
                sections[current_header] = ""
            else:
                sections[current_header] += el.get_text(" ", strip=True) + " "

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