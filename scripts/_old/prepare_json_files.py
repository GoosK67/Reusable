from bs4 import BeautifulSoup
import json
from pathlib import Path
import sys

if __name__ == "__main__":
    html_file = Path(sys.argv[1])  # EXACT 1 file here
    output_folder = Path("output/json")
    output_folder.mkdir(parents=True, exist_ok=True)

    soup = BeautifulSoup(html_file.read_text(encoding="utf-8"), "html.parser")
    sections = {}
    current = "preamble"
    sections[current] = ""

    for tag in soup.find_all(["h1", "h2", "h3", "p", "div"]):
        if tag.name in ["h1", "h2", "h3"]:
            current = tag.text.strip().lower()
            sections[current] = ""
        else:
            sections[current] += tag.text + "\n"

    out = output_folder / f"{html_file.stem}.json"
    out.write_text(json.dumps(sections, indent=2), encoding="utf-8")
    print(f"✔ Parsed → {out}")