from bs4 import BeautifulSoup
from pathlib import Path
import json
import sys
import re

def clean_html(text):
    # remove UTF-8 BOM if present
    if text.startswith("\ufeff"):
        text = text.replace("\ufeff", "", 1)

    # remove <style>, <script>, <xml>, <m:math> blocks
    text = re.sub(r"<style.*?>.*?</style>", "", text, flags=re.S|re.I)
    text = re.sub(r"<script.*?>.*?</script>", "", text, flags=re.S|re.I)
    text = re.sub(r"<xml.*?>.*?</xml>", "", text, flags=re.S|re.I)
    text = re.sub(r"<m:math.*?>.*?</m:math>", "", text, flags=re.S|re.I)

    return text

if __name__ == "__main__":
    html_file = Path(sys.argv[1])
    output_folder = Path("output/json")
    output_folder.mkdir(parents=True, exist_ok=True)

    raw = html_file.read_text(encoding="utf-8", errors="ignore")
    cleaned = clean_html(raw)

    soup = BeautifulSoup(cleaned, "html.parser")

    sections = {}
    current = "preamble"
    sections[current] = ""

    # robust detection
    for tag in soup.find_all(True):
        name = tag.name.lower()

        if name in ["h1", "h2", "h3"]:
            current = tag.get_text(strip=True).lower()
            sections.setdefault(current, "")
            continue

        if name == "p":
            text = tag.get_text(" ", strip=True)
            if text:
                sections[current] += text + "\n"
            continue

    out_path = output_folder / f"{html_file.stem}.json"
    out_path.write_text(json.dumps(sections, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"✔ Robust parsed → {out_path}")