from pathlib import Path
import json
import subprocess
from docx import Document

# 1. INPUT verwijst naar jouw OneDrive-share
INPUT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

# 2. OUTPUT blijft in de repo
OUTPUT = Path("extracted")
OUTPUT.mkdir(exist_ok=True)

PROMPT = ".github/prompts/extract-service-fields.prompt.md"

def call_copilot_prompt(text, prompt_path):
    cmd = [
        "gh", "copilot", "prompt", "apply",
        "--prompt-file", prompt_path,
        "--input", text
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.stdout

def extract_doc(doc_path):
    doc = Document(doc_path)
    text = "\n".join([p.text for p in doc.paragraphs])
    json_output = call_copilot_prompt(text, PROMPT)
    return json_output

def main():
    print(f"Scanning input folder: {INPUT}")

    for sd in INPUT.rglob("SD*.docx"):
        print(f"Extracting: {sd}")
        data = extract_doc(sd)

        out_file = OUTPUT / f"{sd.stem}.json"
        out_file.write_text(data, encoding="utf-8")
        print(f" → Saved as {out_file}")

if __name__ == "__main__":
    main()