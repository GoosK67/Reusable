import os
import json
from pathlib import Path
from docx import Document
from openai import AzureOpenAI

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------

# 1. SD source folder (OneDrive sync)
SD_INPUT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

# 2. Output folder
OUTPUT = Path("extracted")
OUTPUT.mkdir(exist_ok=True)

# 3. Azure OpenAI config
client = AzureOpenAI(
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version="2024-08-01-preview"
)

MODEL = "gpt-4o-mini"   # snel + goedkoop + perfect voor extractie


# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------

def read_docx(path: Path) -> str:
    """Convert DOCX to plain text."""
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def extract_json_from_llm(raw_text: str) -> dict:
    """Send content to LLM and get structured JSON back."""
    prompt = f"""
Je bent een extractie‑engine. Zet onderstaande SD om in strikt JSON.

VELDEN:
- title
- service_summary
- business_context
- key_features
- standard_services
- optional_services
- operational_services
- prerequisites
- out_of_scope
- conditions
- sla
- pricing
- risks
- assumptions
- differentiators
- missing_information

REGELS:
- Gebruik arrays waar mogelijk
- Vul ontbrekende dingen in als null
- Geen vrije tekst buiten JSON
- Verzin niets
- JSON moet parseerbaar zijn

SD INHOUD:
"""

    response = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        messages=[
            {"role": "system", "content": "Je bent een expert in documentextractie."},
            {"role": "user", "content": prompt}
        ]
    )

    clean = response.choices[0].message.content.strip()

    # JSON reparatie indien nodig
    if clean.startswith("```"):
        clean = clean.strip("`").strip()

    return json.loads(clean)


# ---------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------

def main():
    print(f"Scanning SD folder: {SD_INPUT}")

    for file in SD_INPUT.rglob("*.docx"):
        if not file.name.startswith("SD"):
            continue

        print(f"\n➡ Extracting: {file}")

        try:
            raw_text = read_docx(file)
        except Exception as ex:
            print(f"❌ Error reading DOCX: {ex}")
            continue

        try:
            data = extract_json_from_llm(raw_text)
        except Exception as ex:
            print(f"❌ Error from LLM: {ex}")
            continue

        out_path = OUTPUT / f"{file.stem}.json"
        out_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

        print(f"✅ Saved JSON → {out_path}")


if __name__ == "__main__":
    main()