from pathlib import Path
from docx import Document

INPUT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")
OUTPUT = Path("extracted")
OUTPUT.mkdir(exist_ok=True)

def extract_docx_text(doc_path: Path) -> str:
    doc = Document(doc_path)
    return "\n".join([p.text for p in doc.paragraphs])

def main():
    print(f"Scanning SD folder: {INPUT}")

    # 100% OneDrive-safe search
    for sd in INPUT.rglob("*.docx"):
        if not sd.name.startswith("SD"):
            continue

        print(f"Extracting: {sd}")

        try:
            raw_text = extract_docx_text(sd)
        except Exception as e:
            print(f"❌ Could not read file: {sd}")
            print(f"   Error: {e}")
            continue

        out_file = OUTPUT / f"{sd.stem}.raw.txt"
        out_file.write_text(raw_text, encoding="utf-8")
        print(f" → Saved raw extract to {out_file}")

if __name__ == "__main__":
    main()