from pathlib import Path
from extractor import extract_sd
from mapper import map_to_presales
from generator import generate_presales

# 👉 Zet hier het pad van jouw OneDrive sync:
INPUT_ROOT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

OUTPUT_ROOT = Path("output")
TEMPLATE_PATH = Path("templates/presales_template.docx")

def main():
    print("STEP 1: Scanning OneDrive-synchronized SharePoint folders...")
    print(f"Using input folder: {INPUT_ROOT}\n")

    OUTPUT_ROOT.mkdir(exist_ok=True)

    # Loop door alle DOCX in alle mappen en submappen
    for sd in INPUT_ROOT.rglob("*.docx"):
        # Enkel SD’s verwerken
        if sd.name.lower().startswith("sd"):
            print(f"Processing: {sd}")

            # Extract → Map → Generate
            try:
                data = extract_sd(sd)
                fields = map_to_presales(data)

                output_file = OUTPUT_ROOT / f"{sd.stem} - Presales Guide.docx"
                generate_presales(fields, TEMPLATE_PATH, output_file)

                print(f"Generated: {output_file}\n")

            except Exception as e:
                print(f"❌ Error processing {sd}: {e}\n")

    print("🎉 DONE — All Presales Guides generated!")

if __name__ == "__main__":
    main()