from sharepoint_downloader import fetch_all_sd_files
from extractor import extract_sd
from mapper import map_to_presales
from generator import generate_presales
from pathlib import Path

def main():
    print("STEP 1: Downloading SDs from SharePoint...")
    fetch_all_sd_files()

    print("STEP 2: Processing SD files...")
    input_dir = Path("input")
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    for sd in input_dir.rglob("*.docx"):
        print("Processing:", sd.name)

        data = extract_sd(sd)
        fields = map_to_presales(data)

        output_file = output_dir / f"{sd.stem} - Presales Guide.docx"
        generate_presales(fields, "templates/presales_template.docx", output_file)

    print("DONE — All Presales Guides generated!")

if __name__ == "__main__":
    main()