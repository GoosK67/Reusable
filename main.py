from pathlib import Path
from extractor import extract_sections
from mapper import DEFAULT_REWRITE_PROFILE, REWRITE_PROFILES, map_sd_to_template
from generator import fill_template

# 👉 Zet hier het pad van jouw OneDrive sync:
INPUT_ROOT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

OUTPUT_ROOT = Path("output")
TEMPLATE_PATH = Path("templates/presales_template.docx")


def _to_fill_fields(mapped_sections):
    """Convert human-readable template keys to DOCX tag field keys.
    
    When full_section=True with include_tables=True, mapped_sections values
    may be dicts with {text, tables}. We return a structure that preserves this."""
    # Extract the source titles mapping before processing
    source_titles = mapped_sections.pop("_source_titles", {})
    
    result = {}
    for template_field, field_key in [
        ("Product Summary", "ProductSummary"),
        ("Value Proposition", "ValueProposition"),
        ("Product Description", "ProductDescription"),
        ("Requirements & Prerequisites", "Requirements"),
        ("Scope / Out of Scope", "Scope"),
        ("SLA", "SLA"),
        ("Operational Support", "OperationalSupport"),
    ]:
        value = mapped_sections.get(template_field, "")
        if isinstance(value, dict):
            # Structure: {text: str, tables: list}
            result[field_key] = value
        else:
            # Backward compatible: plain string
            result[field_key] = {"text": value, "tables": []}
    
    # Add the source titles mapping for the generator
    result["_source_titles"] = source_titles
    return result


def process_all_sd_files(root_folder, template_path, output_folder, rewrite_profile=DEFAULT_REWRITE_PROFILE):
    """
    Process all SD DOCX files recursively:
    1. extract_sections
    2. map_sd_to_template
    3. fill_template
    """
    root = Path(root_folder)
    template = Path(template_path)
    output = Path(output_folder)

    output.mkdir(parents=True, exist_ok=True)

    template_resolved = template.resolve()
    output_resolved = output.resolve()

    for sd_file in root.rglob("*.docx"):
        sd_resolved = sd_file.resolve()

        # Only process Service Description files (filename must start with "SD").
        if not sd_file.name.lower().startswith("sd"):
            continue

        # Avoid processing the template itself or already generated files.
        if sd_resolved == template_resolved:
            continue
        if output_resolved in sd_resolved.parents:
            continue

        print(f"Processing: {sd_file}")
        try:
            from docx import Document
            source_doc = Document(sd_file)
            sections = extract_sections(sd_file)
            mapped_sections = map_sd_to_template(
                sections,
                rewrite_profile=rewrite_profile,
                full_section=True,
                preserve_titles=True,
                include_tables=True,
            )
            fields = _to_fill_fields(mapped_sections)
            fields["_source_doc"] = source_doc
            fields["_source_sections"] = sections

            output_file = output / f"{sd_file.stem} - Presales Guide.docx"
            fill_template(template, output_file, fields)
            print(f"Generated: {output_file}\n")
        except Exception as exc:
            print(f"Error processing {sd_file}: {exc}\n")


def choose_rewrite_profile():
    """Simple runtime switch to select rewrite tone profile."""
    available_profiles = list(REWRITE_PROFILES.keys())

    print("Available rewrite profiles:")
    for index, profile_name in enumerate(available_profiles, start=1):
        default_label = " (default)" if profile_name == DEFAULT_REWRITE_PROFILE else ""
        print(f"{index}. {profile_name}{default_label}")

    user_choice = input(
        f"Choose rewrite profile [1-{len(available_profiles)}] or press Enter for default: "
    ).strip()

    if not user_choice:
        return DEFAULT_REWRITE_PROFILE

    if user_choice.isdigit():
        selected_index = int(user_choice) - 1
        if 0 <= selected_index < len(available_profiles):
            return available_profiles[selected_index]

    if user_choice in REWRITE_PROFILES:
        return user_choice

    print(f"Invalid choice '{user_choice}', falling back to default profile: {DEFAULT_REWRITE_PROFILE}\n")
    return DEFAULT_REWRITE_PROFILE


def main():
    print("Scanning folders for SD DOCX files...")
    print(f"Using input folder: {INPUT_ROOT}\n")
    selected_profile = choose_rewrite_profile()
    print(f"Using rewrite profile: {selected_profile}\n")

    process_all_sd_files(INPUT_ROOT, TEMPLATE_PATH, OUTPUT_ROOT, rewrite_profile=selected_profile)

    print("Done - all Presales Guides generated.")

if __name__ == "__main__":
    main()