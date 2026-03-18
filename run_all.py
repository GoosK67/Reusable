import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from lxml import etree

BASE_DIR = Path(__file__).parent.resolve()
SCRIPTS_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "log"
DEFAULT_SD_ROOT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

EXTRACT_SCRIPT = SCRIPTS_DIR / "extract_html.py"
PARSE_SCRIPT = SCRIPTS_DIR / "parse_html_sections.py"
AUTO_MAP_SCRIPT = SCRIPTS_DIR / "auto_map_sections.py"
XML_TO_DOCX_SCRIPT = SCRIPTS_DIR / "xml_to_docx.py"
DASHBOARD_SCRIPT = SCRIPTS_DIR / "generate_dashboard.py"

LOG_DIR.mkdir(exist_ok=True)


def log(message: str, sd_name: str = "GENERAL") -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    log_file = LOG_DIR / f"{sd_name}.log"
    with open(log_file, "a", encoding="utf-8", errors="ignore") as f:
        f.write(line + "\n")
    print(line, flush=True)


def run_step(script_path: Path, input_path: Path, sd_name: str) -> None:
    cmd = [sys.executable, str(script_path), str(input_path)]
    log(f"RUN {script_path.name} -> {input_path}", sd_name)
    subprocess.run(cmd, check=True, cwd=str(BASE_DIR))


def sanitize_xml_text(text: str) -> str:
    if text is None:
        return ""
    return "".join(ch for ch in text if ch in ("\t", "\n", "\r") or ord(ch) >= 0x20)


def mapped_json_to_xml(mapped_json_path: Path, sd_name: str) -> Path:
    mapped = json.loads(mapped_json_path.read_text(encoding="utf-8", errors="ignore"))

    root = etree.Element("ServiceDescription")

    for header, entry in mapped.items():
        section = etree.SubElement(root, "Section")
        section.set("name", sanitize_xml_text(str(header)))

        header_el = etree.SubElement(section, "Header")
        header_el.text = sanitize_xml_text(str(header))

        category_el = etree.SubElement(section, "Category")
        category_el.text = sanitize_xml_text(str(entry.get("category", "")))

        content_el = etree.SubElement(section, "Content")
        content_el.text = sanitize_xml_text(str(entry.get("content", "")))

    out_dir = BASE_DIR / "output" / "xml"
    out_dir.mkdir(parents=True, exist_ok=True)

    out_file = out_dir / f"{sd_name}_mapped.xml"
    out_file.write_bytes(
        etree.tostring(root, encoding="UTF-8", xml_declaration=True, pretty_print=True)
    )

    log(f"JSON->XML OK -> {out_file}", sd_name)
    return out_file


def process_one(sd_docx: Path) -> bool:
    sd_name = sd_docx.stem

    try:
        log(f"START PIPELINE for {sd_docx}", sd_name)

        run_step(EXTRACT_SCRIPT, sd_docx, sd_name)

        html_path = BASE_DIR / "extracted_html" / f"{sd_docx.stem}.html"
        if not html_path.exists():
            raise FileNotFoundError(f"Missing HTML output: {html_path}")

        run_step(PARSE_SCRIPT, html_path, sd_name)

        json_path = BASE_DIR / "output" / "json" / f"{sd_docx.stem}.json"
        if not json_path.exists():
            raise FileNotFoundError(f"Missing JSON output: {json_path}")

        run_step(AUTO_MAP_SCRIPT, json_path, sd_name)

        mapped_json_path = BASE_DIR / "output" / "mapped" / f"{sd_docx.stem}_mapped.json"
        if not mapped_json_path.exists():
            raise FileNotFoundError(f"Missing mapped JSON output: {mapped_json_path}")

        mapped_xml_path = mapped_json_to_xml(mapped_json_path, sd_docx.stem)

        run_step(XML_TO_DOCX_SCRIPT, mapped_xml_path, sd_name)

        log("PIPELINE OK", sd_name)
        return True

    except subprocess.CalledProcessError as exc:
        log(f"STEP FAILED ({exc.returncode}): {exc}", sd_name)
        return False
    except Exception as exc:
        log(f"PIPELINE ERROR: {exc}", sd_name)
        return False


def discover_docx(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        return [input_path]

    pattern = "**/SD*.docx" if recursive else "SD*.docx"
    return sorted(input_path.glob(pattern))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run SD pipeline in sequence: extract_html.py -> parse_html_sections.py "
            "-> auto_map_sections.py -> xml_to_docx.py"
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_SD_ROOT),
        help=(
            "Path to one SD .docx file or a folder containing SD*.docx "
            "(default: Product Management Library root)"
        ),
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        default=True,
        help="Recursively search for SD*.docx when input is a folder (default: enabled)",
    )
    parser.add_argument(
        "--no-recursive",
        dest="recursive",
        action="store_false",
        help="Disable recursive search when input is a folder",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)

    if not input_path.exists():
        log(f"Input path not found: {input_path}")
        return 1

    sd_files = discover_docx(input_path, recursive=args.recursive)

    if not sd_files:
        log(f"No SD*.docx files found in: {input_path}")
        return 1

    log(f"Found {len(sd_files)} SD file(s) to process")

    ok_count = 0
    fail_count = 0

    for sd in sd_files:
        if process_one(sd):
            ok_count += 1
        else:
            fail_count += 1

    try:
        log("Generating dashboard...")
        subprocess.run([sys.executable, str(DASHBOARD_SCRIPT)], check=True, cwd=str(BASE_DIR))
        log("Dashboard generated")
    except Exception as exc:
        log(f"Dashboard generation failed: {exc}")

    log(f"DONE - OK: {ok_count} | FAILED: {fail_count}")
    return 0 if fail_count == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
