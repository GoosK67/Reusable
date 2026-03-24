import argparse
import json
import os
import re
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from lxml import etree
from openpyxl import Workbook

BASE_DIR = Path(__file__).parent.resolve()
SCRIPTS_DIR = BASE_DIR / "scripts"
LOG_DIR = BASE_DIR / "log"
MAPPING_XLSX_DIR = BASE_DIR / "output" / "mapping_xlsx"
DEFAULT_SD_ROOT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")

EXTRACT_SCRIPT = SCRIPTS_DIR / "extract_html.py"
PARSE_SCRIPT = SCRIPTS_DIR / "parse_html_sections.py"
AUTO_MAP_SCRIPT = SCRIPTS_DIR / "auto_map_sections.py"
XML_TO_DOCX_SCRIPT = SCRIPTS_DIR / "xml_to_docx.py"
DASHBOARD_SCRIPT = SCRIPTS_DIR / "generate_dashboard.py"

LOG_DIR.mkdir(exist_ok=True)
MAPPING_XLSX_DIR.mkdir(parents=True, exist_ok=True)

# Force UTF-8 stdout/stderr for all child processes so Unicode in log messages
# does not crash on Windows cp1252 console terminals.
UTF8_ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}


def normalize_windows_path(path: Path) -> str:
    p = str(path)
    if os.name != "nt":
        return p
    if p.startswith("\\\\?\\"):
        return p
    if len(p) >= 248:
        return "\\\\?\\" + p
    return p


def path_exists(path: Path) -> bool:
    if path.exists():
        return True
    try:
        return Path(normalize_windows_path(path)).exists()
    except Exception:
        return False


def path_is_file(path: Path) -> bool:
    if path.is_file():
        return True
    try:
        return Path(normalize_windows_path(path)).is_file()
    except Exception:
        return False


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
    subprocess.run(cmd, check=True, cwd=str(BASE_DIR), env=UTF8_ENV)


def run_xml_to_docx_step(xml_path: Path, source_sd_docx: Path, sd_name: str) -> None:
    cmd = [sys.executable, str(XML_TO_DOCX_SCRIPT), str(xml_path), str(source_sd_docx)]
    log(f"RUN {XML_TO_DOCX_SCRIPT.name} -> {xml_path} (source: {source_sd_docx})", sd_name)
    subprocess.run(cmd, check=True, cwd=str(BASE_DIR), env=UTF8_ENV)


def sanitize_xml_text(text: str) -> str:
    if text is None:
        return ""
    return "".join(ch for ch in text if ch in ("\t", "\n", "\r") or ord(ch) >= 0x20)


def _last_pipeline_slice(lines: list[str]) -> list[str]:
    starts = [i for i, line in enumerate(lines) if "START PIPELINE" in line]
    if not starts:
        return lines
    return lines[starts[-1]:]


def _last_xml_to_docx_slice(lines: list[str]) -> list[str]:
    starts = [i for i, line in enumerate(lines) if "START xml_to_docx" in line]
    if not starts:
        return lines
    return lines[starts[-1]:]


def export_sdt_mapping_xlsx(sd_name: str) -> Path:
    """Create an XLSX report showing which source chapters filled which SDT tags."""
    log_file = LOG_DIR / f"{sd_name}.log"
    mapped_log_file = LOG_DIR / f"{sd_name}_mapped.log"
    out_file = MAPPING_XLSX_DIR / f"{sd_name}_sdt_mapping.xlsx"

    lines: list[str] = []
    if mapped_log_file.exists():
        lines = mapped_log_file.read_text(encoding="utf-8", errors="ignore").splitlines()
        lines = _last_xml_to_docx_slice(lines)
    elif log_file.exists():
        lines = log_file.read_text(encoding="utf-8", errors="ignore").splitlines()
        if any("START xml_to_docx" in line for line in lines):
            lines = _last_xml_to_docx_slice(lines)
        else:
            lines = _last_pipeline_slice(lines)

    direct_re = re.compile(r"Filled SDT '([^']+)' with XML section '([^']+)'")
    trace_re = re.compile(r"Trace SDT '([^']+)': (.+)$")
    trace_conflict_re = re.compile(r"Trace SDT '([^']+)' conflict(?: \[severity=([^\]]+)\])?(?: \[([^\]]+)\])?: (.+)$")
    quality_re = re.compile(
        r"Quality SDT '([^']+)': overall=(\d+); coverage=(\d+); specificity=(\d+); "
        r"evidence_count=(\d+); policy_compliance=(\d+); fill_type=([^;]+); low_score=(yes|no)"
    )
    ai_related_re = re.compile(r"AI aangevuld op basis van gerelateerde documenten voor hoofdstuk '([^']+)'")
    ai_missing_re = re.compile(r"AI aangevuld voor ontbrekend hoofdstuk '([^']+)'")
    ai_open_re = re.compile(r"AI open wegens te weinig info voor hoofdstuk '([^']+)'")
    forced_open_re = re.compile(r"Forced open placeholder for hoofdstuk '([^']+)' wegens ontbrekende exacte bron-evidence")

    rows: list[dict[str, str]] = []
    row_num = 0

    for line in lines:
        m = direct_re.search(line)
        if m:
            row_num += 1
            rows.append({
                "order": str(row_num),
                "sdt_tag": m.group(1).strip(),
                "fill_type": "direct_from_sd_chapter",
                "source_chapter": m.group(2).strip(),
                "details": "",
                "evidence_status": "exact_evidence",
                "quality_score": "",
                "coverage_score": "",
                "specificity_score": "",
                "evidence_count_score": "",
                "policy_compliance_score": "",
                "quality_fill_type": "",
                "low_score": "",
            })
            continue

        m = trace_re.search(line)
        if m:
            t_tag = m.group(1).strip()
            t_reason = m.group(2).strip()
            for item in reversed(rows):
                if item.get("fill_type") == "direct_from_sd_chapter" and item.get("sdt_tag") == t_tag and not item.get("details"):
                    item["details"] = t_reason
                    break
            continue

        m = trace_conflict_re.search(line)
        if m:
            t_tag = m.group(1).strip()
            t_sev = (m.group(2) or "warning").strip()
            t_code = (m.group(3) or "n/a").strip()
            t_conflict = m.group(4).strip()
            for item in reversed(rows):
                if item.get("fill_type") == "direct_from_sd_chapter" and item.get("sdt_tag") == t_tag:
                    existing = (item.get("details") or "").strip()
                    extra = f"conflict[{t_sev}/{t_code}]: {t_conflict}"
                    if existing:
                        item["details"] = f"{existing} | {extra}"
                    else:
                        item["details"] = extra
                    break
            continue

        m = ai_related_re.search(line)
        if m:
            row_num += 1
            rows.append({
                "order": str(row_num),
                "sdt_tag": m.group(1).strip(),
                "fill_type": "ai_related_documents",
                "source_chapter": "",
                "details": "AI content based on related documents",
                "evidence_status": "related_docs_evidence",
                "quality_score": "",
                "coverage_score": "",
                "specificity_score": "",
                "evidence_count_score": "",
                "policy_compliance_score": "",
                "quality_fill_type": "",
                "low_score": "",
            })
            continue

        m = ai_missing_re.search(line)
        if m:
            row_num += 1
            rows.append({
                "order": str(row_num),
                "sdt_tag": m.group(1).strip(),
                "fill_type": "ai_missing_chapter",
                "source_chapter": "",
                "details": "AI fallback because no mapped SD chapter",
                "evidence_status": "missing_evidence",
                "quality_score": "",
                "coverage_score": "",
                "specificity_score": "",
                "evidence_count_score": "",
                "policy_compliance_score": "",
                "quality_fill_type": "",
                "low_score": "",
            })
            continue

        m = ai_open_re.search(line)
        if m:
            row_num += 1
            rows.append({
                "order": str(row_num),
                "sdt_tag": m.group(1).strip(),
                "fill_type": "open_too_little_info",
                "source_chapter": "",
                "details": "Not enough info for AI to fill automatically",
                "evidence_status": "missing_evidence",
                "quality_score": "",
                "coverage_score": "",
                "specificity_score": "",
                "evidence_count_score": "",
                "policy_compliance_score": "",
                "quality_fill_type": "",
                "low_score": "",
            })
            continue

        m = forced_open_re.search(line)
        if m:
            row_num += 1
            rows.append({
                "order": str(row_num),
                "sdt_tag": m.group(1).strip(),
                "fill_type": "open_too_little_info",
                "source_chapter": "",
                "details": "Forced open placeholder due to missing exact source evidence",
                "evidence_status": "forced_open_missing_evidence",
                "quality_score": "",
                "coverage_score": "",
                "specificity_score": "",
                "evidence_count_score": "",
                "policy_compliance_score": "",
                "quality_fill_type": "",
                "low_score": "",
            })
            continue

        m = quality_re.search(line)
        if m:
            q_tag = m.group(1).strip()
            q_overall = m.group(2).strip()
            q_coverage = m.group(3).strip()
            q_specificity = m.group(4).strip()
            q_evidence = m.group(5).strip()
            q_policy = m.group(6).strip()
            q_fill_type = m.group(7).strip()
            q_low = m.group(8).strip()

            for item in reversed(rows):
                if item.get("sdt_tag") == q_tag and not str(item.get("quality_score", "")).strip():
                    item["quality_score"] = q_overall
                    item["coverage_score"] = q_coverage
                    item["specificity_score"] = q_specificity
                    item["evidence_count_score"] = q_evidence
                    item["policy_compliance_score"] = q_policy
                    item["quality_fill_type"] = q_fill_type
                    item["low_score"] = q_low
                    break
            continue

    wb = Workbook()
    ws = wb.active
    ws.title = "SDT Mapping"

    ws.append([
        "SD File",
        "Order",
        "SDT Tag",
        "Fill Type",
        "Source Chapter",
        "Details",
        "Evidence Status",
        "Quality Score",
        "Coverage Score",
        "Specificity Score",
        "Evidence Count Score",
        "Policy Compliance Score",
        "Quality Fill Type",
        "Low Score",
    ])

    for item in rows:
        ws.append([
            sd_name,
            int(item["order"]),
            item["sdt_tag"],
            item["fill_type"],
            item["source_chapter"],
            item["details"],
            item.get("evidence_status", "unknown"),
            item.get("quality_score", ""),
            item.get("coverage_score", ""),
            item.get("specificity_score", ""),
            item.get("evidence_count_score", ""),
            item.get("policy_compliance_score", ""),
            item.get("quality_fill_type", ""),
            item.get("low_score", ""),
        ])

    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 48
    ws.column_dimensions["F"].width = 52
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 18
    ws.column_dimensions["L"].width = 20
    ws.column_dimensions["M"].width = 24
    ws.column_dimensions["N"].width = 12

    # Second sheet: compact summary per SDT tag and fill type.
    summary = defaultdict(Counter)
    for item in rows:
        tag = item.get("sdt_tag", "")
        fill_type = item.get("fill_type", "")
        if not tag:
            continue
        summary[tag]["total"] += 1
        summary[tag][fill_type] += 1

    ws_sum = wb.create_sheet(title="SDT Summary")
    ws_sum.append([
        "SD File",
        "SDT Tag",
        "Total",
        "Direct from SD chapter",
        "AI related documents",
        "AI missing chapter",
        "Open too little info",
    ])

    for tag in sorted(summary.keys()):
        c = summary[tag]
        ws_sum.append([
            sd_name,
            tag,
            c.get("total", 0),
            c.get("direct_from_sd_chapter", 0),
            c.get("ai_related_documents", 0),
            c.get("ai_missing_chapter", 0),
            c.get("open_too_little_info", 0),
        ])

    ws_sum.freeze_panes = "A2"
    ws_sum.column_dimensions["A"].width = 60
    ws_sum.column_dimensions["B"].width = 34
    ws_sum.column_dimensions["C"].width = 10
    ws_sum.column_dimensions["D"].width = 24
    ws_sum.column_dimensions["E"].width = 22
    ws_sum.column_dimensions["F"].width = 18
    ws_sum.column_dimensions["G"].width = 20

    wb.save(out_file)
    return out_file


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

        table_facts = entry.get("table_facts", []) if isinstance(entry, dict) else []
        if table_facts:
            facts_el = etree.SubElement(section, "TableFactsJson")
            facts_el.text = sanitize_xml_text(json.dumps(table_facts, ensure_ascii=False))

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

        run_xml_to_docx_step(mapped_xml_path, sd_docx, sd_name)

        try:
            mapping_xlsx = export_sdt_mapping_xlsx(sd_name)
            log(f"SDT mapping XLSX OK -> {mapping_xlsx}", sd_name)
        except Exception as exc:
            log(f"SDT mapping XLSX ERROR: {exc}", sd_name)

        log("PIPELINE OK", sd_name)
        return True

    except subprocess.CalledProcessError as exc:
        log(f"STEP FAILED ({exc.returncode}): {exc}", sd_name)
        return False
    except Exception as exc:
        log(f"PIPELINE ERROR: {exc}", sd_name)
        return False


def discover_docx(input_path: Path, recursive: bool) -> list[Path]:
    if path_is_file(input_path):
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

    if not path_exists(input_path):
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
        subprocess.run([sys.executable, str(DASHBOARD_SCRIPT)], check=True, cwd=str(BASE_DIR), env=UTF8_ENV)
        log("Dashboard generated")
    except Exception as exc:
        log(f"Dashboard generation failed: {exc}")

    log(f"DONE - OK: {ok_count} | FAILED: {fail_count}")
    return 0 if fail_count == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
