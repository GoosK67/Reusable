import html
import re
from pathlib import Path
from zipfile import ZipFile
from collections import Counter, defaultdict

from lxml import etree
from openpyxl import Workbook

BASE = Path(__file__).resolve().parents[1]
LOG_FOLDER = BASE / "log"
OUT_FOLDER = BASE / "dashboard"
OUT_FOLDER.mkdir(exist_ok=True)

OUT_FILE = OUT_FOLDER / "presales_status.html"
OUT_XLSX_FILE = OUT_FOLDER / "presales_status.xlsx"
DOCX_OUT_FOLDER = BASE / "output" / "docx"

HITL_PREFIX = "AI generated, teverifieren door HITL"
LOW_INFO_TEXT = "AI agent heeft te weinig info om dit zelf op te stellen"

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TEMPLATE_DOCX = BASE / "templates" / "presales_template_sdt_v2.docx"
NON_CHAPTER_TAGS = {"Customer"}


def _extract_ts(line: str):
    if line.startswith("[") and "]" in line:
        return line.split("]", 1)[0].strip("[")
    return ""


def _last_pipeline_slice(lines):
    starts = [i for i, line in enumerate(lines) if "START PIPELINE" in line]
    if not starts:
        return lines
    return lines[starts[-1]:]


def _extract_step_name(line: str):
    # Example: "... RUN parse_html_sections.py -> ..."
    if " RUN " not in line or " -> " not in line:
        return ""
    tail = line.split(" RUN ", 1)[1]
    return tail.split(" -> ", 1)[0].strip()


def _extract_why(lines):
    priority_markers = [
        "PIPELINE ERROR:",
        "STEP FAILED",
        "EXTRACT ERROR:",
        "PARSE ERROR:",
        "auto_map ERROR:",
        "XML ERROR:",
        "ERROR:",
    ]
    for line in reversed(lines):
        for marker in priority_markers:
            if marker in line:
                return line
    return ""


def _last_xml_to_docx_slice(lines):
    starts = [i for i, line in enumerate(lines) if "START xml_to_docx" in line]
    if not starts:
        return lines
    return lines[starts[-1]:]


def _find_latest_output_docx(sd_stem: str):
    """Find the most recent FINAL docx for an SD stem."""
    prefix_mapped = f"{sd_stem}_mapped_FINAL"
    prefix_plain = f"{sd_stem}_FINAL"

    candidates = [
        p for p in DOCX_OUT_FOLDER.glob("*_FINAL*.docx")
        if p.name.startswith(prefix_mapped) or p.name.startswith(prefix_plain)
    ]

    if not candidates:
        return None

    return max(candidates, key=lambda p: p.stat().st_mtime)


def _normalize_text(value: str):
    return " ".join((value or "").split()).strip().lower()


def _iter_tagged_sdts(docx_path: Path):
    with ZipFile(docx_path, "r") as z:
        for name in z.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue

            try:
                root = etree.fromstring(z.read(name))
            except Exception:
                continue

            for sdt in root.xpath(".//w:sdt", namespaces=NS):
                tag = sdt.xpath("./w:sdtPr/w:tag/@w:val", namespaces=NS)
                tag = tag[0].strip() if tag else ""
                if not tag or tag in NON_CHAPTER_TAGS:
                    continue

                texts = sdt.xpath(".//w:sdtContent//w:t/text()", namespaces=NS)
                value = " ".join(t.strip() for t in texts if t and t.strip()).strip()
                yield tag, value


def _build_template_pool():
    pool = defaultdict(Counter)
    if not TEMPLATE_DOCX.exists():
        return pool

    for tag, value in _iter_tagged_sdts(TEMPLATE_DOCX):
        pool[tag][_normalize_text(value)] += 1
    return pool


def _count_chapters_from_docx(docx_path: Path):
    """Count chapter completion from tagged SDT controls, using template baseline for open placeholders."""
    sd_count = 0
    ai_count = 0
    open_count = 0

    template_pool = _build_template_pool()

    for tag, value in _iter_tagged_sdts(docx_path):
        normalized_value = _normalize_text(value)

        if not normalized_value or LOW_INFO_TEXT.lower() in normalized_value or "[to be completed]" in normalized_value:
            open_count += 1
            continue

        if HITL_PREFIX.lower() in normalized_value:
            ai_count += 1
            continue

        # Unchanged template placeholder text counts as open.
        if template_pool[tag][normalized_value] > 0:
            template_pool[tag][normalized_value] -= 1
            open_count += 1
            continue

        sd_count += 1

    return {
        "sd_chapters": sd_count,
        "ai_chapters": ai_count,
        "open_chapters": open_count,
    }


def _count_sdt_from_docx(docx_path: Path):
    """Backward-compatible alias (kept to avoid changing callers)."""
    return _count_chapters_from_docx(docx_path)


def parse_embedded_images_count(sd_stem: str) -> int:
    """Extract number of embedded images from xml_to_docx log message."""
    mapped_log = LOG_FOLDER / f"{sd_stem}_mapped.log"
    if not mapped_log.exists():
        return 0

    lines = mapped_log.read_text(encoding="utf-8", errors="ignore").splitlines()
    for line in lines:
        if "Afbeeldingen inline ingevoegd in bijbehorende tekst:" in line:
            try:
                num = int(line.split(":")[-1].strip())
                return num
            except (ValueError, IndexError):
                pass
    return 0


def parse_quality_scores(sd_stem: str):
    mapped_log = LOG_FOLDER / f"{sd_stem}_mapped.log"
    if not mapped_log.exists():
        return {
            "quality_avg": 0,
            "quality_worst": 0,
            "low_score_chapters": 0,
            "low_score_tags": "",
        }

    lines = mapped_log.read_text(encoding="utf-8", errors="ignore").splitlines()
    scoped = _last_xml_to_docx_slice(lines)

    pat = (
        r"Quality SDT '([^']+)': overall=(\d+); coverage=(\d+); specificity=(\d+); "
        r"evidence_count=(\d+); policy_compliance=(\d+); fill_type=([^;]+); low_score=(yes|no)"
    )

    per_tag = {}
    for line in scoped:
        m = re.search(pat, line)
        if not m:
            continue
        tag = m.group(1).strip()
        per_tag[tag] = {
            "overall": int(m.group(2)),
            "low": m.group(8).strip().lower() == "yes",
        }

    if not per_tag:
        return {
            "quality_avg": 0,
            "quality_worst": 0,
            "low_score_chapters": 0,
            "low_score_tags": "",
        }

    scores = [v["overall"] for v in per_tag.values()]
    low_tags = sorted([k for k, v in per_tag.items() if v["low"]])
    return {
        "quality_avg": round(sum(scores) / len(scores), 1),
        "quality_worst": min(scores),
        "low_score_chapters": len(low_tags),
        "low_score_tags": ", ".join(low_tags),
    }


def parse_chapter_counts(sd_stem: str):
    """Read chapter-level counters from final DOCX; fallback to companion _mapped log."""
    docx_path = _find_latest_output_docx(sd_stem)
    if docx_path and docx_path.exists():
        try:
            return _count_chapters_from_docx(docx_path)
        except Exception:
            pass

    mapped_log = LOG_FOLDER / f"{sd_stem}_mapped.log"
    if not mapped_log.exists():
        return {"sd_chapters": 0, "ai_chapters": 0, "open_chapters": 0}

    lines = mapped_log.read_text(encoding="utf-8", errors="ignore").splitlines()
    scoped_lines = _last_xml_to_docx_slice(lines)

    sd_tags = set()
    ai_tags = set()
    open_tags = set()

    for line in scoped_lines:
        if "Filled SDT '" in line and "with XML section" in line:
            tag = line.split("Filled SDT '", 1)[1].split("'", 1)[0].strip()
            if tag:
                sd_tags.add(tag)

        if "AI aangevuld voor ontbrekend hoofdstuk '" in line:
            tag = line.split("AI aangevuld voor ontbrekend hoofdstuk '", 1)[1].split("'", 1)[0].strip()
            if tag:
                ai_tags.add(tag)

        # Backward compatibility with older xml_to_docx log phrasing.
        if "AI fallback ingevuld voor ontbrekend hoofdstuk '" in line:
            tag = line.split("AI fallback ingevuld voor ontbrekend hoofdstuk '", 1)[1].split("'", 1)[0].strip()
            if tag:
                ai_tags.add(tag)

        if "AI open wegens te weinig info voor hoofdstuk '" in line:
            tag = line.split("AI open wegens te weinig info voor hoofdstuk '", 1)[1].split("'", 1)[0].strip()
            if tag:
                open_tags.add(tag)

    return {
        "sd_chapters": len(sd_tags),
        "ai_chapters": len(ai_tags),
        "open_chapters": len(open_tags),
    }


def parse_logfile(path: Path):
    """Parse latest pipeline run for one SD logfile."""
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    scoped_lines = _last_pipeline_slice(lines)

    start = None
    end = None
    status = "UNKNOWN"
    failed_step = ""
    last_success_step = ""
    all_error_lines = []
    why = ""
    current_step = ""

    for line in scoped_lines:
        if "START PIPELINE" in line:
            start = _extract_ts(line)

        if "RUN " in line and " -> " in line:
            current_step = _extract_step_name(line)
            failed_step = current_step

        if "extract_html OK" in line:
            last_success_step = "extract_html.py"

        if "PARSE OK" in line:
            last_success_step = "parse_html_sections.py"

        if "auto_map OK" in line:
            last_success_step = "auto_map_sections.py"

        if "JSON->XML OK" in line:
            last_success_step = "json_to_xml (internal)"

        if "XML" in line and "DOCX OK" in line:
            last_success_step = "xml_to_docx.py"
            status = "OK"
            end = _extract_ts(line)

        if "PIPELINE OK" in line:
            status = "OK"
            end = _extract_ts(line)
            last_success_step = "PIPELINE COMPLETED"

        if "STEP FAILED" in line or "PIPELINE ERROR" in line:
            status = "ERROR"
            end = _extract_ts(line)

        if "ERROR" in line or "FAILED" in line or "Traceback" in line:
            all_error_lines.append(line)

    why = _extract_why(scoped_lines)

    if status == "UNKNOWN":
        if all_error_lines:
            status = "ERROR"
        elif start:
            status = "RUNNING"

    if status == "OK":
        failed_step = ""
        why = ""

    return {
        "sd": path.stem,
        "status": status,
        "start": start,
        "end": end,
        "step": failed_step,
        "last_success_step": last_success_step,
        "why": why,
        "errors": all_error_lines[-10:],
        "embedded_images": parse_embedded_images_count(path.stem),
        **parse_chapter_counts(path.stem),
        **parse_quality_scores(path.stem),
    }


def export_dashboard_xlsx(rows):
    headers = [
        "SD File",
        "Status",
        "Laatste succesvolle stap",
        "Step",
        "Hoofdstukken uit SD",
        "AI aangevuld",
        "Open (te weinig info)",
        "Kwaliteit gem.",
        "Kwaliteit laagste",
        "Low-score hoofdstukken",
        "Low-score tags",
        "Afbeeldingen",
        "Waarom",
        "Start",
        "End",
        "Recent Errors",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.append(headers)

    for row in rows:
        ws.append([
            row.get("sd", ""),
            row.get("status", ""),
            row.get("last_success_step", ""),
            row.get("step", ""),
            row.get("sd_chapters", 0),
            row.get("ai_chapters", 0),
            row.get("open_chapters", 0),
            row.get("quality_avg", 0),
            row.get("quality_worst", 0),
            row.get("low_score_chapters", 0),
            row.get("low_score_tags", ""),
            row.get("embedded_images", 0),
            row.get("why", ""),
            row.get("start", ""),
            row.get("end", ""),
            "\n".join(row.get("errors", []) or []),
        ])

    ws.freeze_panes = "A2"

    # Keep widths readable by default.
    widths = {
        "A": 60,
        "B": 12,
        "C": 26,
        "D": 24,
        "E": 20,
        "F": 14,
        "G": 22,
        "H": 14,
        "I": 14,
        "J": 18,
        "K": 44,
        "L": 12,
        "M": 50,
        "N": 20,
        "O": 20,
        "P": 70,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(OUT_XLSX_FILE)


def build_dashboard():
    logs = list(LOG_FOLDER.glob("*.log"))

    # skip global run logs and internal xml_to_docx auxiliary logs (_mapped.log)
    sd_logs = [
        x for x in logs
        if not x.name.startswith("run_")
        and not x.name.endswith("_mapped.log")
        and x.name != "GENERAL.log"
    ]

    rows = [parse_logfile(log) for log in sd_logs]

    # Primary sort: highest open chapter count first.
    # Secondary sort: errors first, then alphabetical SD name.
    rows.sort(
        key=lambda r: (
            -int(r.get("low_score_chapters", 0)),
            -int(r.get("open_chapters", 0)),
            0 if r["status"] == "ERROR" else 1,
            r["sd"].lower(),
        )
    )

    html_text = """
<html>
<head>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
table { width: 100%; border-collapse: collapse; }
th { background: #222; color: white; padding: 8px; text-align: left; }
td { border: 1px solid #ccc; padding: 6px; vertical-align: top; }
.ok { background: #c8f7c5; }
.err { background: #f6b2b2; }
.run { background: #f8e59a; }
.summary { margin-bottom: 12px; }
.small { font-size: 12px; color: #444; }
.toolbar { margin: 10px 0 14px 0; }
.toolbar input { min-width: 360px; padding: 6px 8px; font-size: 14px; }
</style>
<script>
function filterBySd() {
    const input = document.getElementById('sdFilter');
    const filter = (input.value || '').toLowerCase();
    const table = document.getElementById('statusTable');
    if (!table) return;

    const rows = table.getElementsByTagName('tr');
    for (let i = 1; i < rows.length; i++) {
        const firstCell = rows[i].getElementsByTagName('td')[0];
        if (!firstCell) continue;
        const txt = (firstCell.textContent || firstCell.innerText || '').toLowerCase();
        rows[i].style.display = txt.indexOf(filter) > -1 ? '' : 'none';
    }
}

function shouldAutoRefresh() {
    const input = document.getElementById('sdFilter');
    if (!input) return true;
    const hasFilter = (input.value || '').trim().length > 0;
    const hasFocus = document.activeElement === input;
    return !hasFilter && !hasFocus;
}

setInterval(function () {
    if (shouldAutoRefresh()) {
        window.location.reload();
    }
}, 5000);
</script>
</head>
<body>

<h1>Presales Processing Dashboard</h1>
<p>Auto-refresh every 5 seconds</p>

<div class="summary">
"""

    total = len(rows)
    errors = sum(1 for r in rows if r["status"] == "ERROR")
    running = sum(1 for r in rows if r["status"] == "RUNNING")
    ok = sum(1 for r in rows if r["status"] == "OK")
    total_sd_chapters = sum(r.get("sd_chapters", 0) for r in rows)
    total_ai_chapters = sum(r.get("ai_chapters", 0) for r in rows)
    total_open_chapters = sum(r.get("open_chapters", 0) for r in rows)
    total_low_score_chapters = sum(r.get("low_score_chapters", 0) for r in rows)
    quality_avgs = [float(r.get("quality_avg", 0) or 0) for r in rows if float(r.get("quality_avg", 0) or 0) > 0]
    avg_quality_all = round(sum(quality_avgs) / len(quality_avgs), 1) if quality_avgs else 0

    html_text += f"<strong>Total:</strong> {total} | <strong>OK:</strong> {ok} | <strong>Errors:</strong> {errors} | <strong>Running:</strong> {running}"
    html_text += "<br>"
    html_text += f"<strong>Hoofdstukken uit SD:</strong> {total_sd_chapters} | <strong>AI aangevuld:</strong> {total_ai_chapters} | <strong>Open (te weinig info):</strong> {total_open_chapters}"
    html_text += "<br>"
    html_text += f"<strong>Gem. kwaliteit:</strong> {avg_quality_all} | <strong>Low-score hoofdstukken:</strong> {total_low_score_chapters}"
    html_text += "</div>"

    html_text += """
<div class="toolbar">
    <label for="sdFilter"><strong>Filter SD:</strong></label>
    <input id="sdFilter" type="text" placeholder="Typ (deel van) documentnaam..." onkeyup="filterBySd()" />
</div>

<table id="statusTable">
<tr>
  <th>SD File</th>
  <th>Status</th>
    <th>Laatste succesvolle stap</th>
  <th>Step</th>
    <th>Hoofdstukken uit SD</th>
    <th>AI aangevuld</th>
    <th>Open (te weinig info)</th>
    <th>Kwaliteit gem.</th>
    <th>Kwaliteit laagste</th>
    <th>Low-score hoofdstukken</th>
    <th>Low-score tags</th>
    <th>Afbeeldingen</th>
  <th>Waarom</th>
  <th>Start</th>
  <th>End</th>
  <th>Recent Errors</th>
</tr>
"""

    for row in rows:
        if row["status"] == "OK":
            cls = "ok"
        elif row["status"] == "ERROR":
            cls = "err"
        else:
            cls = "run"

        err_html = "<br>".join(html.escape(x) for x in row["errors"]) if row["errors"] else ""
        why = html.escape(row["why"] or "")
        step = html.escape(row["step"] or "")
        last_success_step = html.escape(row.get("last_success_step", "") or "")

        html_text += f"""
<tr class="{cls}">
  <td>{html.escape(row['sd'])}</td>
  <td>{html.escape(row['status'])}</td>
  <td>{last_success_step}</td>
  <td>{step}</td>
    <td>{row.get('sd_chapters', 0)}</td>
    <td>{row.get('ai_chapters', 0)}</td>
    <td>{row.get('open_chapters', 0)}</td>
    <td>{row.get('quality_avg', 0)}</td>
    <td>{row.get('quality_worst', 0)}</td>
    <td>{row.get('low_score_chapters', 0)}</td>
    <td class="small">{html.escape(row.get('low_score_tags', '') or '')}</td>
    <td>{row.get('embedded_images', 0)}</td>
  <td>{why}</td>
  <td>{html.escape(row['start'] or '')}</td>
  <td>{html.escape(row['end'] or '')}</td>
  <td class="small">{err_html}</td>
</tr>
"""

    html_text += """
</table>
</body>
</html>
"""

    OUT_FILE.write_text(html_text, encoding="utf-8")
    print(f"Dashboard created: {OUT_FILE}")
    export_dashboard_xlsx(rows)
    print(f"Dashboard XLSX created: {OUT_XLSX_FILE}")


if __name__ == "__main__":
    build_dashboard()