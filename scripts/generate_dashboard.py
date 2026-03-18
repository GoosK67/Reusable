import html
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
LOG_FOLDER = BASE / "log"
OUT_FOLDER = BASE / "dashboard"
OUT_FOLDER.mkdir(exist_ok=True)

OUT_FILE = OUT_FOLDER / "presales_status.html"


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

        if "XML→DOCX OK" in line:
            last_success_step = "xml_to_docx.py"

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
    }


def build_dashboard():
    logs = list(LOG_FOLDER.glob("*.log"))

    # skip global run logs
    sd_logs = [x for x in logs if not x.name.startswith("run_")]

    rows = [parse_logfile(log) for log in sd_logs]

    rows.sort(key=lambda r: (0 if r["status"] == "ERROR" else 1, r["sd"].lower()))

    html_text = """
<html>
<head>
<meta http-equiv="refresh" content="5">
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
</style>
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

    html_text += f"<strong>Total:</strong> {total} | <strong>OK:</strong> {ok} | <strong>Errors:</strong> {errors} | <strong>Running:</strong> {running}"
    html_text += "</div>"

    html_text += """

<table>
<tr>
  <th>SD File</th>
  <th>Status</th>
    <th>Laatste succesvolle stap</th>
  <th>Step</th>
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
    print(f"✔ Dashboard created → {OUT_FILE}")


if __name__ == "__main__":
    build_dashboard()