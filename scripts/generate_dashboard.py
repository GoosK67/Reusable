import json
from pathlib import Path
from datetime import datetime

BASE = Path(__file__).resolve().parents[1]
LOG_FOLDER = BASE / "log"
OUT_FOLDER = BASE / "dashboard"
OUT_FOLDER.mkdir(exist_ok=True)

OUT_FILE = OUT_FOLDER / "presales_status.html"


def parse_logfile(path: Path):
    """
    Parse LAST run result for a given SD log.
    We only use the LAST START → LAST FINISHED → LAST ERROR (if any)
    """
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()

    start = None
    end = None
    errors = []

    for line in lines:
        if "START " in line:
            start = line.split("]")[0].strip("[")
        if "FINISHED" in line:
            end = line.split("]")[0].strip("[")
        if "ERROR" in line:
            errors.append(line)

    status = "OK" if not errors else "ERROR"

    return {
        "sd": path.stem,
        "status": status,
        "start": start,
        "end": end,
        "errors": errors[-10:],   # show last 10 errors only
    }


def build_dashboard():
    logs = list(LOG_FOLDER.glob("*.log"))

    # skip global run logs
    sd_logs = [x for x in logs if not x.name.startswith("run_")]

    rows = [parse_logfile(log) for log in sd_logs]

    html = """
<html>
<head>
<meta http-equiv="refresh" content="5">
<style>
body { font-family: Arial; }
table { width: 100%; border-collapse: collapse; }
th { background: #222; color: white; padding: 8px; }
td { border: 1px solid #ccc; padding: 6px; vertical-align: top; }
.ok { background: #c8f7c5; }
.err { background: #f6b2b2; }
.unk { background: #f8e59a; }
</style>
</head>
<body>

<h1>Presales Processing Dashboard</h1>
<p>Auto-refresh every 5 seconds</p>

<table>
<tr>
  <th>SD File</th>
  <th>Status</th>
  <th>Start</th>
  <th>End</th>
  <th>Errors</th>
</tr>
"""

    for row in rows:
        cls = "ok" if row["status"] == "OK" else "err"
        err_html = "<br>".join(row["errors"]) if row["errors"] else ""
        html += f"""
<tr class="{cls}">
  <td>{row['sd']}</td>
  <td>{row['status']}</td>
  <td>{row['start'] or ''}</td>
  <td>{row['end'] or ''}</td>
  <td>{err_html}</td>
</tr>
"""

    html += """
</table>
</body>
</html>
"""

    OUT_FILE.write_text(html, encoding="utf-8")
    print(f"✔ Dashboard created → {OUT_FILE}")


if __name__ == "__main__":
    build_dashboard()