import subprocess
import sys
from pathlib import Path
import json
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import threading
import os

# ---------------------------------------------------------
# FORCE UTF‑8 OUTPUT (NO MORE UNICODE CRASHES)
# ---------------------------------------------------------
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
sys.stderr.reconfigure(encoding="utf-8", errors="ignore")

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------

BASE = Path(__file__).parent.resolve()

SD_ROOT = Path(r"C:\Users\koengo\Cegeka\Product Management - Product Management Library")
SCRIPTS = BASE / "scripts"
TEMPLATE_XML = BASE / "templates" / "presales_template_sdt.xml"
TEMPLATE_DOCX = BASE / "templates" / "presales_template_sdt.docx"
OUTPUT = BASE / "presales"
LOG_FOLDER = BASE / "log"
MAPPED = BASE / "mapped"
DASHBOARD = BASE / "dashboard"

ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

LOG_FOLDER.mkdir(exist_ok=True)
MAPPED.mkdir(exist_ok=True)
DASHBOARD.mkdir(exist_ok=True)

RUN_LOG = LOG_FOLDER / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
LOCK = threading.Lock()

MAX_THREADS = 4


# ---------------------------------------------------------
# SAFE UTF‑8 LOGGING (APPEND‑MODE, NEVER read_text)
# ---------------------------------------------------------

def log(msg, sd_file=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tid = threading.get_ident()
    line = f"[{timestamp}] [TID:{tid}] {msg}\n"

    with LOCK:
        # Global log
        with open(RUN_LOG, "a", encoding="utf-8", errors="ignore") as f:
            f.write(line)

        # SD log
        if sd_file:
            sd_log = LOG_FOLDER / f"{sd_file.stem}.log"
            with open(sd_log, "a", encoding="utf-8", errors="ignore") as f:
                f.write(line)

    print(line, end="", flush=True)


# ---------------------------------------------------------
# DEBUG SUBPROCESS RUNNER
# ---------------------------------------------------------

def run_debug(command, sd_file, step_name):
    log(f"[DEBUG] Running step: {step_name}", sd_file)
    log(f"[DEBUG] Command: {command}", sd_file)

    result = subprocess.run(command, capture_output=True, text=True)

    if result.stdout:
        log(f"[DEBUG][stdout] {result.stdout}", sd_file)
    if result.stderr:
        log(f"[DEBUG][stderr] {result.stderr}", sd_file)

    if result.returncode != 0:
        raise Exception(f"[{step_name}] FAILED with exit code {result.returncode}")

    log(f"[DEBUG] Step '{step_name}' OK", sd_file)


# ---------------------------------------------------------
# XML HELPERS
# ---------------------------------------------------------

def load_sdt_fields(xml_root):
    fields = {}
    for sdt in xml_root.findall(".//w:sdt", ns):
        tag = sdt.find(".//w:tag", ns)
        if tag is not None:
            fields[tag.attrib[f"{{{ns['w']}}}val"]] = sdt
    return fields


def set_sdt_text(sdt_node, text):
    content = sdt_node.find(".//w:sdtContent", ns)
    for child in list(content):
        content.remove(child)

    p = ET.SubElement(content, f"{{{ns['w']}}}p")
    r = ET.SubElement(p, f"{{{ns['w']}}}r")
    t = ET.SubElement(r, f"{{{ns['w']}}}t")
    t.text = text


def auto_map(sections):
    txt = {k.lower(): v for k, v in sections.items()}

    def pick(*keys):
        for k in keys:
            for head in txt:
                if k in head:
                    return txt[head]
        return ""

    return {
        "PRODUCT_SUMMARY": pick("service introduction", "service overview"),
        "CLIENT_NEEDS": pick("needs"),
        "PRODUCT_DESCRIPTION": pick("product"),
        "ARCHITECTURAL_DESCRIPTION": pick("technical implementation"),
        "KEY_FEATURES": pick("features"),
        "SCOPE": pick("scope"),
        "REQUIREMENTS": pick("eligibility"),
        "VALUE_PROPOSITION": pick("value"),
        "DIFFERENTIATORS": pick("differentiators"),
        "OPERATIONAL_SUPPORT": pick("support"),
        "TERMS_CONDITIONS": pick("conditions"),
        "ASSUMPTIONS_RISKS": pick("risks"),
        "PRICING_ELEMENTS": pick("pricing"),
        "SERVICE_DESCRIPTION_LINK": pick("service description"),
    }


# ---------------------------------------------------------
# DOCX BUILDER
# ---------------------------------------------------------

def generate_docx(xml_path, out_docx):
    with zipfile.ZipFile(TEMPLATE_DOCX, "r") as zin:
        with zipfile.ZipFile(out_docx, "w") as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, xml_path.read_bytes())
                else:
                    zout.writestr(item, zin.read(item.filename))


# ---------------------------------------------------------
# PROCESS ONE SD
# ---------------------------------------------------------

def process(sd):
    try:
        log(f"START {sd.name}", sd)

        # 1) Extract HTML
        html_file = Path("extracted_html") / f"{sd.stem}.html"
        run_debug([sys.executable, str(SCRIPTS/"extract_html.py"), str(sd)], sd, "extract_html")

        # 2) Parse HTML
        json_file = Path("output/json") / f"{sd.stem}.json"
        run_debug([sys.executable, str(SCRIPTS/"parse_html_sections.py"), str(html_file)], sd, "parse_html")

        # 3) Auto‑map → mapped/sections_x.json
        sd_map = MAPPED / f"sections_{sd.stem}.json"
        run_debug([sys.executable, str(SCRIPTS/"auto_map_sections.py"), str(json_file)], sd, "auto_map")

        # 4) Load mapping
        sections = json.loads(sd_map.read_text(encoding="utf-8", errors="ignore"))
        mapping = auto_map(sections)

        # 5) Fill XML
        xml_tree = ET.parse(TEMPLATE_XML)
        fields = load_sdt_fields(xml_tree.getroot())

        for tag, text in mapping.items():
            if tag in fields:
                set_sdt_text(fields[tag], text)

        filled_xml = OUTPUT / f"{sd.stem}_filled.xml"
        xml_tree.write(filled_xml, encoding="utf-8")

        # 6) Build DOCX
        out_docx = OUTPUT / f"{sd.stem}_presales.docx"
        generate_docx(filled_xml, out_docx)

        log(f"FINISHED {sd.name}", sd)
        return f"OK {sd.name}"

    except Exception as e:
        log(f"ERROR {sd.name}: {e}", sd)
        return f"ERROR {sd.name}: {e}"


# ---------------------------------------------------------
# DASHBOARD GENERATOR
# ---------------------------------------------------------

def generate_dashboard():
    dashboard_script = SCRIPTS / "generate_dashboard.py"
    if dashboard_script.exists():
        log("Generating dashboard...", None)
        subprocess.run(
            [sys.executable, str(dashboard_script)],
            capture_output=True, text=True
        )


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------

def main():
    OUTPUT.mkdir(exist_ok=True)

    log(f"Scanning SD directory: {SD_ROOT}")
    sd_files = list(SD_ROOT.rglob("SD*.docx"))

    log(f"Found {len(sd_files)} SD files")

    results = []
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as pool:
        futures = [pool.submit(process, sd) for sd in sd_files]
        for f in as_completed(futures):
            results.append(f.result())

    # Summary
    for r in results:
        log(r)

    # Dashboard refresh
    generate_dashboard()

    log("🚀 FULL RUN COMPLETED (UTF‑8 SAFE + DASHBOARD FIXED)")


if __name__ == "__main__":
    main()