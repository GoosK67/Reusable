from pathlib import Path
p = Path("scripts/generate_dashboard.py")
txt = p.read_text(encoding="utf-8")
txt = txt.replace(
    'if "XML\u2192DOCX OK" in line:',
    'if "XML" in line and "DOCX OK" in line:',
)
txt = txt.replace(
    'print(f"\u2714 Dashboard created \u2192 {OUT_FILE}")',
    'print(f"Dashboard created: {OUT_FILE}")',
)
p.write_text(txt, encoding="utf-8")
print("Done")
