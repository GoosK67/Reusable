from pathlib import Path
import subprocess
import json

EXTRACTED = Path("extracted")
OUTPUT = Path("output")
OUTPUT.mkdir(exist_ok=True)

def call_copilot_prompt(json_data, prompt_path):
    cmd = [
        "gh", "copilot", "prompt", "apply",
        "--prompt-file", prompt_path,
        "--input", json.dumps(json_data)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.stdout

for js in EXTRACTED.glob("*.json"):
    with open(js, "r", encoding="utf-8") as f:
        data = json.load(f)

    guide = call_copilot_prompt(data, ".github/prompts/generate-presales-guide.prompt.md")
    md_out = OUTPUT / f"{js.stem}_presales.md"
    md_out.write_text(guide, encoding="utf-8")

    print(f"Generated: {md_out}")