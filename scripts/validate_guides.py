from pathlib import Path
import json
import subprocess

OUTPUT = Path("output")
EXTRACTED = Path("extracted")

def call_prompt(input_text, json_data):
    cmd = [
        "gh", "copilot", "prompt", "apply",
        "--prompt-file", ".github/prompts/qa-review.prompt.md",
        "--input", json.dumps({"guide": input_text, "json": json_data})
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    return res.stdout

for md in OUTPUT.glob("*_presales.md"):
    js = EXTRACTED / (md.stem.replace("_presales", "") + ".json")
    guide = md.read_text()
    data = json.loads(js.read_text())
    review = call_prompt(guide, data)

    review_path = OUTPUT / f"{md.stem}_review.txt"
    review_path.write_text(review)
    print(f"Reviewed: {review_path}")