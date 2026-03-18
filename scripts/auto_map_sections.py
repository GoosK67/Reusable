import json
from pathlib import Path
import sys
import re

if __name__ == "__main__":
    src_json = Path(sys.argv[1])   # EXACT 1 json input

    out_dir = Path("mapped")
    out_dir.mkdir(exist_ok=True)

    out_json = out_dir / f"sections_{src_json.stem}.json"

    data = json.loads(src_json.read_text(encoding="utf-8"))
    result = {}

    def normalize(x):
        x = x.lower().replace("_", " ")
        return re.sub(r"\s+", " ", x).strip()

    for heading, content in data.items():
        result[normalize(heading)] = content.strip()

    out_json.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"✔ Mapped → {out_json}")
``