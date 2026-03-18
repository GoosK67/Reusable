import mammoth
from pathlib import Path
import sys

# Prevent Windows CP1252 console crashes
sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
sys.stderr.reconfigure(encoding="utf-8", errors="ignore")

if __name__ == "__main__":
    sd_file = Path(sys.argv[1])

    out_folder = Path("extracted_html")
    out_folder.mkdir(exist_ok=True)

    out_file = out_folder / f"{sd_file.stem}.html"

    try:
        with open(sd_file, "rb") as f:
            result = mammoth.convert_to_html(f)
            html = result.value
    except Exception as e:
        print("EXTRACT ERROR:", e)
        sys.exit(1)

    try:
        out_file.write_text(html, encoding="utf-8", errors="ignore")
    except Exception as e:
        print("WRITE ERROR:", e)
        sys.exit(1)

    # IMPORTANT: no unicode icons!!
    print(f"HTML saved: {out_file.name}")