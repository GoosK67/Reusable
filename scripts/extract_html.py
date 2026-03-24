from pathlib import Path
import sys
import os
import mammoth
from datetime import datetime

# -----------------------------------------
# LOGGING (ALTIJD APPEND)
# -----------------------------------------
LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)

def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"

    # ALWAYS APPEND
    with open(logfile, "a", encoding="utf-8", errors="ignore") as f:
        f.write(line)

    print(line, end="")


def normalize_windows_path(path: Path) -> str:
    """Return a Windows path string that supports long paths when needed."""
    p = str(path)
    if os.name != "nt":
        return p

    # Keep existing extended paths untouched.
    if p.startswith("\\\\?\\"):
        return p

    # Use long-path prefix for deep paths.
    if len(p) >= 248:
        return "\\\\?\\" + p

    return p


def path_exists(path: Path) -> bool:
    """Check existence using regular and long-path forms."""
    if path.exists():
        return True
    try:
        return Path(normalize_windows_path(path)).exists()
    except Exception:
        return False

# -----------------------------------------
# MAIN EXTRACTOR
# -----------------------------------------
if __name__ == "__main__":
    sd_file = Path(sys.argv[1])
    sd_name = sd_file.stem

    log(f"START extract_html for: {sd_file}", sd_name)

    out_folder = Path("extracted_html")
    out_folder.mkdir(exist_ok=True)

    out_file = out_folder / f"{sd_file.stem}.html"

    try:
        if not path_exists(sd_file):
            raise FileNotFoundError(f"Source not accessible: {sd_file}")

        source_path = normalize_windows_path(sd_file)

        with open(source_path, "rb") as f:
            result = mammoth.convert_to_html(f)
            html = result.value

        out_file.write_text(html, encoding="utf-8", errors="ignore")

        log(f"HTML saved: {out_file}", sd_name)
        log("extract_html OK", sd_name)

        sys.exit(0)

    except Exception as e:
        log(f"EXTRACT ERROR: {e}", sd_name)
        sys.exit(1)