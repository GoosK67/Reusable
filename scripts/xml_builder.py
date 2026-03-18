from pathlib import Path
import sys
import json
from datetime import datetime

# ---------------------------------------------------------
# LOGGING (always append)
# ---------------------------------------------------------
LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)

def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"

    with open(logfile, "a", encoding="utf-8", errors="ignore") as f:
        f.write(line)

    print(line, end="")

# ---------------------------------------------------------
# Convert JSON → XML structure
# ---------------------------------------------------------
def to_xml(mapped: dict) -> str:
    """
    Very simple XML generator for SD templates.
    Produces:
      <ServiceDescription>
         <Section name="...">
             <Header>...</Header>
             <Category>...</Category>
             <Content>...</Content>
         </Section>
      </ServiceDescription>
    """

    xml = ['<ServiceDescription>']

    for header, entry in mapped.items():
        xml.append(f'  <Section name="{header}">')
        xml.append(f'    <Header>{header}</Header>')
        xml.append(f'    <Category>{entry.get("category","")}</Category>')
        xml.append(f'    <Content><![CDATA[{entry.get("content","")}]]></Content>')
        xml.append(f'  </Section>')

    xml.append('</ServiceDescription>')
    return "\n".join(xml)

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    mapped_file = Path(sys.argv[1])
    sd_name = mapped_file.stem

    log(f"START xml_builder for: {mapped_file}", sd_name)

    try:
        mapped = json.loads(
            mapped_file.read_text(encoding="utf-8", errors="ignore")
        )

        xml_text = to_xml(mapped)

        out_folder = Path("output/xml")
        out_folder.mkdir(parents=True, exist_ok=True)

        out_file = out_folder / f"{sd_name}.xml"
        out_file.write_text(
            xml_text,
            encoding="utf-8",
            errors="ignore"
        )

        log(f"XML OK → {out_file}", sd_name)
        sys.exit(0)

    except Exception as e:
        log(f"XML ERROR: {e}", sd_name)
        sys.exit(1)