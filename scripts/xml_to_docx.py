import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree
from datetime import datetime
import difflib
import re

LOG_FOLDER = Path("log")
LOG_FOLDER.mkdir(exist_ok=True)


def log(msg, sd_name="GENERAL"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    logfile = LOG_FOLDER / f"{sd_name}.log"
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(line)
    print(line, end="")


# =====================================================================
#  PERFECTE MAPPING (op basis van jouw XML-sectienamen)  [1](https://cegekagroup-my.sharepoint.com/personal/koen_goos_cegeka_com1/Documents/Microsoft%20Copilot%20Chat%20Files/run_all.py)%20%5BPRD.0.L.135%5D%5BPV0.8%5D%5BDV1.0%5D_mapped.xml)
# =====================================================================
MAPPING = {
    "service introduction": "PRODUCT_SUMMARY",
    "service identification": "PRODUCT_SUMMARY",
    "service reporting": "PRODUCT_SUMMARY",
    "service window": "PRODUCT_SUMMARY",
    "service overview": "CLIENT_NEEDS",
    "goals": "CLIENT_NEEDS",
    "service target audience": "CLIENT_NEEDS",

    "technical implementation": "ARCHITECTURAL_DESCRIPTION",

    "key features": "KEY_FEATURES",
    "operational readiness": "KEY_FEATURES",
    "run services": "KEY_FEATURES",
    "management services": "KEY_FEATURES",
    "governance & reporting": "KEY_FEATURES",
    "process": "KEY_FEATURES",

    "scope": "SCOPE",
    "out_of_scope": "SCOPE",
    "out of scope": "SCOPE",

    "eligibility & prerequisites": "REQUIREMENTS",

    "value proposition": "VALUE_PROPOSITION",
    "value & benefits": "VALUE_PROPOSITION",

    "differentiators": "DIFFERENTIATORS",

    "transition & transformation": "TRANSITION_TRANSFORMATION",

    "client responsibilities": "CLIENT_RESPONSIBILITIES",

    "support model": "OPERATIONAL_SUPPORT",

    "conditions": "TERMS_CONDITIONS",

    "risks": "ASSUMPTIONS_RISKS",
    "assumptions": "ASSUMPTIONS_RISKS",

    "acceptance": "ACCEPTANCE_CRITERIA",

    "service level": "SLA_KPI",

    "pricing": "PRICING_ELEMENTS",
    "delivery model": "PRICING_ELEMENTS",

    "service dependencies": "SERVICE_DESCRIPTION_LINK",
    "service description link": "SERVICE_DESCRIPTION_LINK",
}


# =====================================================================
#  SEMANTISCHE MATCHING ENGINE
# =====================================================================

def normalize(s):
    return s.strip().lower().replace("_", " ").replace("&", "and")


def fuzzy_match(a, b):
    return difflib.SequenceMatcher(None, normalize(a), normalize(b)).ratio() >= 0.70


def sanitize_xml_text(text):
    # Remove XML 1.0 illegal control chars that can make Word parts unreadable.
    if text is None:
        return ""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def _set_sdt_text_preserving_structure(content, new_text, ns):
    text_nodes = content.xpath(".//w:t", namespaces=ns)
    safe_text = sanitize_xml_text(new_text).strip()

    if text_nodes:
        text_nodes[0].text = safe_text
        for node in text_nodes[1:]:
            node.text = ""
        return

    wp = ns["w"]
    first_child = next(iter(content), None)
    first_tag = etree.QName(first_child.tag).localname if first_child is not None else None

    if first_tag == "r":
        r = etree.Element(f"{{{wp}}}r")
        t = etree.SubElement(r, f"{{{wp}}}t")
        t.text = safe_text
        content.append(r)
        return

    p = etree.Element(f"{{{wp}}}p")
    r = etree.SubElement(p, f"{{{wp}}}r")
    t = etree.SubElement(r, f"{{{wp}}}t")
    t.text = safe_text
    content.append(p)


def replace_sdt(xml_root, tag_name, new_text):
    NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    sdts = xml_root.xpath(f".//w:sdt[w:sdtPr/w:tag[@w:val='{tag_name}']]", namespaces=NS)

    count = 0
    for sdt in sdts:
        content = sdt.find("w:sdtContent", NS)
        if content is None:
            continue

        _set_sdt_text_preserving_structure(content, new_text, NS)
        count += 1

    return count


def process_docx(buffer, xml_root_data, sd_name):
    total = 0
    xml_sections = xml_root_data.xpath("//Section/@name")

    for part_name, xml_bytes in list(buffer.items()):

        # Only Word XML parts
        if not (part_name.startswith("word/") and part_name.endswith(".xml")):
            continue

        try:
            xml_root = etree.fromstring(xml_bytes)
        except:
            continue

        # Try to match each XML section semantically
        for section_xml in xml_sections:
            xml_norm = normalize(section_xml)

            matched_key = None
            for key in MAPPING.keys():
                if fuzzy_match(xml_norm, key):
                    matched_key = key
                    break

            if not matched_key:
                continue

            sdt_tag = MAPPING[matched_key]

            element = xml_root_data.xpath(f"//Section[@name='{section_xml}']")
            if not element:
                continue

            content = element[0].findtext("Content", default="").strip()
            if not content:
                continue

            replaced = replace_sdt(xml_root, sdt_tag, content)
            total += replaced

            if replaced > 0:
                log(f"Filled SDT '{sdt_tag}' with XML section '{section_xml}'", sd_name)

        # Replace part in buffer
        buffer[part_name] = etree.tostring(xml_root, encoding="UTF-8", xml_declaration=True)

    return total


# =====================================================================
#  MAIN — met correcte ZIP-merging fix (Word compatible!)
# =====================================================================
if __name__ == "__main__":
    xml_file = Path(sys.argv[1])
    template = Path("templates/presales_template_sdt_GENERATED.docx")
    sd_name = xml_file.stem

    log(f"START xml_to_docx_v3_fixed for: {xml_file}", sd_name)

    if not template.exists():
        log(f"ERROR: Template not found: {template}", sd_name)
        sys.exit(1)

    xml_tree = etree.parse(str(xml_file))
    xml_root = xml_tree.getroot()

    with ZipFile(template, "r") as zin:
        original = {n: zin.read(n) for n in zin.namelist()}

    # make a buffer we can modify
    buffer = dict(original)

    # process all xml parts
    total_filled = process_docx(buffer, xml_root, sd_name)
    log(f"TOTAL SDT FIELDS FILLED: {total_filled}", sd_name)

    # SAFE DOCX REBUILD FIX
    out_folder = Path("output/docx")
    out_folder.mkdir(exist_ok=True)
    out_file = out_folder / f"{sd_name}_FINAL.docx"

    with ZipFile(template, "r") as zin:
       with ZipFile(out_file, "w") as zout:
           for info in zin.infolist():
               name = info.filename
               if name in buffer:
                   zout.writestr(info, buffer[name])
               else:
                   zout.writestr(info, zin.read(name))


    log(f"XML→DOCX OK → {out_file}", sd_name)
    sys.exit(0)