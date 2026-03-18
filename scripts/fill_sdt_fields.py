import json
import xml.etree.ElementTree as ET
from pathlib import Path

TEMPLATE_XML = Path("presales_template_sdt.xml")
JSON_FILE = Path("sections.json")
OUTPUT_XML = Path("presales_filled.xml")

def load_sdt_fields(xml_root):
    fields = {}
    for sdt in xml_root.findall(".//w:sdt", namespaces=ns):
        tag = sdt.find(".//w:tag", namespaces=ns)
        if tag is not None:
            name = tag.attrib.get(f"{{{ns['w']}}}val")
            fields[name] = sdt
    return fields

def set_sdt_text(sdt_node, text):
    content = sdt_node.find(".//w:sdtContent", namespaces=ns)
    for child in list(content):
        content.remove(child)
    p = ET.SubElement(content, f"{{{ns['w']}}}p")
    r = ET.SubElement(p, f"{{{ns['w']}}}r")
    t = ET.SubElement(r, f"{{{ns['w']}}}t")
    t.text = text

ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

def auto_map(json_data):
    mapping = {}
    text_by_heading = {k.lower(): v for k,v in json_data.items()}

    def pick(*candidates):
        for c in candidates:
            for k in text_by_heading:
                if c in k:
                    return text_by_heading[k]
        return ""

    return {
        "PRODUCT_SUMMARY": pick("service introduction", "service description", "overview"),
        "CLIENT_NEEDS": pick("needs", "client"),
        "PRODUCT_DESCRIPTION": pick("product"),
        "ARCHITECTURAL_DESCRIPTION": pick("technical implementation"),
        "KEY_FEATURES": pick("features", "functionalities"),
        "SCOPE": pick("scope"),
        "REQUIREMENTS": pick("prerequisites", "eligibility"),
        "VALUE_PROPOSITION": pick("value", "benefits"),
        "DIFFERENTIATORS": pick("differentiators"),
        "TRANSITION_TRANSFORMATION": pick("transition"),
        "CLIENT_RESPONSIBILITIES": pick("responsibilities"),
        "OPERATIONAL_SUPPORT": pick("support"),
        "TERMS_CONDITIONS": pick("conditions"),
        "ASSUMPTIONS_RISKS": pick("risks", "assumptions"),
        "ACCEPTANCE_CRITERIA": pick("acceptance"),
        "SLA_KPI": pick("sla"),
        "PRICING_ELEMENTS": pick("pricing"),
        "COST_ONE_TIME": pick("one time"),
        "COST_RECURRING": pick("recurring"),
        "CHARGING_MECHANISM": pick("charging"),
        "OTHER_DOCS": pick("documents"),
        "COMMERCIAL_SHEET": pick("solution sheet"),
        "SERVICE_DESCRIPTION_LINK": pick("service description")
    }

def main():
    xml_tree = ET.parse(TEMPLATE_XML)
    root = xml_tree.getroot()
    fields = load_sdt_fields(root)

    data = json.loads(JSON_FILE.read_text())
    mapping = auto_map(data)

    for field, text in mapping.items():
        if field in fields:
            set_sdt_text(fields[field], text)

    xml_tree.write(OUTPUT_XML, encoding="utf-8", xml_declaration=True)

if __name__ == "__main__":
    main()