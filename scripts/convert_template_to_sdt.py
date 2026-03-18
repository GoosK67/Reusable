import zipfile
from lxml import etree
from pathlib import Path
import re
import difflib
import datetime

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def log(msg):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

MAPPING = {
    "Product summary": "PRODUCT_SUMMARY",
    "Understanding the Client’s Needs": "CLIENT_NEEDS",
    "Product Description": "PRODUCT_DESCRIPTION",
    "Architectural description": "ARCHITECTURAL_DESCRIPTION",
    "Key features and functionalities": "KEY_FEATURES",
    "Scope / out-of-scope": "SCOPE",
    "Requirements and Prerequisites": "REQUIREMENTS",
    "Value Proposition": "VALUE_PROPOSITION",
    "Key Differentiators": "DIFFERENTIATORS",
    "Transition & Transformation": "TRANSITION_TRANSFORMATION",
    "Client responsibilities": "CLIENT_RESPONSIBILITIES",
    "Operational Support": "OPERATIONAL_SUPPORT",
    "Terms and Conditions": "TERMS_CONDITIONS",
    "Assumptions & Risks": "ASSUMPTIONS_RISKS",
    "Acceptance criteria": "ACCEPTANCE_CRITERIA",
    "SLA & KPI Management": "SLA_KPI",
    "Cost/Pricing elements": "PRICING_ELEMENTS",
    "One time cost elements": "COST_ONE_TIME",
    "Recurring costing elements": "COST_RECURRING",
    "Charging mechanism": "CHARGING_MECHANISM",
    "Service description": "SERVICE_DESCRIPTION_LINK"
}

def clean_text(t):
    if not t:
        return ""
    t = re.sub(r"<[^>]+>", "", t)
    t = t.replace("_", "").replace("*", "")
    return t.strip().lower()

def fuzzy_match(a, b):
    return difflib.SequenceMatcher(None, clean_text(a), clean_text(b)).ratio() >= 0.55

def build_sdt_block(correct_tag, original_p):
    p_copy = etree.fromstring(etree.tostring(original_p))

    sdt = etree.Element("{%s}sdt" % NS["w"])
    sdtPr = etree.SubElement(sdt, "{%s}sdtPr" % NS["w"])

    tag = etree.SubElement(sdtPr, "{%s}tag" % NS["w"])
    tag.set("{%s}val" % NS["w"], correct_tag)

    lock_el = etree.SubElement(sdtPr, "{%s}lock" % NS["w"])
    lock_el.set("{%s}val" % NS["w"], "sdtContentLocked")

    content = etree.SubElement(sdt, "{%s}sdtContent" % NS["w"])
    content.append(p_copy)

    return sdt

def unwrap_existing_sdt_blocks(doc_xml):
    # Only move children of w:sdtContent back to the parent.
    # Moving w:sdtPr/w:sdtContent themselves to body level makes the document invalid.
    for sdt in doc_xml.xpath("//w:sdt", namespaces=NS):
        parent = sdt.getparent()
        if parent is None:
            continue

        insert_at = parent.index(sdt)
        sdt_content = sdt.find("w:sdtContent", namespaces=NS)

        if sdt_content is not None:
            for child in list(sdt_content):
                sdt_content.remove(child)
                parent.insert(insert_at, child)
                insert_at += 1

        parent.remove(sdt)

def convert_v4():
    in_path = Path("templates/NEW presales_template_sdt.docx")
    out_path = Path("templates/presales_template_sdt_GENERATED.docx")

    if not in_path.exists():
        log("FOUT: template niet gevonden.")
        return

    log(f"Template gevonden: {in_path}")

    with zipfile.ZipFile(in_path, "r") as zin:
        buffer = {n: zin.read(n) for n in zin.namelist()}

    doc_xml = etree.fromstring(buffer["word/document.xml"])

    # verwijder alle bestaande SDTs (reset)
    unwrap_existing_sdt_blocks(doc_xml)

    log("Alle bestaande SDTs verwijderd (clean reset).")

    # Rebuild paragraph list after unwrapping SDTs.
    paragraphs = doc_xml.xpath("//w:p", namespaces=NS)

    vul_positions = []
    for i, p in enumerate(paragraphs):
        text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
        if "VUL_HIER_IN" in text:
            vul_positions.append(i)

    used = set()
    replacements = 0

    for title, correct_tag in MAPPING.items():
        log(f"Zoek titel: {title}")

        title_pos = None

        for i, p in enumerate(paragraphs):
            txt = "".join(p.xpath(".//w:t/text()", namespaces=NS))
            if fuzzy_match(title, txt):
                title_pos = i
                break

        if title_pos is None:
            log(f"!! Geen match voor {title}")
            continue

        vul_target = None
        for j in range(title_pos + 1, len(paragraphs)):
            txt = "".join(paragraphs[j].xpath(".//w:t/text()", namespaces=NS))
            if "VUL_HIER_IN" in txt and j not in used:
                vul_target = j
                used.add(j)
                break

        if vul_target is None:
            log(f"!! Geen VUL_HIER_IN onder {title}")
            continue

        original_p = paragraphs[vul_target]
        new_sdt = build_sdt_block(correct_tag, original_p)

        parent = original_p.getparent()
        parent.replace(original_p, new_sdt)

        log(f"   ✓ SDT toegevoegd: {correct_tag} @ index {vul_target}")
        replacements += 1

    buffer["word/document.xml"] = etree.tostring(doc_xml, encoding="UTF-8", xml_declaration=True)

    with zipfile.ZipFile(out_path, "w") as zout:
        for name, data in buffer.items():
            zout.writestr(name, data)

    log(f"CONVERSIE AFGEWERKT — {replacements} SDTs aangemaakt.")
    log(f"Nieuw SDT-template: {out_path}")

if __name__ == "__main__":
    convert_v4()