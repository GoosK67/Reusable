import sys
from pathlib import Path
from zipfile import ZipFile
from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def scan_xml_part(part_name, xml_bytes):
    """Scan één XML-part voor SDTs."""
    results = []

    # 1. Probeer XML te parsen
    try:
        root = etree.fromstring(xml_bytes)
        sdts = root.xpath(".//w:sdt", namespaces=NS)

        for sdt in sdts:
            tag = sdt.xpath(".//w:sdtPr/w:tag/@w:val", namespaces=NS)
            tag_val = tag[0] if tag else "(GEEN TAG)"
            snippet = etree.tostring(sdt, encoding="unicode")[:300]
            results.append((part_name, "XML", tag_val, snippet))
        return results

    except Exception:
        # 2. RAW fallback: zoek “w:sdt” in plain text (soms zitten SDTs in non‑OOXML parts!)
        text = xml_bytes.decode(errors="ignore")
        if "w:sdt" in text or "<sdt" in text:
            idx = text.find("w:sdt")
            snippet = text[idx-50:idx+300]
            results.append((part_name, "RAW", "(UNKNOWN TAG)", snippet))
        return results


def scan_docx(docx_path):
    print(f"\n=== DEBUG SDT LOCATOR ===")
    print(f"Scanning: {docx_path}\n")

    with ZipFile(docx_path, "r") as zin:
        for name in zin.namelist():
            # Alleen Word parts bekijken
            if not name.startswith("word/"):
                continue

            data = zin.read(name)
            found = scan_xml_part(name, data)

            for part_name, mode, tag_val, snippet in found:
                print("\n--------------------------------------")
                print(f"FOUND SDT in: {part_name}")
                print(f"Parse mode : {mode}")
                print(f"SDT Tag    : {tag_val}")
                print("Snippet:")
                print(snippet)
                print("--------------------------------------\n")

    print("\n=== SCAN COMPLETE ===\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Gebruik: python debug_sdt_locator.py <template.docx>")
        sys.exit(1)

    scan_docx(sys.argv[1])