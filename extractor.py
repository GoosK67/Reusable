from docx import Document

def extract_sd(path):
    doc = Document(path)

    data = {}
    current = None

    for p in doc.paragraphs:
        t = p.text.strip()

        if p.style.name.startswith("Heading"):
            current = t
            data[current] = ""
        elif current:
            data[current] += t + "\n"

    # Tabellen extraheren
    data["tables"] = []
    for table in doc.tables:
        rows = []
        for r in table.rows:
            rows.append([c.text.strip() for c in r.cells])
        data["tables"].append(rows)

    return data