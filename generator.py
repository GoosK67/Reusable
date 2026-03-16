from docx import Document

def generate_presales(fields, template_path, output_path):
    doc = Document(template_path)

    def replace(tag, value):
        for p in doc.paragraphs:
            if tag in p.text:
                p.text = p.text.replace(tag, value)

    for key, value in fields.items():
        replace(f"<{key}>", value)

    doc.save(output_path)