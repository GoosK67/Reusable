from docx import Document


def _replace_in_paragraph(paragraph, replacements):
    updated = paragraph.text
    for tag, value in replacements.items():
        updated = updated.replace(tag, value)

    if updated != paragraph.text:
        paragraph.text = updated


def _replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            _replace_in_container(cell, replacements)


def _replace_in_container(container, replacements):
    for paragraph in container.paragraphs:
        _replace_in_paragraph(paragraph, replacements)

    for table in container.tables:
        _replace_in_table(table, replacements)


def fill_template(template_path, output_path, fields):
    """
    Fill a DOCX template by replacing predefined tags with provided values.

    Required tags:
    <ProductSummary>, <ValueProposition>, <ProductDescription>,
    <Requirements>, <Scope>, <SLA>, <OperationalSupport>
    """
    doc = Document(template_path)

    replacements = {
        "<ProductSummary>": str(fields.get("ProductSummary", "") or ""),
        "<ValueProposition>": str(fields.get("ValueProposition", "") or ""),
        "<ProductDescription>": str(fields.get("ProductDescription", "") or ""),
        "<Requirements>": str(fields.get("Requirements", "") or ""),
        "<Scope>": str(fields.get("Scope", "") or ""),
        "<SLA>": str(fields.get("SLA", "") or ""),
        "<OperationalSupport>": str(fields.get("OperationalSupport", "") or ""),
    }

    _replace_in_container(doc, replacements)

    for section in doc.sections:
        _replace_in_container(section.header, replacements)
        _replace_in_container(section.footer, replacements)

    doc.save(output_path)


def generate_presales(fields, template_path, output_path):
    doc = Document(template_path)

    def replace(tag, value):
        for p in doc.paragraphs:
            if tag in p.text:
                p.text = p.text.replace(tag, value)

    for key, value in fields.items():
        replace(f"<{key}>", value)

    doc.save(output_path)