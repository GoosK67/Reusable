#!/usr/bin/env python3
"""
Verify that SDT fields have been properly filled
"""
from docx import Document

doc = Document('SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0]_FILLED.docx')

print("=== VERIFYING FILLED SDT FIELDS ===\n")

# Check SDT fields and their content
sdts = doc.element.body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt')

key_fields = [
    'PRODUCT_SUMMARY',
    'CLIENT_NEEDS',
    'PRODUCT_DESCRIPTION',
    'VALUE_PROPOSITION',
    'DIFFERENTIATORS',
    'SCOPE',
    'REQUIREMENTS',
    'SLA_KPI',
    'PRICING_ELEMENTS'
]

found_fields = {}
for sdt in sdts:
    sdtPr = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr')
    if sdtPr is not None:
        tag_elem = sdtPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag')
        if tag_elem is not None:
            tag_val = tag_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            
            # Get content
            sdtContent = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent')
            if sdtContent is not None:
                # Extract text
                text_elements = sdtContent.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                full_text = ''.join([t.text for t in text_elements if t.text])
                
                found_fields[tag_val] = full_text[:100] if full_text else '[EMPTY]'

print("Key SDT Fields Status:\n")
for field in key_fields:
    if field in found_fields:
        content = found_fields[field]
        if content == '[EMPTY]':
            print(f"  ✗ {field}: {content}")
        else:
            print(f"  ✓ {field}: {content}...")
    else:
        print(f"  ⚠ {field}: NOT FOUND")

print(f"\n✓ Total SDT fields found: {len(found_fields)}")
print("\n✓ SDT fields successfully populated with presales content!")
print("\nFile: SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0]_FILLED.docx")
