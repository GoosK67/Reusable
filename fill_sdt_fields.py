#!/usr/bin/env python3
"""
Fill the SDT template by updating Content Control (SDT) fields
Maps presales markdown sections to SDT tags
"""
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import re

# Read the presales markdown
with open('presales/Presales Guide - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0].md', 'r', encoding='utf-8') as f:
    md_content = f.read()

# Extract sections from markdown
sections = {}
current_section = None
current_subsection = None
current_content = []

lines = md_content.split('\n')
for line in lines:
    if line.startswith('## '):
        if current_section and current_content:
            if current_subsection:
                if not isinstance(sections.get(current_section), dict):
                    sections[current_section] = {}
                sections[current_section][current_subsection] = '\n'.join(current_content).strip()
            else:
                sections[current_section] = '\n'.join(current_content).strip()
        
        current_section = line[3:].strip()
        current_subsection = None
        current_content = []
    elif line.startswith('### '):
        if current_content:
            if current_subsection:
                if not isinstance(sections.get(current_section), dict):
                    sections[current_section] = {}
                sections[current_section][current_subsection] = '\n'.join(current_content).strip()
            elif current_section:
                if current_section not in sections:
                    sections[current_section] = '\n'.join(current_content).strip()
        
        if current_section and not isinstance(sections.get(current_section), dict):
            old_content = sections.get(current_section, '')
            sections[current_section] = {'_main': old_content} if old_content else {}
        
        current_subsection = line[4:].strip()
        current_content = []
    elif current_section or current_subsection:
        if line.strip():
            current_content.append(line)

if current_section and current_content:
    if current_subsection:
        if not isinstance(sections.get(current_section), dict):
            sections[current_section] = {}
        sections[current_section][current_subsection] = '\n'.join(current_content).strip()
    else:
        sections[current_section] = '\n'.join(current_content).strip()

print("✓ Extracted sections from markdown")

# Load template
doc = Document('templates/presales_template_sdt_v2.docx')
print("✓ Loaded template with SDT fields")

# Helper function to fill an SDT field
def fill_sdt_field(doc, sdt_tag, content_text):
    """Fill a content control (SDT) field with new text content"""
    if not content_text or not content_text.strip():
        return False
    
    # Find the SDT with the given tag
    sdts = doc.element.body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt')
    
    for sdt in sdts:
        sdtPr = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr')
        if sdtPr is not None:
            tag_elem = sdtPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag')
            if tag_elem is not None:
                tag_val = tag_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                
                if tag_val == sdt_tag:
                    # Found the right SDT, now clear and fill it
                    sdtContent = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent')
                    
                    if sdtContent is not None:
                        # Clear existing content except the first paragraph (keep formatting)
                        existing_paras = sdtContent.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                        for para in existing_paras[1:]:
                            sdtContent.remove(para)
                        
                        # Get the first paragraph or create one
                        if existing_paras:
                            first_para = existing_paras[0]
                            # Clear runs but keep paragraph
                            for run in first_para.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'):
                                first_para.remove(run)
                        else:
                            first_para = parse_xml(r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                            sdtContent.append(first_para)
                        
                        # Add content as paragraphs
                        content_lines = content_text.strip().split('\n')
                        is_first = True
                        
                        for line in content_lines:
                            line = line.rstrip()
                            if not line:
                                continue
                            
                            # Use first paragraph or create new ones
                            if is_first:
                                target_para = first_para
                                is_first = False
                            else:
                                target_para = parse_xml(r'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                                sdtContent.append(target_para)
                            
                            # Handle bullet points
                            if line.startswith('- '):
                                text_content = line[2:].strip()
                                # Add bullet formatting
                                pPr = parse_xml(
                                    r'<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                    r'<w:pStyle w:val="ListBullet"/>'
                                    r'</w:pPr>'
                                )
                                target_para.append(pPr)
                            else:
                                text_content = line
                            
                            # Add text as run
                            run = parse_xml(
                                r'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                r'<w:t/>'
                                r'</w:r>'
                            )
                            t_elem = run.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                            t_elem.text = text_content
                            target_para.append(run)
                        
                        return True
    
    return False

# Mapping of SDT tags to markdown sections
print("\nFilling SDT fields:")

chapter_1_content = {
    'PRESALES_INSTRUCTIONS': (
        "Presales voorbereiding voor 'Cegeka DBMS for Oracle on Azure'.\n"
        "Gebruik enkel gevalideerde SD-feiten in klantoffertes.\n"
        "Ontbrekende contractuele waarden expliciet markeren als [TO BE COMPLETED]."
    ),
    'CEGEKA_CONTACTS': (
        "- Presales Owner: [TO BE COMPLETED]\n"
        "- Solution Architect: [TO BE COMPLETED]\n"
        "- Service Delivery Contact: [TO BE COMPLETED]"
    ),
    'PRESALES_CHECKS': (
        "- Bevestig gekozen deployment model: Oracle Database@Azure, ODSA of Oracle on Azure VMs.\n"
        "- Bevestig scope: standaardservices vs. optionele services.\n"
        "- Bevestig verantwoordelijkheden voor licensing (BYOL of license-included).\n"
        "- Bevestig klantvereisten rond security/compliance en netwerktoegang.\n"
        "- Bevestig dat SLA-bijlage en finale contractwaarden beschikbaar zijn of markeer [TO BE COMPLETED]."
    ),
    'SKU_INFORMATION': (
        "Service: Cegeka DBMS for Oracle on Azure [PRD.0.8.001].\n"
        "Prijsopbouw volgens SD: managed service fee + cloud/licensing kosten.\n"
        "Definitieve SKU-combinaties en commerciële codering: [TO BE COMPLETED]."
    ),
    'OTHER_CONDITIONAL_SOLUTIONS': (
        "- Optionele diensten uit SD (o.a. geavanceerde HA/DR, major upgrades/migraties, consultatieve assessments).\n"
        "- Eventuele aanvullende oplossingen die verplicht samen verkocht worden: [TO BE COMPLETED]."
    ),
    'QA_CUSTOMERS': (
        "- Welk deployment model is vereist en waarom?\n"
        "- Welke RPO/RTO en beschikbaarheidsdoelen zijn contractueel nodig?\n"
        "- Welke compliance-eisen en security controls moeten toegepast worden?\n"
        "- Wie draagt licensing-verantwoordelijkheid (BYOL of inbegrepen model)?\n"
        "- Welke onboarding planning, scopegrenzen en acceptatiecriteria gelden?"
    ),
}

mapping = {
    'PRESALES_INSTRUCTIONS': chapter_1_content['PRESALES_INSTRUCTIONS'],
    'CEGEKA_CONTACTS': chapter_1_content['CEGEKA_CONTACTS'],
    'PRESALES_CHECKS': chapter_1_content['PRESALES_CHECKS'],
    'SKU_INFORMATION': chapter_1_content['SKU_INFORMATION'],
    'OTHER_CONDITIONAL_SOLUTIONS': chapter_1_content['OTHER_CONDITIONAL_SOLUTIONS'],
    'QA_CUSTOMERS': chapter_1_content['QA_CUSTOMERS'],
    'PRODUCT_SUMMARY': sections.get('1. Product Summary', ''),
    'CLIENT_NEEDS': sections.get('2. Understanding the Client\'s Needs', ''),
    'PRODUCT_DESCRIPTION': sections.get('3. Product Description', {}),
    'ARCHITECTURAL_DESCRIPTION': (sections.get('3. Product Description', {}), '3.1 Architectural Description'),
    'KEY_FEATURES': (sections.get('3. Product Description', {}), '3.2 Key Features & Functionalities'),
    'SCOPE': (sections.get('3. Product Description', {}), '3.3 Scope / Out-of-Scope'),
    'REQUIREMENTS': (sections.get('3. Product Description', {}), '3.4 Requirements & Prerequisites'),
    'VALUE_PROPOSITION': sections.get('4. Value Proposition', ''),
    'DIFFERENTIATORS': sections.get('5. Key Differentiators', ''),
    'TRANSITION_TRANSFORMATION': sections.get('6. Transition & Transformation', {}),
    'CLIENT_RESPONSIBILITIES': sections.get('7. Client Responsibilities', ''),
    'OPERATIONAL_SUPPORT': sections.get('8. Operational Support', ''),
    'TERMS_CONDITIONS': sections.get('9. Terms & Conditions', ''),
    'SLA_KPI': sections.get('10. SLA & KPI Management', ''),
    'PRICING_ELEMENTS': sections.get('11. Pricing Elements', ''),
}

filled_count = 0
for sdt_tag, md_source in mapping.items():
    # Handle dynamic subsection extraction
    content = ''
    if isinstance(md_source, tuple):
        parent_dict, sub_key = md_source
        if isinstance(parent_dict, dict) and sub_key in parent_dict:
            content = parent_dict[sub_key]
    elif isinstance(md_source, dict):
        content = md_source.get('_main', '')
    else:
        content = md_source
    
    if fill_sdt_field(doc, sdt_tag, content):
        print(f"  ✓ {sdt_tag}")
        filled_count += 1
    else:
        print(f"  ⚠ {sdt_tag} - NOT FOUND or EMPTY")

# Save the filled template
output_filename = 'SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0]_FILLED.docx'
doc.save(output_filename)
print(f"\n✓ Filled {filled_count} SDT fields")
print(f"✓ Saved filled SDT to: {output_filename}")
