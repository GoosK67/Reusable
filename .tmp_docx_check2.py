from zipfile import ZipFile
from pathlib import Path
from lxml import etree
from collections import Counter

p = Path(r"output/docx/SD -  Managed Security Infrastructure [PRD.4.2.001][PV1.0][DV1.0]_mapped_FINAL.docx")
print('exists', p.exists())
with ZipFile(p, 'r') as z:
    names = set(z.namelist())
    rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
    nsr = {'r':'http://schemas.openxmlformats.org/package/2006/relationships'}
    doc = etree.fromstring(z.read('word/document.xml'))
    ns = {
        'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }

    # relationship consistency
    rel_map = {rel.get('Id'): rel.get('Target') for rel in rels.xpath('/r:Relationships/r:Relationship', namespaces=nsr)}
    blips = doc.xpath('.//a:blip/@r:embed', namespaces=ns)
    print('embedded_blips', blips)
    for rid in blips:
        tgt = rel_map.get(rid)
        print('rid', rid, 'target', tgt, 'exists', ('word/' + tgt) in names if tgt else False)

    # duplicate IDs
    docpr = [int(x) for x in doc.xpath('.//wp:docPr/@id', namespaces=ns) if str(x).isdigit()]
    cnvpr = [int(x) for x in doc.xpath('.//pic:cNvPr/@id', namespaces=ns) if str(x).isdigit()]
    d1 = [k for k,v in Counter(docpr).items() if v>1]
    d2 = [k for k,v in Counter(cnvpr).items() if v>1]
    print('docPr ids', sorted(docpr))
    print('cNvPr ids', sorted(cnvpr))
    print('dup docPr', d1)
    print('dup cNvPr', d2)

    # check all xml parsable
    bad=[]
    for n in names:
        if n.endswith('.xml') or n.endswith('.rels'):
            try: etree.fromstring(z.read(n))
            except Exception as e: bad.append((n,str(e)))
    print('bad_xml', bad)
