from zipfile import ZipFile
from pathlib import Path
from lxml import etree
from collections import Counter

p = Path(r"output/docx/SD - IBM Power on Premise [DV0.9]_mapped_FINAL.docx")
print('exists', p.exists())
with ZipFile(p, 'r') as z:
    names = set(z.namelist())
    print('entries', len(names))
    # 1) XML parse health
    bad = []
    xml_parts = [n for n in names if n.endswith('.xml') or n.endswith('.rels')]
    for n in xml_parts:
        try:
            etree.fromstring(z.read(n))
        except Exception as e:
            bad.append((n, str(e)))
    print('bad_xml', len(bad))
    for n,e in bad[:10]:
        print('  ', n, e)

    # 2) document relationships targets exist
    rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
    nsr = {'r':'http://schemas.openxmlformats.org/package/2006/relationships'}
    misses = []
    for rel in rels.xpath('/r:Relationships/r:Relationship', namespaces=nsr):
        rid = rel.get('Id')
        tgt = rel.get('Target','')
        typ = rel.get('Type','')
        if '/image' in typ:
            key = 'word/' + tgt.replace('\\','/') if not tgt.startswith('/') else tgt[1:]
            if key not in names:
                misses.append((rid, tgt, key))
    print('missing_image_targets', len(misses))
    for m in misses[:10]:
        print('  ', m)

    # 3) duplicate drawing ids in document
    doc = etree.fromstring(z.read('word/document.xml'))
    ns = {
        'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }
    docpr = [int(x) for x in doc.xpath('.//wp:docPr/@id', namespaces=ns) if str(x).isdigit()]
    cnvpr = [int(x) for x in doc.xpath('.//pic:cNvPr/@id', namespaces=ns) if str(x).isdigit()]
    d1 = [k for k,v in Counter(docpr).items() if v>1]
    d2 = [k for k,v in Counter(cnvpr).items() if v>1]
    print('docPr_count', len(docpr), 'dup', len(d1))
    print('cNvPr_count', len(cnvpr), 'dup', len(d2))

    # 4) content types coverage for images in word/media
    ct = etree.fromstring(z.read('[Content_Types].xml'))
    nsc = {'ct':'http://schemas.openxmlformats.org/package/2006/content-types'}
    defaults = {d.get('Extension','').lower(): d.get('ContentType','') for d in ct.xpath('/ct:Types/ct:Default', namespaces=nsc)}
    media_ext = sorted({Path(n).suffix.lower().lstrip('.') for n in names if n.startswith('word/media/') and '.' in Path(n).name})
    print('media_ext', media_ext)
    missing_ct = [e for e in media_ext if e and e not in defaults]
    print('missing_content_types', missing_ct)
