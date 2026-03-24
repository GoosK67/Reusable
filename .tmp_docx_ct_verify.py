from zipfile import ZipFile
from pathlib import Path
from lxml import etree
p = Path(r"output/docx/SD -  Managed Security Infrastructure [PRD.4.2.001][PV1.0][DV1.0]_mapped_FINAL.docx")
with ZipFile(p,'r') as z:
    ct = etree.fromstring(z.read('[Content_Types].xml'))
    ns = {'ct':'http://schemas.openxmlformats.org/package/2006/content-types'}
    defs = {d.get('Extension','').lower(): d.get('ContentType','') for d in ct.xpath('/ct:Types/ct:Default', namespaces=ns)}
    print('png', defs.get('png'))
    print('jpg', defs.get('jpg'))
    print('jpeg', defs.get('jpeg'))
