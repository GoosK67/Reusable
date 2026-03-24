from zipfile import ZipFile
from pathlib import Path
from lxml import etree
p = Path(r"output/docx/SD -  Managed Security Infrastructure [PRD.4.2.001][PV1.0][DV1.0]_mapped_FINAL.docx")
with ZipFile(p,'r') as z:
    names = set(z.namelist())
    ct = etree.fromstring(z.read('[Content_Types].xml'))
    ns = {'ct':'http://schemas.openxmlformats.org/package/2006/content-types'}
    defs = {d.get('Extension','').lower(): d.get('ContentType','') for d in ct.xpath('/ct:Types/ct:Default', namespaces=ns)}
    media_ext = sorted({Path(n).suffix.lower().lstrip('.') for n in names if n.startswith('word/media/') and '.' in Path(n).name})
    print('media_ext', media_ext)
    print('defaults_subset', {k:defs.get(k) for k in media_ext})
