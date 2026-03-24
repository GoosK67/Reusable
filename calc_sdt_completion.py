from docx import Document
import re

ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

candidates = [
    'SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0].docx',
    'SDT - Cegeka DBMS for Oracle on Azure [PRD.0.8.001][PV1.0][DV1.0]_FILLED.docx',
    'SDT_DBMS_Oracle_Azure_FILLED.docx',
]


def calc_for_file(fn: str):
    doc = Document(fn)
    sdts = doc.element.body.findall('.//' + ns + 'sdt')

    total = len(sdts)
    filled = 0
    empty = 0
    todo = 0

    for s in sdts:
        c = s.find(ns + 'sdtContent')
        txt = ''
        if c is not None:
            ts = c.findall('.//' + ns + 't')
            txt = ' '.join((t.text or '').strip() for t in ts)
            txt = re.sub(r'\s+', ' ', txt).strip()

        if txt:
            filled += 1
        else:
            empty += 1

        up = txt.upper()
        if '[TO BE COMPLETED]' in up or 'TO BE COMPLETED' in up:
            todo += 1

    pct = round((filled / total * 100), 2) if total else 0.0
    return total, filled, empty, todo, pct


for fn in candidates:
    try:
        total, filled, empty, todo, pct = calc_for_file(fn)
        print(f'FILE={fn}')
        print(f'TOTAL_SDT={total}')
        print(f'FILLED_SDT={filled}')
        print(f'EMPTY_SDT={empty}')
        print(f'TODO_MARKERS={todo}')
        print(f'PERCENT_FILLED={pct}')
        print('---')
    except Exception as ex:
        print(f'FILE={fn}')
        print(f'ERROR={ex}')
        print('---')
