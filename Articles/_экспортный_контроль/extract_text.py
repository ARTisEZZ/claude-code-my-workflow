"""Полный текст статьи из .docx для экспертизы экспортного контроля — ничего не пропуская.

    python extract_text.py СТАТЬЯ.docx [ВЫХОД.txt]
    python extract_text.py --selfcheck

Берётся всё, что python-docx теряет: текст в порядке документа с таблицами (строка = « | »), сноски, концевые
сноски, колонтитулы, надписи, формулы Word (OMML), подписи и замещающий текст рисунков. Формулы MathType — это
OLE-объекты-картинки, их текста в .docx нет: скрипт их считает, и эксперт обязан сверить формулы по article.md.
В конце — сводка всех чисел с единицами для сверки со списками.
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
O = '{urn:schemas-microsoft-com:office:office}'

UNITS = (r'ГГц|МГц|кГц|Гц|GHz|MHz|дБм|дБ|dBm|dB|мВт|Вт|mW|W|мкм|нм|мм|см|µm|um|nm|мА/мм|мА|А|мСм/мм|мСм|Ом·мм|Ом/□|Ом|'
         r'В|V|К|K|°C|пФ/мм|пФ|фФ|пГн|нГн|см⁻²|см-2|см⁻³|см-3|см2/\(В·с\)|%|с|мкс|нс|пс')
NUM = re.compile(r'(?<![\w.])[-−+]?\d+(?:[.,]\d+)?(?:\s?[·x×]\s?10[⁻\-−]?\^?\d+)?\s?(?:' + UNITS + r')(?![\w])')


def text_of(el):
    """Текст элемента: обычный (w:t), формулы Word (m:t), табуляции и переносы.
    Верхний индекс — «^», нижний — «_»: иначе 10¹⁷ см⁻³ читается как «1017 см-3»."""
    out, va = [], None
    for e in el.iter():
        if e.tag == W + 'r':
            v = e.find(f'{W}rPr/{W}vertAlign')
            va = v.get(W + 'val') if v is not None else None
        elif e.tag == M + 'r':
            va = None
        if e.tag in (W + 't', M + 't') and e.text:
            out.append({'superscript': '^', 'subscript': '_'}.get(va, '') + e.text)
        elif e.tag in (W + 'tab',):
            out.append('\t')
        elif e.tag in (W + 'br', W + 'cr'):
            out.append('\n')
    return ''.join(out)


def body_lines(root):
    """Абзацы и таблицы в порядке документа; таблица — строками «ячейка | ячейка»."""
    body = root.find(W + 'body')
    lines = []
    for el in body:
        if el.tag == W + 'p':
            lines.append(text_of(el))
        elif el.tag == W + 'tbl':
            lines.append('[ТАБЛИЦА]')
            for tr in el.iter(W + 'tr'):
                lines.append(' | '.join(text_of(tc).strip() for tc in tr.findall(W + 'tc')))
            lines.append('[/ТАБЛИЦА]')
    return lines


def extract(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    doc = ET.fromstring(z.read('word/document.xml'))
    parts = [('ОСНОВНОЙ ТЕКСТ', body_lines(doc))]
    for n in sorted(names):
        if re.fullmatch(r'word/(footnotes|endnotes|header\d*|footer\d*|comments)\.xml', n):
            root = ET.fromstring(z.read(n))
            lines = [text_of(p) for p in root.iter(W + 'p')]
            lines = [x for x in lines if x.strip()]
            if lines:
                parts.append((n.split('/')[-1], lines))
    alts = [f"{d.get('name', '')}: {d.get('descr', '')} {d.get('title', '')}".strip()
            for d in doc.iter(WP + 'docPr') if d.get('descr') or d.get('title')]
    ole = [o.get('ProgID', '?') for o in doc.iter(O + 'OLEObject')]
    omml = sum(1 for _ in doc.iter(M + 'oMath'))
    return parts, alts, ole, omml


def report(path):
    parts, alts, ole, omml = extract(path)
    out = []
    for title, lines in parts:
        out.append(f'===== {title} =====')
        out.extend(x for x in lines if x.strip())
    out.append('===== ЗАМЕЩАЮЩИЙ ТЕКСТ РИСУНКОВ =====')
    out.extend(alts or ['(нет)'])
    from collections import Counter
    out.append(f'===== ОБЪЕКТЫ =====\nформул Word (OMML): {omml}; OLE-объектов: {len(ole)} {dict(Counter(ole))}')
    if ole:
        out.append('ВНИМАНИЕ: OLE-формулы (MathType и т.п.) не содержат текста — сверить формулы по article.md / PDF.')
    full = '\n'.join(out)
    nums = sorted(set(m.group(0).strip() for m in NUM.finditer(full)))
    out.append(f'===== ЧИСЛА С ЕДИНИЦАМИ ({len(nums)}) =====')
    out.append('; '.join(nums))
    words = len(re.findall(r'\w+', full))
    out.append(f'===== СЛОВ ≈ {words} =====')
    return '\n'.join(out)


def selfcheck():
    import io
    xml = (f'<w:document xmlns:w="{W[1:-1]}" xmlns:m="{M[1:-1]}"><w:body>'
           '<w:p><w:r><w:t>Частота 33 ГГц, Pнас 0,5 Вт, n = 10</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>17</w:t></w:r><w:r><w:t> см</w:t></w:r></w:p>'
           '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>NF</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1,2 дБ</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
           '<w:p><m:oMath><m:r><m:t>F=1+T/290</m:t></m:r></m:oMath></w:p></w:body></w:document>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        z.writestr('word/document.xml', xml)
        z.writestr('word/footnotes.xml', f'<w:footnotes xmlns:w="{W[1:-1]}"><w:footnote><w:p><w:r><w:t>Сноска 40 ГГц</w:t></w:r></w:p></w:footnote></w:footnotes>')
    buf.seek(0)
    r = report(buf)
    for s in ('33 ГГц', '0,5 Вт', 'NF | 1,2 дБ', 'F=1+T/290', 'Сноска 40 ГГц', 'формул Word (OMML): 1', '1,2 дБ', '10^17 см'):
        assert s in r, s
    print('selfcheck ок')


if __name__ == '__main__':
    if sys.argv[1] == '--selfcheck':
        selfcheck()
    else:
        sys.stdout.reconfigure(encoding='utf-8')
        text = report(sys.argv[1])
        if len(sys.argv) > 2:
            open(sys.argv[2], 'w', encoding='utf-8').write(text)
            print('записано:', sys.argv[2], len(text), 'симв.')
        else:
            print(text)
