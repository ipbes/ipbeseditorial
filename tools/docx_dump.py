"""Dump a .docx as plain text, one paragraph per line, prefixed with its Word style.

Standard library only. Usage:
    python3 tools/docx_dump.py input.docx [output.txt]

Each line: [StyleName] paragraph text
Tables are marked with [TABLE] / [/TABLE]; headers/footers are dumped first,
marked [HEADER:file] / [FOOTER:file]. Images are noted as {IMAGE:name}.
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def style_names(z):
    names = {}
    try:
        root = ET.fromstring(z.read('word/styles.xml'))
    except KeyError:
        return names
    for s in root.iter(W + 'style'):
        sid = s.get(W + 'styleId')
        n = s.find(W + 'name')
        names[sid] = n.get(W + 'val') if n is not None else sid
    return names


def rels(z, part):
    path = re.sub(r'([^/]+)$', r'_rels/\1.rels', part)
    out = {}
    try:
        root = ET.fromstring(z.read(path))
    except KeyError:
        return out
    for r in root:
        out[r.get('Id')] = r.get('Target')
    return out


def para_text(p, rel):
    parts = []
    for el in p.iter():
        if el.tag == W + 't':
            parts.append(el.text or '')
        elif el.tag == W + 'tab':
            parts.append('\t')
        elif el.tag in (W + 'br', W + 'cr'):
            parts.append(' / ')
        elif el.tag == W + 'noBreakHyphen':
            parts.append('‑')
        elif el.tag == W + 'footnoteReference':
            parts.append('[^%s]' % el.get(W + 'id'))
        elif el.tag == A + 'blip':
            rid = el.get(R + 'embed')
            parts.append('{IMAGE:%s}' % rel.get(rid, rid))
    return ''.join(parts)


def dump_part(z, part, names, out, prefix=''):
    try:
        root = ET.fromstring(z.read(part))
    except KeyError:
        return
    rel = rels(z, part)
    body = root.find(W + 'body')
    if body is None:
        body = root

    def walk(node, depth=0):
        for child in node:
            if child.tag == W + 'p':
                ps = child.find(W + 'pPr/' + W + 'pStyle')
                sid = ps.get(W + 'val') if ps is not None else 'Normal'
                txt = para_text(child, rel)
                num = child.find(W + 'pPr/' + W + 'numPr')
                mark = '#' if num is not None else ''
                out.append('%s[%s%s] %s' % (prefix + '  ' * depth, names.get(sid, sid), mark, txt))
            elif child.tag == W + 'tbl':
                out.append(prefix + '  ' * depth + '[TABLE]')
                for row in child.iter(W + 'tr'):
                    cells = []
                    for c in row.findall(W + 'tc'):
                        cells.append(' '.join(para_text(p, rel) for p in c.iter(W + 'p')).strip())
                    out.append(prefix + '  ' * depth + ' | '.join(cells))
                out.append(prefix + '  ' * depth + '[/TABLE]')
            elif child.tag in (W + 'sdt', W + 'sdtContent', W + 'customXml'):
                walk(child, depth)
    walk(body)


def main():
    src = sys.argv[1]
    z = zipfile.ZipFile(src)
    names = style_names(z)
    out = []
    for n in sorted(z.namelist()):
        m = re.match(r'word/(header|footer)\d*\.xml$', n)
        if m:
            out.append('[%s:%s]' % (m.group(1).upper(), n))
            dump_part(z, n, names, out, prefix='  ')
    out.append('[BODY]')
    dump_part(z, 'word/document.xml', names, out)
    if 'word/footnotes.xml' in z.namelist():
        out.append('[FOOTNOTES]')
        dump_part(z, 'word/footnotes.xml', names, out)
    text = '\n'.join(out) + '\n'
    if len(sys.argv) > 2:
        open(sys.argv[2], 'w', encoding='utf-8').write(text)
    else:
        sys.stdout.write(text)


if __name__ == '__main__':
    main()
