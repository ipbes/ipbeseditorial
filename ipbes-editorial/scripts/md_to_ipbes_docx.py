"""Convert a Markdown draft into a Word document in the IPBES style.

Uses the IPBES template (styles, numbering, page set-up, headers) in ../assets/, made by
tools/make_template.py in the repository. Requires python-docx:  pip install --user python-docx

Usage:
    python3 scripts/md_to_ipbes_docx.py draft.md output.docx [options]

Options:
    --numbered          number body paragraphs (style Normal_number), as in reports
    --symbol SYMBOL     document symbol for the page headers, e.g. IPBES/13/5
    --number-headings   add I., II. / A., B. / 1., 2. to headings that have no number
    --template FILE     template to use (default: ../assets/ipbes-document-template.docx, next to this script)

Markdown supported:
    # Title                      -> BB_Title
    ## I. Chapter                -> CH1   (### -> CH2, #### -> CH3, ##### -> CH4)
    paragraphs                   -> Normal_number (--numbered) or Normal_No_number
    1. / (a) / a. / - lists      -> numbered subparagraphs (a), (i) ... or literal labels
    > quotation                  -> indented block, no quotation marks
    | tables |                   -> table with horizontal rules only; a preceding line
                                    "Table 1: Title" or "**Table 1** Title" becomes Title_table,
                                    and a line "(United States dollars)" becomes the headnote
    **bold**, *italic*, [text](url), footnotes [^1] with "[^1]: text" definitions

Typography: straight quotes and apostrophes become curly; " - " and "--" become an en dash.
The script reports what it changed. It does NOT apply the editorial rules (see
ipbes-editorial/SKILL.md) - do that before or after converting.
"""
import argparse
import copy
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, '..', 'assets', 'ipbes-document-template.docx')

ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV',
         'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV']
LOWER_ROMAN = [r.lower() for r in ROMAN]
LETTERS = 'abcdefghijklmnopqrstuvwxyz'

stats = {}


def count(what, n=1):
    if n:
        stats[what] = stats.get(what, 0) + n


# ---------------------------------------------------------------- typography
def typography(text):
    """Curly quotes/apostrophes and en dashes, outside URLs and code."""
    out = []
    for i, chunk in enumerate(re.split(r'(`[^`]*`|\]\([^)]*\)|https?://\S+)', text)):
        if i % 2 == 1:
            out.append(chunk)
            continue
        n0 = chunk
        chunk, n = re.subn(r'(?<=\w)\'(?=\w)', '\u2019', chunk); count('apostrophes made curly', n)
        chunk, n = re.subn(r'(^|[\s(\[\u2014\u2013-])\'', '\\1\u2018', chunk); count('single quotes made curly', n)
        chunk, n = re.subn(r'\'', '\u2019', chunk); count('single quotes made curly', n)
        chunk, n = re.subn(r'(^|[\s(\[\u2014\u2013-])"', '\\1\u201c', chunk); count('double quotes made curly', n)
        chunk, n = re.subn(r'"', '\u201d', chunk); count('double quotes made curly', n)
        chunk, n = re.subn(r' -{1,2} ', ' \u2013 ', chunk); count('spaced hyphens made en dashes', n)
        chunk, n = re.subn(r'\u2014', '\u2013', chunk); count('em dashes made en dashes', n)
        chunk, n = re.subn(r'(?<=\S)  +(?=\S)', ' ', chunk); count('double spaces removed', n)
        out.append(chunk)
    return ''.join(out)


# ---------------------------------------------------------------- inline runs
INLINE = re.compile(r'(\*\*[^*]+\*\*|__[^_]+__|\*[^*\s][^*]*\*|_[^_\s][^_]*_|\[\^[^\]]+\]|\[[^\]]+\]\([^)]+\)|`[^`]+`)')


class Footnotes:
    def __init__(self, doc, defs):
        self.doc = doc
        self.defs = defs
        self.part = None
        for rel in doc.part.rels.values():
            if rel.reltype.endswith('/footnotes'):
                self.part = rel.target_part
        self.next_id = 1
        self.root = None
        if self.part is not None:
            from docx.oxml import parse_xml
            self.root = root = parse_xml(self.part.blob)
            # drop existing (non-separator) footnotes from the template
            for fn in list(root.findall(qn('w:footnote'))):
                if fn.get(qn('w:type')) is None:
                    root.remove(fn)
            ids = [int(fn.get(qn('w:id'))) for fn in root.findall(qn('w:footnote'))]
            self.next_id = max(ids + [0]) + 1

    def add(self, paragraph, label):
        text = self.defs.get(label)
        if text is None or self.part is None:
            paragraph.add_run('[^%s]' % label)
            return
        fid = self.next_id
        self.next_id += 1
        # reference in body
        run = paragraph.add_run()
        rpr = run._r.get_or_add_rPr()
        rs = OxmlElement('w:rStyle'); rs.set(qn('w:val'), self._style_id('footnote reference', 'FootnoteReference'))
        rpr.append(rs)
        ref = OxmlElement('w:footnoteReference'); ref.set(qn('w:id'), str(fid))
        run._r.append(ref)
        # footnote content
        fn = OxmlElement('w:footnote'); fn.set(qn('w:id'), str(fid))
        p = OxmlElement('w:p')
        ppr = OxmlElement('w:pPr'); ps = OxmlElement('w:pStyle'); ps.set(qn('w:val'), self._style_id('footnote text', 'FootnoteText'))
        ppr.append(ps); p.append(ppr)
        r = OxmlElement('w:r'); rpr = OxmlElement('w:rPr'); rs = OxmlElement('w:rStyle')
        rs.set(qn('w:val'), self._style_id('footnote reference', 'FootnoteReference')); rpr.append(rs); r.append(rpr)
        r.append(OxmlElement('w:footnoteRef')); p.append(r)
        r2 = OxmlElement('w:r'); t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = ' ' + typography(text)
        r2.append(t); p.append(r2)
        fn.append(p)
        self.root.append(fn)
        count('footnotes created')

    def save(self):
        if self.root is not None:
            from lxml import etree
            self.part._blob = etree.tostring(self.root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _style_id(self, name, fallback):
        for s in self.doc.styles:
            if s.name.lower() == name:
                return s.style_id
        return fallback


def add_runs(paragraph, text, footnotes, italic=False):
    for piece in INLINE.split(text):
        if not piece:
            continue
        if piece.startswith('[^'):
            footnotes.add(paragraph, piece[2:-1])
        elif piece.startswith('**') or piece.startswith('__'):
            r = paragraph.add_run(piece[2:-2]); r.bold = True; r.italic = italic or None
        elif (piece.startswith('*') or piece.startswith('_')) and len(piece) > 2:
            r = paragraph.add_run(piece[1:-1]); r.italic = True
        elif piece.startswith('['):
            m = re.match(r'\[([^\]]+)\]\(([^)]+)\)', piece)
            add_hyperlink(paragraph, m.group(1), m.group(2))
        elif piece.startswith('`'):
            r = paragraph.add_run(piece[1:-1]); r.italic = italic or None
        else:
            r = paragraph.add_run(piece); r.italic = italic or None


def add_hyperlink(paragraph, text, url):
    rid = paragraph.part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), rid)
    r = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)
    c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rpr.append(c)
    r.append(rpr); t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; r.append(t)
    h.append(r); paragraph._p.append(h)


# ---------------------------------------------------------------- numbering
def set_num(paragraph, num_id, level):
    ppr = paragraph._p.get_or_add_pPr()
    numpr = OxmlElement('w:numPr')
    il = OxmlElement('w:ilvl'); il.set(qn('w:val'), str(level))
    ni = OxmlElement('w:numId'); ni.set(qn('w:val'), str(num_id))
    numpr.append(il); numpr.append(ni)
    ppr.insert(0, numpr) if ppr.find(qn('w:pStyle')) is None else ppr.find(qn('w:pStyle')).addnext(numpr)


def body_num_id(doc):
    """numId used by the Normal_number style in the template."""
    st = doc.styles['Normal_number']
    n = st.element.find('.//' + qn('w:numId'))
    return int(n.get(qn('w:val'))) if n is not None else None


def set_indent(paragraph, left_cm, hanging_cm=0):
    from docx.shared import Cm
    pf = paragraph.paragraph_format
    pf.left_indent = Cm(left_cm)
    if hanging_cm:
        pf.first_line_indent = Cm(-hanging_cm)


def label_for(kind, i):
    if kind == 'num':
        return '%d.' % (i + 1)
    if kind == 'alpha':
        return '(%s)' % LETTERS[i % 26]
    return '(%s)' % LOWER_ROMAN[i % len(LOWER_ROMAN)]


# ---------------------------------------------------------------- tables
def table_rules(table):
    """Horizontal rules only: top, below header row, bottom."""
    tbl = table._tbl
    tblpr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for side in ('top', 'bottom'):
        e = OxmlElement('w:' + side); e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), '000000')
        borders.append(e)
    for side in ('left', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side); e.set(qn('w:val'), 'nil'); borders.append(e)
    tblpr.append(borders)
    # rule under the header row
    for cell in table.rows[0].cells:
        tcpr = cell._tc.get_or_add_tcPr()
        b = OxmlElement('w:tcBorders'); e = OxmlElement('w:bottom')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), '000000')
        b.append(e); tcpr.append(b)
    # repeat header row
    trpr = table.rows[0]._tr.get_or_add_trPr()
    h = OxmlElement('w:tblHeader'); trpr.append(h)


NUMERIC = re.compile(r'^[\s(]*[-–$€£]?\s*[\d.,  ]+%?\)?\s*$|^[–.]{1,2}$')


def add_table(doc, rows, footnotes):
    ncols = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=ncols)
    for i, r in enumerate(rows):
        for j in range(ncols):
            txt = r[j].strip() if j < len(r) else ''
            cell = t.cell(i, j)
            p = cell.paragraphs[0]
            p.style = doc.styles['Normal-pool']
            if NUMERIC.match(txt) and i > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                txt = re.sub(r'(?<=\d),(?=\d{3}\b)', '\u00a0', txt)  # thousands: space in tables
            add_runs(p, typography(txt), footnotes, italic=(i == 0))
    table_rules(t)
    table_layout(t, ncols)
    count('tables')
    return t


TEXT_INDENT = 1247     # 2.2 cm: tables start where the paragraph text starts
TEXT_WIDTH = 8250      # A4 width minus IPBES margins (2.5 + 1.75 cm) minus the 2.2 cm indent, in twips


def table_layout(table, ncols):
    """Indent the table to the text column and share the width: first column wider."""
    from docx.shared import Twips
    tblpr = table._tbl.tblPr
    for tag in ('w:tblInd', 'w:tblW'):
        for old in tblpr.findall(qn(tag)):
            tblpr.remove(old)
    w = OxmlElement('w:tblW'); w.set(qn('w:w'), str(TEXT_WIDTH)); w.set(qn('w:type'), 'dxa'); tblpr.append(w)
    ind = OxmlElement('w:tblInd'); ind.set(qn('w:w'), str(TEXT_INDENT)); ind.set(qn('w:type'), 'dxa'); tblpr.append(ind)
    table.autofit = False
    first = TEXT_WIDTH // 2 if ncols > 1 else TEXT_WIDTH
    rest = (TEXT_WIDTH - first) // (ncols - 1) if ncols > 1 else 0
    widths = [first] + [rest] * (ncols - 1)
    for row in table.rows:
        for j, cell in enumerate(row.cells):
            cell.width = Twips(widths[j])
    # Word draws a fixed-layout table from the grid, not from the cell widths: keep them equal.
    for col, w in zip(table._tbl.tblGrid.findall(qn('w:gridCol')), widths):
        col.set(qn('w:w'), str(w))
    order_tblpr(tblpr)


# Element order required by the Word schema (CT_TblPr); Word may ignore elements out of order.
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
               'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
               'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']


def order_tblpr(tblpr):
    rank = {qn('w:' + n): i for i, n in enumerate(TBLPR_ORDER)}
    children = sorted(tblpr, key=lambda e: rank.get(e.tag, len(rank)))
    for e in children:
        tblpr.remove(e)
        tblpr.append(e)


# ---------------------------------------------------------------- main conversion
HEAD_NUM = re.compile(r'^(?:([IVXLC]+)\.|([A-Z])\.|(\d+)\.|\(([a-z])\))\s+')


LABEL = r'(?:\(?(?:[a-z]|[ivx]+|\d+)\)|[a-z]\.|\d+\.)'


def normalise_list_line(line):
    """Accept the list forms an assistant may write: "**(a)** text" -> "(a) text";
    "- (a) text" or "* (a) text" -> "(a) text" (the bullet is dropped, the label kept)."""
    line = re.sub(r'^(\s*)\*\*(' + LABEL + r')\*\*\s+', r'\1\2 ', line)
    line = re.sub(r'^(\s*)[-*+]\s+(' + LABEL + r')\s+', r'\1\2 ', line)
    return line


def _end(text):
    return re.sub(r'[*_\s]+$', '', text or '')


def ends_with_colon(text):
    return _end(text).endswith(':')


def continues_list(text):
    """A list item ending with ';', '; and' or '; or' is followed by another item of the same list."""
    t = _end(text)
    return t.endswith(';') or t.endswith('; and') or t.endswith('; or')


def cap_first(text):
    """Capitalize the first letter of a list item (after Markdown emphasis or an opening
    quote). URLs and items that start with a figure or a capital are left alone."""
    m = re.match(r'^([\s*_"“\'‘]*)([a-z])', text)
    if not m or re.match(r'(https?:|www\.)', text[m.end(1):], re.I):
        return text
    count('capitals after a colon')
    return text[:m.start(2)] + m.group(2).upper() + text[m.end(2):]


def convert(md_text, out, template, numbered, symbol, number_headings):
    doc = Document(template)
    body = doc.element.body
    # empty the template body, keep final sectPr
    for child in list(body):
        if child.tag != qn('w:sectPr'):
            body.remove(child)

    # footnote definitions
    defs = {}
    def grab(m):
        defs[m.group(1)] = m.group(2).strip()
        return ''
    md_text = re.sub(r'^\[\^([^\]]+)\]:\s*(.+)$', grab, md_text, flags=re.M)
    footnotes = Footnotes(doc, defs)
    num_id = body_num_id(doc)

    lines = [normalise_list_line(l) for l in md_text.splitlines()]
    i = 0
    counters = {2: 0, 3: 0, 4: 0, 5: 0}
    para_buf = []
    pending_table_title = None

    last_text = ['']          # text of the last body paragraph (to detect a chapeau ending with ':')
    # True while the next paragraph or list item follows a colon (or continues a list
    # introduced by one): IPBES practice is a capital letter there.
    colon_chain = [False]

    def flush_para():
        nonlocal para_buf
        if not para_buf:
            return
        text = typography(' '.join(s.strip() for s in para_buf))
        if colon_chain[0]:
            text = cap_first(text)
        colon_chain[0] = ends_with_colon(text) or (colon_chain[0] and continues_list(text))
        last_text[0] = text
        para_buf = []
        p = doc.add_paragraph(style='Normal_number' if numbered else 'Normal_No_number')
        add_runs(p, text, footnotes)
        count('paragraphs')

    while i < len(lines):
        line = lines[i]
        s = line.strip()

        if not s:
            flush_para(); i += 1; continue

        # headings
        m = re.match(r'^(#{1,5})\s+(.*)$', s)
        if m:
            flush_para()
            level = len(m.group(1))
            text = typography(m.group(2).strip().rstrip('.') if not re.search(r'\b[a-z]\.$', m.group(2)) else m.group(2).strip())
            last_text[0] = ''
            colon_chain[0] = False
            if level == 1:
                p = doc.add_paragraph(style='BB_Title')
                add_runs(p, text, footnotes)
                count('title')
            else:
                style = {2: 'CH1', 3: 'CH2', 4: 'CH3', 5: 'CH4'}[level]
                for deeper in range(level + 1, 6):
                    counters[deeper] = 0
                hm = HEAD_NUM.match(text)
                if hm:
                    label = hm.group(0).strip(); text = text[hm.end():]
                    counters[level] += 1
                elif number_headings:
                    n = counters[level]; counters[level] += 1
                    label = {2: ROMAN[n] + '.', 3: LETTERS[n].upper() + '.', 4: '%d.' % (n + 1), 5: '(%s)' % LETTERS[n]}[level]
                    count('heading numbers added')
                else:
                    label = ''
                p = doc.add_paragraph(style=style)
                if label:
                    p.add_run('\t' + label + '\t')
                else:
                    p.add_run('\t\t')
                add_runs(p, text, footnotes)
                count('headings')
            i += 1; continue

        # table title lines
        mt = re.match(r'^\**(Table|Figure)\s+([0-9IVX]+(?:\.\d+)?)\**\s*[:.\u2013-]?\s*(.*)$', s)
        if mt and (i + 1 < len(lines)):
            flush_para()
            colon_chain[0] = False
            kind, num, title = mt.group(1), mt.group(2), mt.group(3).strip('* ')
            p = doc.add_paragraph(style='Title_table' if kind == 'Table' else ('Title_figure' if 'Title_figure' in [x.name for x in doc.styles] else 'Title_table'))
            r = p.add_run('%s %s' % (kind, num))
            r.add_break()
            add_runs(p, typography(title), footnotes)
            # headnote
            if i + 1 < len(lines) and re.match(r'^\(.*\)$', lines[i + 1].strip()):
                hp = doc.add_paragraph(style='Title_table')
                hr = hp.add_run(lines[i + 1].strip()); hr.bold = False
                i += 1
            i += 1; continue

        # tables
        if s.startswith('|'):
            flush_para()
            colon_chain[0] = False
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                row = lines[i].strip().strip('|').split('|')
                if not all(re.match(r'^\s*:?-{2,}:?\s*$', c) for c in row):
                    rows.append(row)
                i += 1
            add_table(doc, rows, footnotes)
            continue

        # block quote
        if s.startswith('>'):
            flush_para()
            colon_chain[0] = False
            q = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                q.append(lines[i].strip()[1:].strip()); i += 1
            p = doc.add_paragraph(style='Normal_No_number')
            set_indent(p, 3.3)
            add_runs(p, typography(' '.join(q).strip('“”"')), footnotes)
            count('block quotations')
            continue

        # lists
        lm = re.match(r'^(\s*)([-*+]|\d+[.)]|\(?[a-z]\)|[a-z]\.|\(?[ivx]+\))\s+(.*)$', line)
        if lm:
            flush_para()
            items = []
            while i < len(lines):
                lm = re.match(r'^(\s*)([-*+]|\d+[.)]|\(?[a-z]\)|[a-z]\.|\(?[ivx]+\))\s+(.*)$', lines[i])
                if not lm:
                    if lines[i].strip() and items and lines[i].startswith('  '):
                        items[-1][2] += ' ' + lines[i].strip(); i += 1; continue
                    break
                depth = len(lm.group(1).replace('\t', '    ')) // 2
                marker = lm.group(2)
                if marker in '-*+':
                    kind = 'bullet'
                elif marker[0].isdigit():
                    kind = 'num'
                elif re.match(r'^\(?[ivx]+\)$', marker) and depth > 0:
                    kind = 'roman'
                else:
                    kind = 'alpha'
                items.append([depth, kind, lm.group(3)])
                i += 1
            base = min(d for d, _, _ in items)
            # A top-level numbered list (1., 2.) with no chapeau (previous paragraph not
            # ending with ':') continues the paragraph numbering, as in IPBES reports.
            # A list after a chapeau, or lettered (a), (b), is a list of subparagraphs.
            after_chapeau = last_text[0].rstrip().endswith(':')
            top_kind = next(k for d, k, _ in items if d == base)
            continues_paras = (top_kind == 'num' and not after_chapeau)
            shift = 0 if continues_paras else 1
            idx = {}
            in_chain = colon_chain[0]
            whole_list = in_chain          # a list introduced by a colon: every item
            for depth, kind, text in items:
                # IPBES practice: items on their own lines after a colon start with a
                # capital; the run continues while items end with ';', '; and', '; or'.
                if in_chain or whole_list:
                    text = cap_first(text)
                in_chain = ends_with_colon(text) or (in_chain and continues_list(text))
                d = depth - base
                for k in list(idx):
                    if k > d:
                        del idx[k]
                n = idx.get(d, 0); idx[d] = n + 1
                text = typography(text)
                if kind == 'bullet' and not numbered:
                    p = doc.add_paragraph(style='Normal_No_number')
                    set_indent(p, 2.2 + 0.6 * (d + 1), 0.6)
                    p.add_run('\u2022\t')
                elif numbered and num_id:
                    # level 0 = paragraph number (1.), level 1 = (a), level 2 = (i)
                    p = doc.add_paragraph(style='Normal_number')
                    if d + shift > 0:
                        set_num(p, num_id, min(d + shift, 4))
                else:
                    p = doc.add_paragraph(style='Normal_No_number')
                    lvl = d + shift
                    if lvl == 0:
                        p.add_run(label_for('num', n) + '\t')
                    else:
                        set_indent(p, 2.2 + 1.1 * lvl, 1.1)
                        p.add_run(label_for(['alpha', 'roman', 'num'][min(lvl - 1, 2)], n) + '\t')
                add_runs(p, text, footnotes)
                count('list items')
            last_text[0] = text
            colon_chain[0] = in_chain
            continue

        para_buf.append(s)
        i += 1
    flush_para()

    # header symbol
    if symbol:
        for rel in doc.part.rels.values():
            if 'header' in rel.reltype or 'footer' in rel.reltype:
                for t in rel.target_part.element.iter(qn('w:t')):
                    if t.text and 'IPBES/XX/X' in t.text:
                        t.text = t.text.replace('IPBES/XX/X', symbol)
                        count('header symbols set')

    footnotes.save()
    doc.save(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('markdown')
    ap.add_argument('output')
    ap.add_argument('--numbered', action='store_true')
    ap.add_argument('--symbol')
    ap.add_argument('--number-headings', action='store_true')
    ap.add_argument('--template', default=DEFAULT_TEMPLATE)
    a = ap.parse_args()
    if not os.path.exists(a.template):
        sys.exit('Template not found: %s (run tools/make_template.py first)' % a.template)
    md = open(a.markdown, encoding='utf-8').read()
    convert(md, a.output, a.template, a.numbered, a.symbol, a.number_headings)
    print('Written:', a.output)
    for k, v in sorted(stats.items()):
        print('  %-32s %d' % (k, v))
    if not a.symbol:
        print('  note: page headers still show IPBES/XX/X - use --symbol to set the document symbol')


if __name__ == '__main__':
    main()
