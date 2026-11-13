"""Make an empty IPBES Word template from an existing IPBES report.

Keeps the styles, numbering, page set-up and header/footer layout of the report's
first section (the masthead page set-up: different first page, even/odd headers
showing the symbol, page numbers). Removes all body content and replaces the
symbol and job number in headers/footers with placeholders.

Requires python-docx (pip install --user python-docx).

Usage:
    python3 tools/make_template.py temp/2026-12.docx ipbes-editorial/assets/ipbes-document-template.docx
"""
import copy
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SYMBOL_PLACEHOLDER = 'IPBES/XX/X'

# Numbered paragraphs in IPBES reports: number at 2.2 cm (aligned with heading text),
# wrapped lines aligned with the number, text after the number at 3.3 cm.
# In the reports this is set on every paragraph; here it is put into the numbering
# definition so that any document made from the template looks right.
PARA_LEFT = 1247      # 2.2 cm in twips
PARA_TEXT = 1871      # 3.3 cm: tab after the number


def fix_body_numbering(doc):
    """Set level indents of the Normal_number list to the IPBES report layout."""
    st = doc.styles['Normal_number'].element.find('.//' + qn('w:numId'))
    if st is None:
        return
    num_id = st.get(qn('w:val'))
    numbering = doc.part.numbering_part.element
    num = next(n for n in numbering.findall(qn('w:num')) if n.get(qn('w:numId')) == num_id)
    abs_id = num.find(qn('w:abstractNumId')).get(qn('w:val'))
    absn = next(a for a in numbering.findall(qn('w:abstractNum')) if a.get(qn('w:abstractNumId')) == abs_id)
    levels = list(absn.findall(qn('w:lvl'))) + list(num.iter(qn('w:lvl')))
    for lvl in levels:
        # only level 0 ("1.") is wrong in the source; (a), (i) levels already match the reports
        if lvl.get(qn('w:ilvl')) != '0':
            continue
        ppr = lvl.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr'); lvl.append(ppr)
        for old in ppr.findall(qn('w:tabs')) + ppr.findall(qn('w:ind')):
            ppr.remove(old)
        tabs = OxmlElement('w:tabs'); tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'num'); tab.set(qn('w:pos'), str(PARA_TEXT)); tabs.append(tab)
        ind = OxmlElement('w:ind'); ind.set(qn('w:left'), str(PARA_LEFT)); ind.set(qn('w:firstLine'), '0')
        ppr.append(tabs); ppr.append(ind)


def main(src, dst):
    doc = Document(src)
    body = doc.element.body

    # Section properties of the first section (masthead page): headers, footers, titlePg
    first_sect = next(body.iter(qn('w:sectPr')))
    sect = copy.deepcopy(first_sect)
    # remove 'continuous' type etc. if any
    for t in sect.findall(qn('w:type')):
        sect.remove(t)

    for child in list(body):
        body.remove(child)
    p = doc.add_paragraph('')
    p.style = doc.styles['Normal_No_number']
    body.append(sect)

    # Scrub every header/footer part reachable from the document
    for rel in doc.part.rels.values():
        if 'header' in rel.reltype or 'footer' in rel.reltype:
            el = rel.target_part.element
            # Job number paragraphs: blank entirely (text is split over many runs)
            for p in el.iter(qn('w:p')):
                ps = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
                if ps is not None and 'jobnumber' in ps.get(qn('w:val'), '').lower():
                    for t in p.iter(qn('w:t')):
                        t.text = ''
            for t in el.iter(qn('w:t')):
                if t.text:
                    t.text = re.sub(r'IPBES/\d+/[\w./]+', SYMBOL_PLACEHOLDER, t.text)
                    t.text = re.sub(r'K\d{7}\s*\[E\]', '', t.text)
                    t.text = re.sub(r'^\s*\d{6}\s*$', '', t.text)

    fix_body_numbering(doc)

    cp = doc.core_properties
    cp.title = 'IPBES document'
    cp.subject = ''
    cp.author = 'IPBES secretariat'
    cp.keywords = ''
    cp.comments = ''
    cp.last_modified_by = 'IPBES secretariat'
    doc.save(dst)
    print('template written to', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
