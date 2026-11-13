"""Make an empty IPBES Word template from an existing IPBES report.

Keeps the styles, numbering, page set-up and header/footer layout of the report's
first section (the masthead page set-up: different first page, even/odd headers
showing the symbol, page numbers). Removes all body content and replaces the
symbol and job number in headers/footers with placeholders.

Requires python-docx (pip install --user python-docx).

Usage:
    python3 tools/make_template.py temp/2026-12.docx skills/ipbes-editorial/assets/ipbes-document-template.docx
"""
import copy
import re
import sys

from docx import Document
from docx.oxml.ns import qn

SYMBOL_PLACEHOLDER = 'IPBES/XX/X'


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
