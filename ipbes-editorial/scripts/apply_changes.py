"""Apply an editorial change list to a Word document as tracked changes.

The Python twin of the Word macro IPBES_ApplyChangeList. The original document
is kept as it is (styles, numbering, masthead, footnotes); each change becomes a
real Word revision (deletion + insertion) that the author can accept or reject,
with the rule as a comment.

Requires python-docx (pip install --user python-docx).

Usage:
    python3 scripts/apply_changes.py document.docx changes.md corrected.docx [--author NAME]

The change list can be:
    - a Markdown table  | No. | Original | Revised | Rule |
    - a CSV file with columns Original, Revised[, Rule]
    - a JSON list of {"original": ..., "revised": ..., "rule": ...}
    - a Word document containing a table with columns Original and Revised

Safety rules (same as the macro):
    - a change is applied only if its original text is found exactly once,
      inside one paragraph;
    - otherwise it is reported (not found / found several times) and not applied;
    - only the part that differs is revised, snapped to word boundaries.
"""
import argparse
import copy
import csv
import datetime
import json
import os
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

QUOTE_MAP = str.maketrans({'‘': "'", '’': "'", '“': '"', '”': '"',
                           ' ': ' ', '‑': '-', '\u001e': '-'})


def norm(t):
    """One-to-one character normalization (lengths unchanged)."""
    return t.translate(QUOTE_MAP)


# ---------------------------------------------------------------- change lists
def clean_cell(t):
    t = t.strip()
    if len(t) >= 2 and t[0] == '`' and t[-1] == '`':
        t = t[1:-1]
    if len(t) >= 2 and t[0] in '"“' and t[-1] in '"”':
        t = t[1:-1]
    return t.replace('\\|', '|').strip()


def rows_to_changes(header, rows):
    h = [x.strip().lower() for x in header]

    def col(*names):
        for i, x in enumerate(h):
            if any(n in x for n in names):
                return i
        return None
    co, cr, cu, cn = col('original'), col('revised', 'corrected'), col('rule', 'reason'), col('no')
    if co is None or cr is None:
        sys.exit('Change list needs columns "Original" and "Revised" (found: %s)' % ', '.join(header))
    out = []
    for i, r in enumerate(rows, 1):
        g = lambda k: clean_cell(r[k]) if k is not None and k < len(r) else ''
        out.append({'no': g(cn) or str(i), 'original': g(co), 'revised': g(cr), 'rule': g(cu)})
    return out


def load_changes(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.json':
        data = json.load(open(path, encoding='utf-8'))
        return [{'no': str(d.get('no', i)), 'original': d.get('original', ''), 'revised': d.get('revised', ''),
                 'rule': d.get('rule', '')} for i, d in enumerate(data, 1)]
    if ext == '.csv':
        rows = list(csv.reader(open(path, encoding='utf-8-sig')))
        return rows_to_changes(rows[0], rows[1:])
    if ext == '.docx':
        d = Document(path)
        for t in d.tables:
            header = [c.text for c in t.rows[0].cells]
            if any('original' in x.lower() for x in header):
                return rows_to_changes(header, [[c.text for c in r.cells] for r in t.rows[1:]])
        sys.exit('No table with an "Original" column found in %s' % path)
    # Markdown table
    lines = [l.strip() for l in open(path, encoding='utf-8') if l.strip().startswith('|')]
    split = lambda l: [c for c in re.split(r'(?<!\\)\|', l.strip().strip('|'))]
    rows = [split(l) for l in lines if not re.match(r'^\|?\s*:?-{2,}', l)]
    if not rows:
        sys.exit('No Markdown table found in %s' % path)
    return rows_to_changes(rows[0], rows[1:])


# ---------------------------------------------------------------- paragraph model
def simple_runs(p):
    """Direct-child runs of a paragraph, with their text. Runs containing anything
    other than text (tabs, breaks, fields, footnote marks) are kept as 'opaque'."""
    out = []
    for child in p:
        if child.tag != qn('w:r'):
            continue
        texts = [c for c in child if c.tag not in (qn('w:rPr'),)]
        if texts and all(c.tag == qn('w:t') for c in texts):
            out.append((child, ''.join(c.text or '' for c in texts), True))
        else:
            # represent opaque content by a private-use character so it never matches
            out.append((child, '', False))
    return out


def para_text(p):
    return ''.join(t for _, t, _ in simple_runs(p))


def all_paragraphs(doc):
    return list(doc.element.body.iter(qn('w:p')))


def split_run(run, offset):
    """Split a simple text run at offset; returns (left, right)."""
    texts = [c for c in run if c.tag == qn('w:t')]
    full = ''.join(t.text or '' for t in texts)
    right = copy.deepcopy(run)
    for c in list(run):
        if c.tag == qn('w:t'):
            run.remove(c)
    for c in list(right):
        if c.tag == qn('w:t'):
            right.remove(c)
    lt = OxmlElement('w:t'); lt.set(qn('xml:space'), 'preserve'); lt.text = full[:offset]; run.append(lt)
    rt = OxmlElement('w:t'); rt.set(qn('xml:space'), 'preserve'); rt.text = full[offset:]; right.append(rt)
    run.addnext(right)
    return run, right


class Reviser:
    def __init__(self, author):
        self.author = author
        self.date = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        self.next_id = 90000

    def _mark(self, tag):
        e = OxmlElement(tag)
        e.set(qn('w:id'), str(self.next_id)); self.next_id += 1
        e.set(qn('w:author'), self.author)
        e.set(qn('w:date'), self.date)
        return e

    def revise(self, p, start, end, new_text):
        """Replace characters [start, end) of paragraph p with new_text as tracked change.
        Returns the list of new runs (for comments) or None if not possible."""
        # boundaries must fall in simple runs; split them (layout recomputed for each cut)
        def layout():
            out, pos = [], 0
            for r, t, ok in simple_runs(p):
                out.append((r, pos, pos + len(t), ok))
                pos += len(t)
            return out

        def cut(at):
            for r, a, b, ok in layout():
                if a < at < b:
                    if not ok:
                        return False
                    split_run(r, at - a)
                    return True
            return True
        if not cut(start) or not cut(end):
            return None
        runs = simple_runs(p)
        pos = 0
        inside, before = [], None
        for r, t, ok in runs:
            a, b = pos, pos + len(t)
            if a >= start and b <= end and b > a:
                if not ok:
                    return None
                inside.append(r)
            elif b <= start:
                before = r
            pos = b
        # deletion
        anchor = None
        if inside:
            dl = self._mark('w:del')
            inside[0].addprevious(dl)
            for r in inside:
                for t in r.findall(qn('w:t')):
                    t.tag = qn('w:delText')
                dl.append(r)
            anchor = dl
            template = inside[0]
        else:
            template = before if before is not None else (runs[0][0] if runs else None)
        new_runs = []
        if new_text:
            ins = self._mark('w:ins')
            r = OxmlElement('w:r')
            if template is not None and template.find(qn('w:rPr')) is not None:
                r.append(copy.deepcopy(template.find(qn('w:rPr'))))
            t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = new_text
            r.append(t); ins.append(r)
            if anchor is not None:
                anchor.addnext(ins)
            elif before is not None:
                before.addnext(ins)
            else:
                ppr = p.find(qn('w:pPr'))
                if ppr is not None:
                    ppr.addnext(ins)
                else:
                    p.insert(0, ins)
            new_runs.append(r)
        return new_runs or ([anchor] if anchor is not None else [])


def diff_region(a, b):
    """Same algorithm as the macro: common prefix/suffix snapped to word boundaries."""
    lo, lr = len(a), len(b)
    p = 0
    while p < lo and p < lr and a[p] == b[p]:
        p += 1
    while p > 0 and a[p - 1] != ' ':
        p -= 1
    s = 0
    while s < lo - p and s < lr - p and a[lo - s - 1] == b[lr - s - 1]:
        s += 1
    while s > 0 and lo - s > 0 and a[lo - s] != ' ' and a[lo - s - 1] != ' ':
        s -= 1
    return p, s


# ---------------------------------------------------------------- main
def apply(doc_path, changes_path, out_path, author):
    doc = Document(doc_path)
    changes = load_changes(changes_path)
    rev = Reviser(author)
    report = []
    for ch in changes:
        orig, new, rule = ch['original'], ch['revised'], ch['rule']
        if not orig or orig == new:
            report.append((ch['no'], 'skipped', 'original empty or identical to revised', orig)); continue
        if '\n' in orig:
            report.append((ch['no'], 'skipped', 'spans more than one paragraph: apply by hand', orig)); continue
        hits = []
        for p in all_paragraphs(doc):
            t = norm(para_text(p))
            start = t.find(norm(orig))
            while start != -1:
                hits.append((p, start))
                start = t.find(norm(orig), start + 1)
        if not hits:
            report.append((ch['no'], 'NOT FOUND', 'text not found exactly (check spelling, quotes, footnote marks)', orig)); continue
        if len(hits) > 1:
            report.append((ch['no'], 'found %d times' % len(hits), 'not applied: give more context', orig)); continue
        p, start = hits[0]
        a = norm(para_text(p))[start:start + len(orig)]
        pre, suf = diff_region(a, norm(new))
        new_mid = new[pre:len(new) - suf]
        res = rev.revise(p, start + pre, start + len(orig) - suf, new_mid)
        if res is None:
            report.append((ch['no'], 'skipped', 'change crosses a tab, field or footnote mark: apply by hand', orig)); continue
        if rule and res:
            try:
                from docx.text.run import Run
                doc.add_comment([Run(r, None) for r in res if r.tag == qn('w:r')] or None, text='[IPBES] ' + rule,
                                author=author, initials='IP')
            except Exception:
                pass
        report.append((ch['no'], 'applied', rule, orig))
    doc.save(out_path)
    return report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('document'); ap.add_argument('changes'); ap.add_argument('output')
    ap.add_argument('--author', default='IPBES editorial assistant')
    a = ap.parse_args()
    report = apply(a.document, a.changes, a.output, a.author)
    applied = sum(1 for r in report if r[1] == 'applied')
    print('Written: %s  (%d of %d changes applied as tracked changes)' % (a.output, applied, len(report)))
    for no, status, note, orig in report:
        if status != 'applied':
            print('  %-4s %-15s %s: %s' % (no, status, note, orig[:70]))


if __name__ == '__main__':
    main()
