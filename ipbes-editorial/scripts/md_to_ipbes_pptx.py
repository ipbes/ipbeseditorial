"""Convert a plain-text slide outline into a PowerPoint deck in the IPBES style.

The outline format is described in references/presentations.md, section 2, and
test/ipbes-test-slides.md is a full example. Each slide is laid out from what it
contains (points, a key figure, cards, a table, a chart, a picture, a question),
on a copy of the standard IPBES slides.

Usage:
    python3 md_to_ipbes_pptx.py slides.md out.pptx [--template other-standard-deck.pptx]

By default the standard deck in ../assets/ is used.

Image paths in the outline are relative to the outline file. The script makes no
language edits: correct the text first (SKILL.md, request D). It prints a
warning for any slide whose text may not fit; split that slide and run again.
Needs python-pptx.
"""
import argparse
import math
import os
import re
import sys

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ipbes_pptx import (SECTION_DIVIDERS, TEAL, TEMPLATE_SLIDES, WHITE, Deck, add_paragraphs)  # noqa: E402

DEFAULT_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'assets',
                                'Standard Slides_IPBES Intro.pptx')
SUPERSCRIPT = str.maketrans('0123456789', '⁰¹²³⁴⁵⁶⁷⁸⁹')
X0, W = 0.6, 8.8              # content area left edge and width (inches)
TOP, BOTTOM = 1.6, 5.3        # content area under the title
GAP = 0.2


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def runs(text):
    """Inline markup -> [(text, opts)]: **bold**, [text](url), [^1] footnote markers."""
    text = re.sub(r'\[\^(\d+)\]', lambda m: m.group(1).translate(SUPERSCRIPT), text)
    out = []
    for part in re.split(r'(\*\*[^*]+\*\*|\[[^\]]+\]\([^)]+\))', text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            out.append((part[2:-2], {'bold': True}))
        elif part.startswith('['):
            label, url = re.match(r'\[([^\]]+)\]\(([^)]+)\)', part).groups()
            out.append((label, {'color': TEAL, 'link': url}))
        else:
            out.append((part, {}))
    return out


def plain(text):
    return ''.join(t for t, _ in runs(text))


def parse(path):
    """Return (cover, slides). Each slide is a dict of the blocks it contains."""
    lines = open(path, encoding='utf-8').read().splitlines()
    cover = {'title': None, 'subtitle': []}
    slides, cur, mode = [], None, None
    base = os.path.dirname(os.path.abspath(path))

    def new(title):
        return {'title': title, 'text': [], 'numbered': [], 'stat': None, 'cards': [], 'chips': None,
                'chart': None, 'tables': [], 'pictures': [], 'question': None, 'notes': [],
                'footnotes': [], 'template': None}

    for raw in lines:
        line = raw.rstrip()
        s = line.strip()
        if s.startswith('<!--'):
            continue
        m = re.match(r'^\[template:\s*(\w+)\]$', s)
        if m:
            slides.append(dict(new(None), template=m.group(1)))
            cur, mode = None, None
            continue
        m = re.match(r'^\[section\]\s*(?:(\d+)\.?\s+)?(.+)$', s)
        if m:
            slides.append(dict(new(m.group(2).strip()), section=m.group(1) or ''))
            cur, mode = None, None
            continue
        if line.startswith('## '):
            cur = new(line[3:].strip())
            slides.append(cur)
            mode = None
            continue
        if line.startswith('# '):
            cover['title'] = line[2:].strip()
            cur, mode = None, 'cover'
            continue
        if not s:
            if mode in ('cards', 'chart', 'table'):
                mode = None
            continue
        if mode == 'cover' and cur is None:
            cover['subtitle'].append(s)
            continue
        if cur is None:
            continue
        if s.startswith('[^') and ']:' in s:
            num, text = re.match(r'\[\^(\d+)\]:\s*(.*)', s).groups()
            cur['footnotes'].append(runs(num.translate(SUPERSCRIPT) + ' ' + text))
            continue
        if s.lower().startswith('notes:'):
            cur['notes'].append(s[6:].strip())
            continue
        if s.startswith('? '):
            cur['question'] = s[2:].strip()
            continue
        m = re.match(r'^\[stat\]\s*(.+?)\s*\|\s*(.+)$', s)
        if m:
            cur['stat'] = (m.group(1), m.group(2))
            continue
        m = re.match(r'^\[chips(?::\s*([^\]]+))?\]\s*(.+)$', s)
        if m:
            cur['chips'] = (m.group(1), [c.strip() for c in m.group(2).split('|')])
            continue
        if s == '[cards]':
            mode = 'cards'
            continue
        m = re.match(r'^\[chart\]\s*(.+)$', s)
        if m:
            cur['chart'] = {'title': m.group(1), 'cats': [], 'vals': [], 'highlight': None}
            mode = 'chart'
            continue
        m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)$', s)
        if m:
            cur['pictures'].append((m.group(1), os.path.join(base, m.group(2))))
            continue
        if mode == 'chart':
            m = re.match(r'^-\s*(.+?):\s*([\d,.]+)\s*(\*)?$', s)
            if m:
                ch = cur['chart']
                if m.group(3):
                    ch['highlight'] = len(ch['vals'])
                ch['cats'].append(m.group(1).replace(' / ', '\n'))
                ch['vals'].append(float(m.group(2).replace(',', '')))
                continue
        if s.startswith('|'):
            if mode != 'table':
                cur['tables'].append([])
                mode = 'table'
            if re.match(r'^\|[\s:|-]+\|$', s):
                continue
            cur['tables'][-1].append([c.strip() for c in s.strip('|').split('|')])
            continue
        if mode == 'cards':
            m = re.match(r'^(\s*)[-*]\s+(.*)$', line)
            if m and not m.group(1):
                head, _, body = m.group(2).partition(':')
                head = head.strip().strip('*').strip()
                cur['cards'].append({'head': head, 'body': [body.strip()] if body.strip() else []})
                continue
            if m and cur['cards']:
                cur['cards'][-1]['body'].append(m.group(2).strip())
                continue
        m = re.match(r'^(\s*)(\d+\.|\([a-z]\)|[a-z]\))\s+(.*)$', line)
        if m:
            cur['numbered'].append((m.group(2), m.group(3), 1 if m.group(1) else 0))
            continue
        m = re.match(r'^(\s*)[-*]\s+(.*)$', line)
        if m:
            cur['text'].append((m.group(2), len(m.group(1)) // 2, True))
            continue
        cur['text'].append((s, 0, False))
    return cover, slides


# ---------------------------------------------------------------------------
# Measuring (rough: Helvetica averages about half an em per character)
# ---------------------------------------------------------------------------
def lines_for(text, width, size):
    per_line = max(1, int(width * 144 / size))
    return max(1, math.ceil(len(text) / per_line))


def text_height(items, width, size, space_after):
    h = 0.15
    for text, level, bullet in items:
        w = width - 0.25 * (level + 1) - 0.2
        h += lines_for(plain(text), w, size) * size * 1.2 / 72 + space_after / 72
    return h


def table_height(rows, col_w, size):
    h = 0
    for r in rows:
        n = max(lines_for(plain(c), cw - 0.12, size) for c, cw in zip(r, col_w))
        h += max(0.25 if size <= 9 else 0.27, n * size * 1.2 / 72 + 0.08)
    return h


def table_widths(rows, width):
    ncol = len(rows[0])
    lens = [max(len(plain(r[c])) for r in rows if c < len(r)) for c in range(ncol)]
    numeric = [all(re.match(r'^[\s$€£\d,.%()¹²³⁴⁵⁶⁷⁸⁹⁰-]*$', plain(r[c])) for r in rows[1:]) for c in range(ncol)]
    mins = [0.55 if numeric[c] else 1.0 for c in range(ncol)]
    raw = [max(mins[c], min(lens[c], 60) * 0.08) for c in range(ncol)]
    scale = width / sum(raw)
    return [w * scale for w in raw], numeric


# ---------------------------------------------------------------------------
# Building
# ---------------------------------------------------------------------------
def put_text(deck, slide, items, x, y, w, h, size):
    paras = []
    for text, level, bullet in items:
        paras.append((None, {'bullet': bullet, 'level': level, 'size': size if level == 0 else size - 1,
                             'space_after': 8 if size >= 13 else 5, 'runs': runs(text)}))
    tb = deck.textbox(slide, x, y, w, h, 'Text')
    add_paragraphs(tb.text_frame, paras)


def fit_size(items, width, height, sizes=(14, 13, 12, 11, 10)):
    for size in sizes:
        if text_height(items, width, size, 8 if size >= 13 else 5) <= height:
            return size, True
    return sizes[-1], False


def build_slide(deck, sd, number, warnings):
    title = sd['title']
    long_title = len(title) > 70
    slide = deck.new_slide(title, size=21 if long_title else 24)
    y = TOP
    bottom = BOTTOM
    if sd['footnotes']:
        bottom -= 0.15 + 0.17 * sum(max(1, math.ceil(len(plain_r(f)) / 150)) for f in sd['footnotes'])
    if sd['chips']:
        bottom -= 0.6

    has_content = any(sd[k] for k in ('text', 'numbered', 'stat', 'cards', 'chart', 'tables', 'pictures'))

    # Question: a box under the title, or a discussion slide when nothing answers it.
    if sd['question'] and not has_content:
        deck.card(slide, X0, 1.75, W, 1.35, 'Question', anchor=MSO_ANCHOR.MIDDLE, items=[
            ('Question', {'bullet': False, 'bold': True, 'color': TEAL, 'size': 12, 'space_after': 4}),
            (None, {'bullet': False, 'italic': True, 'size': 13, 'space_after': 0, 'runs': runs(sd['question'])})])
        deck.card(slide, 3.7, 3.5, 2.6, 0.55, 'For discussion', fill=TEAL, pill=True, anchor=MSO_ANCHOR.MIDDLE, items=[
            ('For discussion', {'bullet': False, 'bold': True, 'color': WHITE, 'size': 14, 'space_after': 0,
                                'align': PP_ALIGN.CENTER})])
        y = 4.2
    elif sd['question']:
        qh = 0.62 if len(sd['question']) > 160 else 0.45
        box = deck.card(slide, X0, y, W, qh, 'Question', anchor=MSO_ANCHOR.MIDDLE)
        add_paragraphs(box.text_frame, [(None, {'bullet': False, 'space_after': 0, 'runs':
                                                [('Question: ', {'bold': True, 'color': TEAL, 'size': 10})]
                                                + [(t, dict(o, italic=True, size=10)) for t, o in runs(sd['question'])]})])
        y += qh + GAP

    avail = bottom - y
    text = sd['text']
    side = sd['stat'] or sd['chart'] or (sd['pictures'] and (text or sd['numbered']))

    # Picture alone: centred, as large as the area allows.
    if sd['pictures'] and not (text or sd['numbered'] or sd['tables'] or sd['cards'] or sd['chart']):
        alt, path = sd['pictures'][0]
        pic = deck.picture(slide, path, y=y, h=avail, name='Picture', alt=alt)
        if pic.width > Inches(W):
            ratio = Inches(W) / pic.width
            pic.width, pic.height = Inches(W), int(pic.height * ratio)
            pic.left = int((deck.prs.slide_width - pic.width) / 2)
        y = bottom

    elif side:
        # Two columns: text or numbered list on the left, figure / chart / picture on the right.
        if sd['chart']:
            left_w, right_w = 5.3, W - 5.3 - GAP
        elif sd['stat']:
            left_w, right_w = 5.6, W - 5.6 - 0.4
        else:
            left_w, right_w = 4.4, W - 4.4 - GAP
        right_x = X0 + W - right_w
        if sd['chart']:
            ch = sd['chart']
            deck.column_chart(slide, X0, y, left_w, avail, ch['cats'], ch['vals'], ch['title'], highlight=ch['highlight'])
            if text:
                size, ok = fit_size(text, right_w - 0.3, avail - 0.3, (12, 11, 10))
                box_h = min(avail - 0.2, text_height(text, right_w - 0.3, size, 6) + 0.25)
                box = deck.card(slide, right_x, y + 0.1, right_w, box_h, 'Points')
                add_paragraphs(box.text_frame, [(None, {'bullet': b, 'level': lv, 'size': size, 'space_after': 6,
                                                        'runs': runs(t)}) for t, lv, b in text])
                if not ok:
                    warnings.append(f'slide {number} ("{title}"): the points beside the chart may not fit')
        else:
            if sd['numbered']:
                deck.numbered(slide, X0 - 0.1, y, left_w, avail, [(n, runs(t), lv) for n, t, lv in sd['numbered']])
            used_h = 0
            if text:
                size, ok = fit_size(text, left_w, avail)
                put_text(deck, slide, text, X0, y, left_w, avail, size)
                used_h = text_height(text, left_w, size, 8 if size >= 13 else 5)
                if not ok:
                    warnings.append(f'slide {number} ("{title}"): the text may not fit; split the slide')
            if sd['stat']:
                value, label = sd['stat']
                deck.card(slide, right_x, y + 0.1, right_w, min(2.5, avail - 0.2), 'Key figure', anchor=MSO_ANCHOR.MIDDLE, items=[
                    (value, {'bullet': False, 'bold': True, 'color': TEAL, 'size': 54 if len(value) <= 5 else 36,
                             'space_after': 0, 'align': PP_ALIGN.CENTER}),
                    (None, {'bullet': False, 'size': 12, 'space_after': 0, 'align': PP_ALIGN.CENTER, 'runs': runs(label)})])
            elif sd['pictures']:
                alt, path = sd['pictures'][0]
                pic = deck.picture(slide, path, y=y, h=avail, name='Picture', alt=alt)
                if pic.width > Inches(right_w):
                    ratio = Inches(right_w) / pic.width
                    pic.width, pic.height = Inches(right_w), int(pic.height * ratio)
                pic.left = int(Inches(right_x))
        y = bottom if (sd['chart'] or sd['pictures']) else y + max(used_h, 2.6) + GAP

    else:
        # Stacked full width: text, numbered list, cards, tables.
        blocks = [k for k in ('text', 'numbered', 'cards', 'tables') if sd[k]]
        if sd['numbered']:
            n = len(sd['numbered'])
            size = 11 if n <= 12 else 10
            need = sum(lines_for(plain(t), W - 0.6, size) for _, t, _ in sd['numbered']) * size * 1.3 / 72 + 0.1
            if sd['cards'] and len(sd['cards']) == 1:
                # A numbered list with one side card (related items).
                card = sd['cards'][0]
                deck.numbered(slide, X0 - 0.1, y, 6.0, avail, [(m, runs(t), lv) for m, t, lv in sd['numbered']], size=size - 1 if n > 12 else size)
                deck.card(slide, X0 + 6.25, y + 0.1, W - 6.25, min(avail, 0.6 + 0.45 * len(card['body'])), 'Side card', items=
                          [(card['head'], {'bullet': False, 'bold': True, 'color': TEAL, 'size': 12, 'space_after': 6})]
                          + [(None, {'size': 11, 'space_after': 5, 'runs': runs(b)}) for b in card['body']])
                sd = dict(sd, cards=[])
                need = sum(lines_for(plain(t), 5.4, size) for _, t, _ in sd['numbered']) * size * 1.3 / 72 + 0.1
            else:
                deck.numbered(slide, X0 - 0.1, y, W, need, [(m, runs(t), lv) for m, t, lv in sd['numbered']], size=size)
            if need > avail:
                warnings.append(f'slide {number} ("{title}"): the numbered list may not fit; split the slide')
            y += min(need, avail) + GAP
        if text:
            remaining = bottom - y
            share = remaining if len(blocks) == 1 else remaining * 0.45
            size, ok = fit_size(text, W, share)
            h = text_height(text, W, size, 8 if size >= 13 else 5)
            put_text(deck, slide, text, X0, y, W, min(h, remaining), size)
            if not ok:
                warnings.append(f'slide {number} ("{title}"): the text may not fit; split the slide')
            y += min(h, remaining) + 0.1
        if sd['cards']:
            cards = sd['cards']
            n = len(cards)
            per_row = n if n <= 3 else 4
            cw = (W - (per_row - 1) * 0.15) / per_row
            nrows = math.ceil(n / per_row)
            remaining = bottom - y
            ch = min(2.6, (remaining - (nrows - 1) * 0.15) / nrows)
            need = max(0.6 + sum(lines_for(plain(b), cw - 0.5, 11) * 0.22 + 0.08 for b in c['body']) for c in cards)
            if need > ch + 0.05:
                warnings.append(f'slide {number} ("{title}"): the cards may not fit; shorten them or split the slide')
            ch = min(ch, max(need, 0.7))
            for k, c in enumerate(cards):
                cx = X0 + (k % per_row) * (cw + 0.15)
                cy = y + (k // per_row) * (ch + 0.15)
                items = [(c['head'], {'bullet': False, 'bold': True, 'color': TEAL, 'size': 13 if per_row <= 3 else 12,
                                      'space_after': 6 if c['body'] else 0})]
                items += [(None, {'bullet': per_row <= 3, 'size': 11 if per_row <= 3 else 10, 'space_after': 5,
                                  'runs': runs(b)}) for b in c['body']]
                deck.card(slide, cx, cy, cw, ch, f'Card: {plain(c["head"])}',
                          anchor=MSO_ANCHOR.TOP if per_row <= 3 else MSO_ANCHOR.MIDDLE, items=items)
            y += nrows * ch + (nrows - 1) * 0.15 + GAP
        if sd['tables']:
            tables = sd['tables']
            tw = W if len(tables) == 1 else (W - GAP) / 2
            size = 10 if len(tables) == 1 else 9
            for k, rows in enumerate(tables):
                ncol = max(len(r) for r in rows)
                rows = [r + [''] * (ncol - len(r)) for r in rows]
                col_w, numeric = table_widths(rows, tw)
                bold_rows = tuple(i for i, r in enumerate(rows) if i and r[0].startswith('**'))
                clean = [[plain(c) for c in r] for r in rows]
                right = tuple(c for c in range(1, ncol) if numeric[c])
                tx = X0 + k * (tw + GAP)
                deck.table(slide, tx, y, tw, clean, col_w, f'Table {k + 1}', size=size, header_size=size,
                           bold_rows=bold_rows, right_cols=right)
                if y + table_height(clean, col_w, size) > bottom + 0.05:
                    warnings.append(f'slide {number} ("{title}"): table {k + 1} may not fit; split it or the slide')
            y += max(table_height([[plain(c) for c in r] for r in t], table_widths(t, tw)[0], size) for t in tables) + GAP

    # Chips: a row of labels, optionally with a label on the left.
    if sd['chips']:
        label, chips = sd['chips']
        cy = min(y + 0.1, BOTTOM - 0.5 - (0 if not sd['footnotes'] else 0.15 + 0.17 * len(sd['footnotes'])))
        cx = X0
        if label:
            lb = deck.textbox(slide, X0, cy, 1.55, 0.42, 'Chips label', anchor=MSO_ANCHOR.MIDDLE, margin=0)
            add_paragraphs(lb.text_frame, [(label, {'bullet': False, 'bold': True, 'color': TEAL, 'size': 11})])
            cx = X0 + 1.6
        cw = min(1.8, (X0 + W - cx - 0.12 * (len(chips) - 1)) / len(chips))
        for k, chip in enumerate(chips):
            deck.card(slide, cx + k * (cw + 0.12), cy, cw, 0.42, f'Chip: {chip}', fill=TEAL, pill=True,
                      anchor=MSO_ANCHOR.MIDDLE, items=[(chip, {'bullet': False, 'bold': True, 'color': WHITE,
                                                              'size': 10.5, 'space_after': 0, 'align': PP_ALIGN.CENTER})])
        y = cy + 0.42 + GAP

    # Footnotes: one 9 pt block at the bottom.
    if sd['footnotes']:
        fh = 0.15 + 0.17 * sum(max(1, math.ceil(len(plain_r(f)) / 150)) for f in sd['footnotes'])
        fb = deck.textbox(slide, X0, min(y, BOTTOM - fh + 0.1), W, fh, 'Footnotes', margin=0)
        add_paragraphs(fb.text_frame, [(None, {'bullet': False, 'size': 9, 'space_after': 0,
                                               'runs': [(t, dict(o, size=9)) for t, o in f]}) for f in sd['footnotes']])
    if sd['notes']:
        deck.notes(slide, '\n'.join(sd['notes']))
    return slide


def plain_r(run_list):
    return ''.join(t for t, _ in run_list)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('outline')
    ap.add_argument('output')
    ap.add_argument('--template', default=DEFAULT_TEMPLATE,
                    help='the standard IPBES slides (default: assets/Standard Slides_IPBES Intro.pptx)')
    args = ap.parse_args()

    if not os.path.exists(args.template):
        sys.exit(f'standard deck not found: {args.template}')
    cover, slides = parse(args.outline)
    deck = Deck(args.template)
    order, used, warnings = [], set(), []
    if cover['title']:
        c = deck.template_slide(TEMPLATE_SLIDES['cover'])
        deck.set_cover_title(c, plain(cover['title']), plain(' '.join(cover['subtitle'])) or None)
        order.append(c)
        used.add('cover')
    dividers = list(SECTION_DIVIDERS)
    for sd in slides:
        number = len(order) + 1
        if 'section' in sd:
            if not dividers:
                sys.exit(f'more than {len(SECTION_DIVIDERS)} [section] slides; the standard deck has only '
                         f'{len(SECTION_DIVIDERS)} section dividers')
            div = deck.template_slide(dividers.pop(0))
            deck.set_section(div, (sd['section'] + '.') if sd['section'] else '', plain(sd['title']))
            order.append(div)
            continue
        if sd['template']:
            name = sd['template']
            if name not in TEMPLATE_SLIDES:
                sys.exit(f'unknown template slide "{name}"; use one of: {", ".join(TEMPLATE_SLIDES)}')
            if name in used:
                sys.exit(f'template slide "{name}" is used twice; a template slide can be used only once')
            used.add(name)
            order.append(deck.template_slide(TEMPLATE_SLIDES[name]))
            continue
        order.append(build_slide(deck, sd, number, warnings))
    deck.finish(order, args.output)
    print(f'{args.output} written ({len(order)} slides)')
    for w in warnings:
        print('WARNING:', w)


if __name__ == '__main__':
    main()
