"""Helpers for building IPBES PowerPoint decks from the standard IPBES slides.

Start from a copy of the standard deck (for example "Standard Slides_IPBES
Intro.pptx"), keep the template slides you need, add new content slides in the
house style and drop the rest. See references/presentations.md.

    from ipbes_pptx import Deck
    deck = Deck(DEFAULT_TEMPLATE)   # assets/Standard Slides_IPBES Intro.pptx
    cover = deck.template_slide(1)
    deck.set_cover_title(cover, 'Title of the presentation', 'Subtitle')
    s = deck.new_slide('1. Major cost categories')
    deck.question(s, 'Could the secretariat share ...')
    deck.bullets(s, 0.6, 2.4, 5.6, 2.9, ['First point', 'Second point'])
    deck.finish([cover, s, deck.template_slide(41)], 'out.pptx')

Needs python-pptx (pip install python-pptx). Standard library otherwise.
"""
import copy
import os

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

DEFAULT_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'assets',
                                'Standard Slides_IPBES Intro.pptx')

# House style of the standard IPBES slides.
TEAL = RGBColor(0x03, 0x74, 0x73)      # titles, headings, table headers, bullets
GREY = RGBColor(0x59, 0x59, 0x59)      # body text
ORANGE = RGBColor(0xED, 0x7D, 0x31)    # theme accent2: highlights (fills, large text only)
LIGHT = RGBColor(0xE6, 0xF1, 0xF1)     # light teal tint: question boxes, cards, total rows
ROW = RGBColor(0xF2, 0xF2, 0xF2)       # alternate table rows
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = 'Helvetica LT Std'              # as in the template; PowerPoint falls back to Arial
NO_STYLE_TABLE = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'  # "No Style, No Grid"

# Slides of assets/Standard Slides_IPBES Intro.pptx (8 slides; numbers are 1-based).
# Check them with references/presentations.md section 3 if the standard deck changes.
TEMPLATE_SLIDES = {
    'cover': 1,
    'outline': 2,          # the outline of the standard IPBES introduction
    'what_is_ipbes': 4,
    'thank_you': 8,
}
# Section dividers (big number, photo, title), used in this order by set_section().
SECTION_DIVIDERS = [3, 5, 6, 7]
# A plain content slide whose background (teal band + ipbes watermark) is copied
# onto every new slide. The band is a slide background, not part of the layout.
BACKGROUND_SOURCE = 2


def style_run(run, size, color=GREY, bold=False, italic=False):
    """Font, size and colour of one run, with the template's East Asian/complex fonts."""
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set('typeface', 'Helvetica Neue')


def set_bullet(paragraph, level=0):
    """Template bullets: teal Wingdings square at level 0, a dot below it."""
    ppr = paragraph._p.get_or_add_pPr()
    ppr.set('marL', str(int(Inches(0.25 + 0.25 * level))))
    ppr.set('indent', str(int(-Inches(0.22))))
    for tag in ('a:buNone', 'a:buChar', 'a:buFont', 'a:buSzPct', 'a:buClr'):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    clr = etree.SubElement(ppr, qn('a:buClr'))
    etree.SubElement(clr, qn('a:srgbClr')).set('val', '037473')
    etree.SubElement(ppr, qn('a:buSzPct')).set('val', '100000')
    font = etree.SubElement(ppr, qn('a:buFont'))
    char = etree.SubElement(ppr, qn('a:buChar'))
    if level == 0:
        font.set('typeface', 'Wingdings')
        char.set('char', '§')
    else:
        font.set('typeface', 'Arial')
        char.set('char', '•')


def no_bullet(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn('a:buNone')) is None:
        etree.SubElement(ppr, qn('a:buNone'))


def add_paragraphs(tf, items, size=12, space_after=4):
    """Fill a text frame.

    items: list of str or (str, opts). str items are bulleted. opts keys:
    bullet (default True), level, bold, italic, color, size, space_after,
    align (default left), runs = [(text, run_opts)] with run_opts taking
    bold/italic/color/size/link.
    """
    for i, item in enumerate(items):
        text, opts = (item, {}) if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(opts.get('space_after', space_after))
        p.alignment = opts.get('align', PP_ALIGN.LEFT)
        for rtext, ropts in opts.get('runs') or [(text, {})]:
            r = p.add_run()
            r.text = rtext
            style_run(r, ropts.get('size', opts.get('size', size)),
                      ropts.get('color', opts.get('color', GREY)),
                      ropts.get('bold', opts.get('bold', False)),
                      ropts.get('italic', opts.get('italic', False)))
            if 'link' in ropts:
                r.hyperlink.address = ropts['link']
        if opts.get('bullet', True):
            set_bullet(p, opts.get('level', 0))
        else:
            no_bullet(p)
    return tf


def money(v, cents=False):
    """$1,234 (UN style: "$", never "US$"). cents=True keeps two decimals."""
    return f'${v:,.2f}' if cents else f'${v:,.0f}'


class Deck:
    """A copy of the standard deck being turned into a new presentation."""

    def __init__(self, template_path=DEFAULT_TEMPLATE):
        self.prs = Presentation(template_path)
        self._template = list(self.prs.slides)
        bg_slide = self._template[BACKGROUND_SOURCE - 1]
        self._bg_slide = bg_slide
        self._bg = bg_slide._element.cSld.find(qn('p:bg'))
        self.content_layout = bg_slide.slide_layout

    # -- template slides ----------------------------------------------------
    def template_slide(self, number):
        """A slide of the standard deck, 1-based as in PowerPoint."""
        return self._template[number - 1]

    def set_cover_title(self, cover, title, subtitle=None):
        """Replace the cover title (26pt bold white) and optional subtitle (16pt light)."""
        p = cover.shapes.title.text_frame.paragraphs[0]
        first_rpr = copy.deepcopy(p._p.find(qn('a:r')).find(qn('a:rPr')))
        for el in list(p._p):
            if el.tag in (qn('a:r'), qn('a:br'), qn('a:fld')):
                p._p.remove(el)
        end = p._p.find(qn('a:endParaRPr'))

        def make(tag, size, text=None, light=False):
            el = etree.Element(qn(tag))
            rpr = copy.deepcopy(first_rpr)
            rpr.set('sz', str(size * 100))
            if light:
                rpr.set('b', '0')
                rpr.find(qn('a:latin')).set('typeface', 'Helvetica LT Std Light')
            el.append(rpr)
            if text is not None:
                etree.SubElement(el, qn('a:t')).text = text
            return el

        parts = [make('a:r', 26, title)]
        if subtitle:
            parts += [make('a:br', 26), make('a:br', 16), make('a:r', 16, subtitle, light=True)]
        for el in parts:
            if end is not None:
                end.addprevious(el)
            else:
                p._p.append(el)

    def set_section(self, slide, number, title):
        """Put a new number and title on a section divider, keeping the template's formatting."""
        def replace(shape, text):
            p = shape.text_frame.paragraphs[0]
            first = p.runs[0]
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
            first.text = text
            for extra in shape.text_frame.paragraphs[1:]:
                extra._p.getparent().remove(extra._p)
        replace(slide.shapes.title, title)
        for sh in slide.shapes:
            if (sh.has_text_frame and not sh.is_placeholder and sh.text_frame.paragraphs[0].runs
                    and (sh.text_frame.paragraphs[0].runs[0].font.size or 0) >= Pt(60)):
                replace(sh, number)

    # -- new slides -------------------------------------------------------
    def new_slide(self, title, size=24):
        """Content slide with the template background and a house-style title.

        The title box has room for two lines and is bottom-aligned, so the last
        line sits at the same height on every slide. Content starts at y=1.6".
        Use size=21-22 only for titles that would otherwise need three lines.
        """
        s = self.prs.slides.add_slide(self.content_layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        bg = copy.deepcopy(self._bg)
        for blip in bg.iter(qn('a:blip')):
            part = self._bg_slide.part.related_part(blip.get(qn('r:embed')))
            blip.set(qn('r:embed'), s.part.relate_to(part, RT.IMAGE))
        s._element.cSld.insert(0, bg)
        tb = self.textbox(s, 0.5, 0.62, 9.0, 0.88, 'Title 1', anchor=MSO_ANCHOR.BOTTOM)
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = title
        style_run(r, size, TEAL, bold=True)
        return s

    def textbox(self, slide, x, y, w, h, name, anchor=MSO_ANCHOR.TOP, margin=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        if margin is not None:
            tf.margin_left = tf.margin_right = Inches(margin)
            tf.margin_top = tf.margin_bottom = Inches(margin * 0.6)
        return tb

    def bullets(self, slide, x, y, w, h, items, size=14, space_after=10, name='Body'):
        tb = self.textbox(slide, x, y, w, h, name)
        add_paragraphs(tb.text_frame, items, size=size, space_after=space_after)
        return tb

    def numbered(self, slide, x, y, w, h, items, size=11, name='Numbered list'):
        """Agenda-style list: teal bold numbers with hanging indents.

        items: list of (number, text, level), e.g. ('1.', 'Opening', 0), ('(a)', 'Agenda', 1).
        text may be a list of (run_text, run_opts) for bold or linked parts.
        """
        tb = self.textbox(slide, x, y, w, h, name)
        tf = tb.text_frame
        for i, (num, text, lvl) in enumerate(items):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.space_after = Pt(2)
            indent = Inches(0.4 if lvl == 0 else 0.35)
            left = Inches(0.4 + 0.35 * lvl)
            ppr = para._p.get_or_add_pPr()
            ppr.set('marL', str(int(left)))
            ppr.set('indent', str(int(-indent)))
            etree.SubElement(ppr, qn('a:buNone'))
            tabs = etree.SubElement(ppr, qn('a:tabLst'))
            etree.SubElement(tabs, qn('a:tab')).set('pos', str(int(left)))
            r = para.add_run()
            r.text = num + '\t'
            style_run(r, size, TEAL, bold=True)
            for rtext, ropts in (text if isinstance(text, list) else [(text, {})]):
                r = para.add_run()
                r.text = rtext
                style_run(r, size, ropts.get('color', GREY), ropts.get('bold', False), ropts.get('italic', False))
                if 'link' in ropts:
                    r.hyperlink.address = ropts['link']
        return tb

    def card(self, slide, x, y, w, h, name, fill=LIGHT, items=None, anchor=MSO_ANCHOR.TOP, pill=False):
        """Rounded box with no outline. pill=True for fully rounded chips."""
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.name = name
        shp.adjustments[0] = 0.5 if pill else 0.08
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
        shp.line.fill.background()
        shp.shadow.inherit = False
        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05 if pill else 0.15)
        tf.margin_top = tf.margin_bottom = Inches(0.1)
        if items:
            add_paragraphs(tf, items)
        return shp

    def question(self, slide, text, y=1.6, h=0.62):
        """The audience's question in a tinted box under the title (10pt italic).

        Shorten long questions here and put the full wording in the notes.
        """
        shp = self.card(slide, 0.6, y, 8.8, h, 'Question', anchor=MSO_ANCHOR.MIDDLE)
        add_paragraphs(shp.text_frame, [(None, {'bullet': False, 'space_after': 0, 'runs': [
            ('Question: ', {'bold': True, 'color': TEAL, 'size': 10}),
            (text, {'italic': True, 'size': 10}),
        ]})])
        return shp

    def table(self, slide, x, y, w, rows, col_w, name, size=10, header_size=10,
              bold_rows=(), right_cols=()):
        """Teal header row, alternating grey rows; bold_rows get a teal tint (totals)."""
        row_h = 0.25 if size <= 9 else 0.27
        gf = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(row_h * len(rows)))
        gf.name = name
        tbl = gf.table
        tbl._tbl.tblPr.find(qn('a:tableStyleId')).text = NO_STYLE_TABLE
        for c, cw in enumerate(col_w):
            tbl.columns[c].width = Inches(cw)
        for r, row in enumerate(rows):
            tbl.rows[r].height = Inches(row_h)
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0.06)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                header, strong = r == 0, r in bold_rows
                cell.fill.solid()
                cell.fill.fore_color.rgb = TEAL if header else (LIGHT if strong else (ROW if r % 2 == 0 else WHITE))
                p = cell.text_frame.paragraphs[0]
                cell.text_frame.word_wrap = True
                p.alignment = PP_ALIGN.RIGHT if (c in right_cols and not header) else PP_ALIGN.LEFT
                cell_size = header_size if header else size
                # Empty cells need a size too, or they fall back to 18pt and stretch the row.
                etree.SubElement(p._p, qn('a:endParaRPr')).set('sz', str(int(cell_size * 100)))
                if str(val):
                    run = p.add_run()
                    run.text = str(val)
                    style_run(run, cell_size, WHITE if header else (TEAL if strong else GREY), bold=header or strong)
        return gf

    def picture(self, slide, path, y=1.6, h=3.85, name='Picture', alt=''):
        """Picture centred horizontally under the title. Always give alt text."""
        pic = slide.shapes.add_picture(path, 0, Inches(y), height=Inches(h))
        pic.left = int((self.prs.slide_width - pic.width) / 2)
        pic.name = name
        pic._element.nvPicPr.cNvPr.set('descr', alt)
        return pic

    def column_chart(self, slide, x, y, w, h, categories, values, title, highlight=None,
                     name='Chart', number_format='#,##0'):
        """Native column chart in the house style: teal bars, data labels, no gridlines.

        highlight: index of one bar to colour orange. Use "\n" in a category to
        break its label over two lines.
        """
        cd = CategoryChartData()
        cd.categories = categories
        cd.add_series(title, values)
        gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        gf.name = name
        ch = gf.chart
        ch.has_legend = False
        ch.has_title = True
        ch.chart_title.text_frame.text = title
        run = ch.chart_title.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb, run.font.name = Pt(11), True, TEAL, FONT
        plot = ch.plots[0]
        plot.gap_width = 70
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size, dl.font.bold, dl.font.color.rgb, dl.font.name = Pt(11), True, GREY, FONT
        dl.number_format, dl.number_format_is_linked = number_format, False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = TEAL
        if highlight is not None:
            pt = ser.points[highlight]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = ORANGE
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        va.minimum_scale = 0
        va.maximum_scale = max(values) * 1.15  # room for the labels above the bars
        ca = ch.category_axis
        ca.tick_labels.font.size, ca.tick_labels.font.color.rgb, ca.tick_labels.font.name = Pt(10), GREY, FONT
        ca.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
        return gf

    @staticmethod
    def notes(slide, text):
        slide.notes_slide.notes_text_frame.text = text

    # -- output -------------------------------------------------------------
    def finish(self, order, out_path):
        """Keep only the slides in `order`, in that order, and save.

        Dropped slides (and media only they used) are left out of the file, so
        an 89 MB standard deck becomes a few MB.
        """
        keep = {id(s) for s in order}
        sld_lst = self.prs.slides._sldIdLst
        by_slide = {}
        for sld_id in list(sld_lst):
            slide = self.prs.slides.part.related_part(sld_id.rId).slide
            by_slide[id(slide)] = sld_id
            if id(slide) not in keep:
                self.prs.part.drop_rel(sld_id.rId)
            sld_lst.remove(sld_id)
        for s in order:
            sld_lst.append(by_slide[id(s)])
        self.prs.save(out_path)
        return out_path
