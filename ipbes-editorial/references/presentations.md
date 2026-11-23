# IPBES presentations (PowerPoint)

How to turn plain text and the images a colleague provides into a PowerPoint deck that is laid out in the IPBES style of the standard slides and checked against the house style. The converter is `scripts/md_to_ipbes_pptx.py`; it uses the helpers in `scripts/ipbes_pptx.py`. A full sample outline is `test/ipbes-test-slides.md`.

## 1. Workflow: from text and images to a deck

**Input:** the text of the presentation (an outline, notes or a document, in any form) and any images to use. **Output:** a `.pptx` deck, a change table for the language corrections, and queries for the author.

1. **Read everything first.** Identify the title, the sections, the points of each slide, the figures and the images. Note what is missing: a section without content, a table without headers, a figure without a source, an image without a caption.
2. **Check the language** with the same rules as any IPBES text (`house-style.md`, `ipbes-terms.md`, `countries.csv`, `inclusive-language.md`) and the editing rules of SKILL.md: change only what a rule requires and never change meaning, figures or names. Put voice problems ("must", "should") in the queries. On slides also:
   - titles in sentence case, with no full stop at the end;
   - list items in a parallel form, with no full stops after items that are not sentences;
   - "$", not "US$"; "per cent" in sentences, "%" acceptable in a table, chart or stat label;
   - full names of IPBES bodies at first mention, abbreviations only where the house style allows.

   Record every change in the change table (`No. | Original | Revised | Rule`).
3. **Fit the text without cutting it.** Slides take little text. If a slide's text does not fit, split it over two slides or move detail to the **speaker notes**; never shorten or paraphrase silently. If you show a shortened form on the slide (for example a long question), keep the full wording in the notes and list the shortening in the change table.
4. **Write the outline.** Rewrite the text in the outline format (section 2): one `##` heading per slide, and a marker for each kind of content (points, numbered list, cards, table, chart, key figure, picture, question, footnotes, notes). Use template slides where the content matches them (section 3). Vary the patterns: do not put every slide in bullets. Apply the language corrections in the outline.
5. **Place the images.**
   - Use only the images provided, or photos and logos already in the standard deck. Never generate, redraw or distort images or logos.
   - Keep the aspect ratio. Scale an image to fit the content area (up to 8.8 × 3.85 inches under the title).
   - Give every image alt text from its caption or content.
   - An image with a lot of small text (a schedule, a table as a picture) gets a slide of its own.
6. **Convert** the outline (section 7):

   ```
   python3 scripts/md_to_ipbes_pptx.py slides.md slides.pptx
   ```

   The script lays out each slide from what it contains and prints a warning for any slide whose text may not fit. Split that slide in the outline and convert again; do not shrink text below the sizes in section 4.
7. **Check** the text, the file and every slide visually (section 9).
8. **Return:**
   - the deck;
   - the change table and a short summary with counts;
   - **queries for the author**: missing content, unlabelled or inconsistent figures, voice, anything you had to choose (for example table headers);
   - a note of any text moved to the speaker notes and any image not used.

Never invent content. A section with no content becomes a slide that says so (for example a question with a "For discussion" label); do not fill it and do not leave empty boxes.

## 2. The outline format

The outline is plain text with a few markers, close to Markdown. Write one block per slide; the converter chooses the layout from the blocks.

| Write | Gives |
|---|---|
| `# Title` and the lines under it | The cover: title, then the subtitle |
| `## 1. Slide title` | A new slide. Number titles in the text if the slides are numbered |
| `- Point` (and `  - sub-point`, two spaces in) | Bullets with teal squares; sub-points with dots |
| A line with no marker | A paragraph without a bullet |
| `**bold**`, `[link text](https://…)` | Bold text; a link |
| `1. Item`, then `   (a) Sub-item` | Numbered list with hanging indents (agenda style) |
| `[cards]` then `- **Heading**: text` or `- **Heading**` with `  - points` | Cards: three or fewer side by side, four or more in a grid. One card after a numbered list sits beside it |
| A Markdown table (`\| A \| B \|`, then `\|---\|---\|`) | A teal table; numbers right-aligned; a row whose first cell is `**bold**` is a total. Two tables on one slide sit side by side |
| `[^1]` in the text and `[^1]: Note` on its own line | Superscript markers and 9 pt footnotes under the content |
| `[chart] Chart title` then `- Label: 120` lines; ` / ` breaks a label; `*` after the number highlights the bar | A column chart, with any points of the slide in a card beside it |
| `[stat] 8% \| label` | A key figure in a card beside the points |
| `![Alt text](image.png)` | The picture: alone on the slide, or beside the points. The alt text is required |
| `? Question` | A question box under the title. With nothing else on the slide, a "For discussion" slide |
| `[chips: Label] A \| B \| C` | A row of teal labels, with an optional label on the left |
| `Notes: text` | Speaker notes |
| `[section] 2. Section title` | A section divider of the standard deck with this number and title. Each divider has a different photo; there are four |
| `[template: thank_you]` | A slide of the standard deck as it is: `what_is_ipbes`, `outline` (the outline of the standard IPBES introduction) or `thank_you`. Each one can be used once |

**Example** (an extract of `test/ipbes-test-slides.md`; all names and figures are fictitious):

```
# Planning a regional capacity-building workshop
Information for participating institutions

## 1. Purpose of the workshop
- Introduce the methodological assessment to experts from the region
- Strengthen the use of **Indigenous and local knowledge** in national assessments
  - Exchange experience between national platforms
- Background documents are on the [IPBES website](https://www.ipbes.net)
Notes: Explain that the workshop is part of objective 2 of the rolling work programme.

[section] 2. Organization of the workshop

## 2. Draft programme
1. Opening of the workshop.
2. Organizational matters:
   (a) Adoption of the programme;
   (b) Election of the co-chairs.
3. Closure of the workshop.

[cards]
- **Also organized around the workshop**
  - Field visit

## 3. Estimated budget
| Item | Quantity | Unit cost | Total |
|---|---|---|---|
| Venue rental[^1] | 3 days | $4,000 | $12,000 |
| Participant travel[^2] | 25 | $1,500 | $37,500 |
| **Total** | | | **$49,500** |

[^1]: Including audiovisual equipment.
[^2]: Economy class, average fare.

## 4. Funding arrangements
- Costs are shared between the host institution and the secretariat
- Unspent funds are refunded
[stat] 8% | programme support costs on funds transferred to the secretariat

## 5. Participation in previous workshops
[chart] Participants per workshop
- 2023 / Africa: 40
- 2025 / Latin America: 62 *

- Participation has grown each year

## 6. Division of responsibilities
[cards]
- **Host institution**
  - Venue and local logistics
- **Secretariat**
  - Programme and invitations

## 7. Practical arrangements
? Which documents do participants need to enter the country?
- Visa support letters are issued by the host institution
[chips: Participant categories] Experts | Fellows | Observers

## 8. Interpretation
? Will interpretation be provided in a second language?

## 9. Workshop location
![Venue of the workshop seen from the main entrance](venue.jpg)

[template: thank_you]
```

## 3. The standard deck

Always build on a copy of the standard deck, **Standard Slides_IPBES Intro.pptx** (about 12 MB, 8 slides, 16:9 at 10 × 5.625 inches), which is in `assets/`. The converter and `Deck()` use it by default; pass `--template` (or a path to `Deck`) only for a newer version. Do not build a deck from scratch: the cover, background band, watermark and logos live in the template.

Slides worth reusing as they are (numbers as in PowerPoint):

| Slide | Contents |
|---|---|
| 1 | Cover: landscape photo, title in the teal panel, partner logos (UNEP, UNESCO, FAO, UNDP). `set_cover_title()` replaces the title |
| 2 | Outline of the standard IPBES introduction. Its background (teal band and watermark) is copied onto every new slide |
| 3, 5, 6, 7 | Section dividers: big number, title and a photo (kingfisher, reindeer herder, coral reef, gorilla). `set_section()` replaces the number and title; the converter uses them in this order |
| 4 | What is IPBES? (name, mission, members, photo of the secretariat in Bonn) |
| 8 | Thank you / ¡Gracias! / Merci !, with the secretariat's address |

The numbers are in `TEMPLATE_SLIDES`, `SECTION_DIVIDERS` and `BACKGROUND_SOURCE` in `scripts/ipbes_pptx.py`. If the standard deck is updated, run `slide_text.py` (section 9) and check that they still hold. A larger version of the deck, with slides on the structure of IPBES, assessments, the timeline up to 2030 and the second global assessment, is kept by the secretariat outside this skill.

Only the slides used are kept in the output, so a deck is usually 5–11 MB, most of it the photos of the cover, dividers and thank-you slide.

## 4. Design of the standard slides

| Element | Style |
|---|---|
| Title | Bold, teal `#037473`, Helvetica LT Std. Template slides use 30 pt at x = 0.5", y ≈ 0.93". New slides use 24 pt in a box with room for two lines, bottom-aligned, so the last line sits at the same height on every slide. Use 21–22 pt only for a title that would otherwise need three lines. |
| Body text | Grey `#595959`, Helvetica LT Std, 12–14 pt (10–11 pt in dense slides). Left-aligned. |
| Bullets | Teal Wingdings square (§) at the first level; a dot below it. |
| Highlight | Orange: theme `accent2` `#ED7D31` (the brand orange is `#D05606`). Use for fills and large or bold text only: orange fails contrast for small text on white. |
| Cards, question boxes, total rows | Light teal tint `#E6F1F1`, rounded corners, no outline. |
| Tables | Teal header row with white bold text; alternate rows `#F2F2F2`/white; totals in teal bold on the light tint. Numbers right-aligned. |
| Charts | Native PowerPoint charts (not pictures), teal bars, one bar in orange to make a point, data labels on, no gridlines or value axis. |
| Background | The teal band along the top and the faint "ipbes" watermark are a **picture background on each slide**, not part of the layout. A slide added from a layout has neither; copy the background from a template content slide (slide 2). `Deck.new_slide()` does this. |
| Layout names | The layouts are in French ("Titre et contenu", "Titre seul"…). Template content slides use "Titre et contenu" with ordinary text boxes, not placeholders. |

## 5. Slide patterns

The converter applies these patterns from the outline markers. To build a slide by hand (section 8), use the same measurements. Content starts at y = 1.6" under the title (or 2.4" under a question box), 8.8" wide from x = 0.6".

| Content | Pattern |
|---|---|
| Title of the presentation | Template cover (slide 1) with `set_cover_title(title, subtitle)` |
| Thanks and contact | Template slide 41 |
| A few points of explanation | Bullets, 13–14 pt, left 5.6" wide; add a card or image on the right |
| One key figure (a rate, a total) | Points on the left; a card on the right with the figure in 54 pt teal and a short label |
| Three parallel items (roles, options) | Three cards 2.8" wide, 0.2" apart, each with a teal heading and bullets |
| Four to nine short items (categories) | Card grid: four cards per row, 0.13" apart; a long item gets a full-width card above the grid |
| A numbered list (agenda, steps) | Teal bold numbers with hanging indents and tab stops, 10 pt; sub-items (a), (b) one level in; a side card for related items |
| A table of figures | `table()`; totals as a bold row; footnote markers (¹ ² ³) with the notes in one 9 pt line below. Two small tables fit side by side at 4.3" each and 9 pt |
| Figures to compare (over time, by place) | `column_chart()`; highlight the bar the slide is about |
| An image with its own content (schedule, diagram) | Picture slide: the image alone under the title |
| Short labels or categories | Chips: `card(..., pill=True)` in teal with white bold text |
| A question with its answer | Question box (`question()`) under the title, the answer below; full question in the notes |
| A question without an answer | The question in a large card and a "For discussion" chip |

## 6. Language on slides

The language rules are those of any IPBES text (step 2 of section 1). The converter makes no language edits: correct the outline before converting it, and list each correction in the change table.

## 7. Converting the outline

```
python3 scripts/md_to_ipbes_pptx.py slides.md slides.pptx
```

- Image paths are relative to the outline file.
- Long titles (over about 70 characters) are set at 21 pt instead of 24 pt so that they fit in two lines.
- Body text shrinks from 14 pt to no less than 10 pt to fit; below that the script warns instead. Split the slide in the outline.
- A template slide can be used only once in a deck (python-pptx cannot copy slides), so a deck can have at most four `[section]` dividers.

Test: `python3 ipbes-editorial/scripts/md_to_ipbes_pptx.py test/ipbes-test-slides.md test-slides.pptx` (16 slides, no warnings).

## 8. Building slides by hand

For a layout the outline format does not cover, write a short script with the helpers in `scripts/ipbes_pptx.py` (needs `python-pptx`):

```python
from ipbes_pptx import Deck, money
deck = Deck()                                            # the standard deck in assets/
cover = deck.template_slide(1)
deck.set_cover_title(cover, 'Title of the presentation', 'Subtitle')
s = deck.new_slide('Title of the slide')                 # background + title
deck.bullets(s, 0.6, 1.6, 5.6, 3.6, ['First point', 'Second point'])
deck.card(s, 6.6, 1.7, 2.8, 2.5, 'Key figure', items=[...])
t = deck.new_slide('Costs by category')
deck.table(t, 0.6, 1.6, 8.8, rows, [5.4, 1.7, 1.7], 'Costs', bold_rows=(len(rows) - 1,), right_cols=(1, 2))
c = deck.new_slide('Participants by session')
deck.column_chart(c, 0.6, 1.6, 6.0, 3.6, ['2023', '2024', '2025'], [700, 800, 1000], 'Participants', highlight=2)
p = deck.new_slide('Organization of work')
deck.picture(p, 'schedule.png', alt='Organization of work over six days')
deck.notes(s, 'Speaker notes')
deck.finish([cover, s, t, c, p, deck.template_slide(41)], 'out.pptx')
```

Other helpers: `question()` for a tinted question box, `textbox()` for free text, and `add_paragraphs()` for mixed runs (bold labels, links) and bullet levels.

## 9. Checking the result

1. **Text:** `python3 slide_text.py deck.pptx` (from the Anthropic pptx skill) lists every slide's text, tables, chart data and notes. Check order, figures and typos against the source and the change table, and search for leftovers (`xxx`, `lorem`, `TODO`, placeholder headers).
2. **Figures:** recalculate every total and percentage in tables and charts. Flag inconsistencies to the author; do not correct them.
3. **File:** `python3 validate.py deck.pptx --original "assets/Standard Slides_IPBES Intro.pptx"` (also from that skill). Pass `--original` so the template's own schema errors are not reported as yours.
4. **Visual:** render every slide and look at each one for text that overflows or overlaps, uneven gaps and titles at different heights.
   - If LibreOffice is not installed and you cannot install it, use the portable AppImage, extracted without installing:

     ```
     curl -sSL -o lo.AppImage https://appimages.libreitalia.org/LibreOffice-fresh.basic-x86_64.AppImage
     chmod +x lo.AppImage && ./lo.AppImage --appimage-extract
     squashfs-root/opt/libreoffice*/program/soffice -env:UserInstallation=file:///tmp/lo/profile --headless --convert-to pdf deck.pptx
     pdftoppm -jpeg -r 110 deck.pdf slide
     ```

   - LibreOffice has no Helvetica LT Std and substitutes a **wider** font, so titles and labels look longer in the preview than in PowerPoint. If text only just fits in the preview, it fits in PowerPoint; if it overflows in the preview, shorten the layout anyway, since not every colleague has Helvetica LT Std (PowerPoint then uses Arial).
   - To check that a rebuilt deck matches an earlier one, render both and compare the images pixel by pixel.

## 10. Pitfalls already met

| Problem | Fix |
|---|---|
| New slides have no teal band or watermark | Copy the `<p:bg>` of slide 2 and relate its image to the new slide (`Deck.new_slide`). |
| Long titles wrap into the content | 24 pt, two-line title box bottom-aligned at y = 0.62–1.5"; content starts at 1.6". |
| Table rows with empty cells are too tall | An empty cell has no size, so PowerPoint uses 18 pt. Give every cell an `endParaRPr` with the table's size (`Deck.table`). |
| Text in shapes is centred | python-pptx shapes centre text by default; `add_paragraphs` left-aligns unless `align` is given. |
| Chips wrap onto two lines | Reduce side margins to 0.05" and allow about 1.7" width at 10.5 pt. |
| Python dict keyed by slide fails | python-pptx `Slide` objects are not hashable; use lists of pairs. |
| Searching a script for "³" fails | A file written with the character itself does not contain the escape `³`. Search for the character. |
| Gradient fills, 8-digit colours | Not supported by python-pptx; use a picture or a flat colour. |
