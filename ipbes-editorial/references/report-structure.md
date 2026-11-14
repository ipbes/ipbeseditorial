# IPBES document structure, masthead and Word template

Sources:

- The Word files of the Plenary reports. The layout described here has been stable since IPBES-10 (2023) and is identical in IPBES-12 (2026).
- The UN Editorial Manual pages on mastheads and cover pages (https://www.un.org/dgacm/en/content/editorial-manual/mast-cover), headings, paragraphs and tables.

---

## 1. Masthead

**Check the masthead only if the document has one.** Information documents, web text and drafts often have none. Do not add a masthead unless asked.

### 1.1 Layout

The current layout (since IPBES-10) is four borderless tables separated by tiny `A_Spacer` paragraphs (1-point text):

| Block | Left | Right |
|---|---|---|
| 1 | "UNITED NATIONS" (two lines) | UN and partner logo strip |
| — | `A_Spacer` | |
| 2 | (empty) | Document symbol, for example **IPBES/12/11** |
| — | `A_Spacer` | |
| 3 | IPBES logo, then "Intergovernmental Science-Policy / Platform on Biodiversity and / Ecosystem Services" (three lines) | "Distr.: General" / date of issue / "Original: English" |
| — | `A_Spacer` | |
| 4 | "Plenary of the Intergovernmental Science-Policy / Platform on Biodiversity and Ecosystem Services" / "Twelfth session" / "Manchester, United Kingdom of Great Britain and Northern Ireland, 3–8 February 2026" | (empty) |

The document title follows in style `BB_Title`:

- For the session report: "Report of the Plenary of the Intergovernmental Science‑Policy Platform on Biodiversity and Ecosystem Services on the work of its twelfth session".
- For a pre-session document, the title is preceded by "Item X of the provisional agenda*". The asterisk points to a footnote giving the provisional agenda symbol (UN Manual).

### 1.2 Logos

These are the logos found in the IPBES-8 to IPBES-12 reports. Reference copies are in this skill, in `assets/masthead/`.

| Logo | File | Size in the masthead | Check |
|---|---|---|---|
| UN and partner strip | `assets/masthead/un-partner-logos-strip.png` | about 11.1 × 1.1 cm (ratio about 10:1) | Black and white. Order from left: UN emblem, UNEP, UNESCO, FAO ("Food and Agriculture Organization of the United Nations"), UNDP. All five must be present. |
| IPBES logo | `assets/masthead/ipbes-logo-masthead-mono.png` | about 3.1 × 1.4 cm (ratio about 2.2:1) | **Black (mono), without strapline.** Not the colour logo, not the strapline version and not reversed. |

**Logo errors to flag:**

- a logo is missing (an empty cell where block 1 or block 3 should have an image);
- a partner logo is missing or the partners are out of order;
- a colour IPBES logo, or one with the strapline "Science and Policy for People and Nature", is used in the masthead (masthead documents are printed in black and white);
- the logo is stretched (the ratio differs from the reference by more than about 5 per cent), cropped or recoloured;
- an old logo, such as the 2013–2019 multi-image strip or a "BES" text logo, appears in a new document.

The full-colour logos and their variants, for covers, web and other material, are in the repository folder `assets/logo/` (not part of this skill). Their use is set out in the *IPBES Visual Guidelines* (2021), repository `assets/reference-pages/03-logo-page.png`:

- five colour versions: original, medium teal, mono, reversed and whiteout;
- clear space equal to the height of the ascender of the "b";
- never distort or dismantle the logo.

### 1.3 Text and spacing checks

- **Symbol.** It is the same in the masthead, the page headers and any references in the text: "IPBES/12/11", not "IPBES-12/11" or "IPBES/12/11/".
- **Date of issue.** Write "12 February 2026" (day month year, no comma). It is on the line after "Distr.: General", followed by "Original: English".
- **Session line.**
  - The ordinal is in words with an initial capital: "Twelfth session".
  - The venue is the city plus the country in its UNTERM short form. When the session is held at the seat (Bonn), the reports usually still add the country ("Bonn, Germany").
  - Write the dates as "3–8 February 2026". Use "28 August–2 September 2023" across months.
- **Platform name.** It appears twice in the masthead (block 3 and block 4) and must be the full name both times.
- **Spacing.**
  - Exactly one `A_Spacer` paragraph sits between the tables.
  - There are no empty `Normal` paragraphs between the blocks or before the title.
  - There are no double spaces in the masthead text.
  - Line breaks inside the name follow the pattern above.
- **Page headers and footers.**
  - Header (`Header-pool`, 9-point bold): the symbol, on even and odd pages.
  - Footer (`Footer-pool`): the page number.
  - The first page has a separate footer (`Footer-jobnumber`) with the job number and date, for example "K2601747[E]	170326". The job number comes from conference services; never invent one. If it is missing in a final document, flag it.
  - "Different first page" is on.

## 2. Page set-up

| Setting | Value |
|---|---|
| Paper | A4 (21 × 29.7 cm), portrait. Landscape sections are allowed for wide tables. |
| Margins | top 1.6 cm, bottom 2.5 cm, left 2.5 cm, right 1.75 cm |
| Header and footer distance | 0.95 cm and 1.72 cm |
| Base font | Times New Roman, 10 point |
| Language | English (United Kingdom) |

## 3. Paragraph styles

These styles are used in all reports from IPBES-2 to IPBES-12. Apply styles; do not format manually.

| Style | Use | Format |
|---|---|---|
| `BB_Title` | Document title | 14 pt bold, indented 2.2 cm |
| `CH1` | Chapter: "I.", "II." | 14 pt bold, hanging indent 2.2 cm, keep with next |
| `CH2` | Section: "A.", "B." | 12 pt bold, hanging indent 2.2 cm |
| `CH3` | Subsection: "1.", "2." | 10 pt bold, hanging indent 2.2 cm |
| `CH4` | Lower subsection: "(a)" | 10 pt bold, hanging indent 2.2 cm |
| `Normal_number` | Numbered paragraphs. Automatic numbering: 1., then (a), then (i), then a., then i. | 10 pt, 6 pt after |
| `Normal_No_number` | Unnumbered text (chapeaus of decisions, agenda lists, quotations) | 10 pt, indent 2.2 cm, 6 pt after |
| `Normal-pool` | Base style for lists and annex text | 10 pt |
| `Title_table` | "Table 1", line break, title in bold; the headnote "(United States dollars)" as a separate `Title_table` paragraph | 10 pt bold, indent 2.2 cm, keep with next |
| `Title_figure` | "Figure I", line break, title | as `Title_table` |
| `ZZ_Anx_header` | "Annex" | 14 pt bold |
| `ZZ_Anx_title` | Title of the annex | 14 pt bold, indent 2.2 cm |
| `AA_Title` | Session block (older mastheads) | bold |
| `A_Spacer` | Spacer between masthead tables | 1 pt |
| `Header-pool`, `Footer-pool`, `Footer-jobnumber` | Headers and footers | 9 pt bold (header and page number) |

Notes:

- The heading number is followed by a tab, and the hanging indent holds the number in the margin column ("\tI.\tOpening of the session").
- In IPBES-12 the annex heading block ("Annex" and "Decisions adopted by the Plenary…") was typed in `Normal`, not `ZZ_Anx_header` and `ZZ_Anx_title`. Accept either, but suggest the proper styles.
- Decision titles are `CH1` without a number: "Decision IPBES-12/1: Implementation of the rolling work programme…".
- The parts inside a decision are `CH2` with a Roman numeral on its own line and the title on the next line.

## 4. Structure of a session report

The CH1 headings follow the agenda. Since IPBES-10 they have been:

1. I. Opening of the session
2. II. Organizational matters
   - A. Adoption of the agenda and organization of work
   - B. Status of the membership of the Platform
   - C. Election of officers (when there are elections)
3. III. Admission of observers
4. IV. Credentials of representatives
5. V. Report of the Executive Secretary on progress in the implementation of the rolling work programme of the Platform up to 2030
6. VI. Financial and budgetary arrangements for the Platform
7. VII. Assessing knowledge (followed by the assessment considered)
8. VIII. Building capacity, strengthening knowledge foundations and supporting policy
9. IX. Improving the effectiveness of the Platform
10. X. Requests, input and suggestions for additional elements of the rolling work programme of the Platform up to 2030
11. XI. Organization of the Plenary; dates and venues of future sessions of the Plenary
12. XII. Institutional arrangements: United Nations collaborative partnership arrangement for the work of the Platform and its secretariat
13. XIII. Adoption of the decisions and the report of the session
14. XIV. Closure of the session

These are followed by the Annex, "Decisions adopted by the Plenary of the Intergovernmental Science-Policy Platform on Biodiversity and Ecosystem Services at its twelfth session". It contains a list of the decisions, then each decision with its own annexes ("Annex I to decision IPBES-12/1").

**Checks:**

- The headings match the adopted agenda (section II.A), in the same order and with the same wording.
- The Roman numerals have no gaps.
- Each decision mentioned in the text ("The decision is set out in the annex to the present report") exists in the annex.

## 5. Standard formulas

These formulas recur in every recent report. They are deliberate, so keep them; do not "improve" them. When drafting, use them.

| Situation | Formula |
|---|---|
| Opening | "The twelfth session of the Plenary of the Intergovernmental Science-Policy Platform on Biodiversity and Ecosystem Services (IPBES) was held at [venue], in [city], [country], from 3 to 8 February 2026." |
| | "Following a [musical performance / video presentation], the session was opened at 9.55 a.m. by the Chair of IPBES, [Full Name], who welcomed participants." |
| | "Opening statements were delivered by [Full Name], [title]; [Full Name], [title]; and …" |
| Group statements | "The representatives of Ghana on behalf of the Group of African States; Indonesia on behalf of the Group of Asia-Pacific States; …" or "the representative of X, speaking on behalf of …" |
| Introducing an item | "Introducing the item, the Chair drew attention to the note by the secretariat on … (IPBES/12/6)…". For a sub-item: "Introducing the sub-item, …" |
| Agenda | "The Plenary adopted the following agenda on the basis of the provisional agenda (IPBES/12/1):" followed by the agenda as a numbered list in `Normal_No_number`. Items end with a full stop; sub-items end with ";" and "." on the last one. |
| | "The Plenary agreed to follow the organization of work set out in annex I to document IPBES/12/1/Add.1." |
| Working arrangements | "The Plenary agreed to establish two working groups." / "The Plenary agreed to request working group II to consider the item, working on the basis of the Chair’s informal note." |
| | "The Plenary established a contact group on budget, [co-]chaired by [Name] ([Country]), to further consider the financial and budgetary arrangements for the Platform." |
| Information items | "The Plenary took note of the information provided and welcomed the work carried out." |
| Decisions | "Subsequently, the Plenary considered a draft decision on the matter (IPBES/12/L.3)." / "Subsequently, the Plenary adopted decision IPBES-12/1 on …, as orally amended, including part II thereof, entitled “Assessing knowledge”." / "The decision is set out in the annex to the present report." |
| Credentials | "The Bureau, with the assistance of the secretariat, examined the credentials of the representatives of IPBES members submitted in accordance with rule 13 of the rules of procedure." … "The Plenary approved the report of the Bureau on credentials." |
| Observers | "…the interim procedure for the admission of observers to sessions of the Plenary, as described in paragraph 22 of the report of the first session of the Plenary (IPBES/1/12)…" |
| Adoption | "The Plenary adopted decisions IPBES-12/1–IPBES-12/3, as set out in the annex to the present report." / "The Plenary adopted the present report on the basis of the draft report that had been circulated, on the understanding that the report would be finalized by the secretariat under the supervision of the Bureau." |
| Closure | "Following the customary exchange of courtesies, the Chair declared the session closed at 2.20 p.m. on 8 February 2026." |

## 6. Decisions

- Title: "Decision IPBES-12/1: [title]" (`CH1`, no number).
- Opening: "The Plenary," on its own line (`Normal_No_number`, roman type).
- **Preambular paragraphs** begin with an italic participle ("*Recalling*", "*Welcoming*", "*Acknowledging with appreciation*") and end with a comma. They are not numbered.
- **Operative paragraphs** are numbered (`Normal_number`). They begin with an italic verb in the third person ("*Decides*", "*Requests* the Executive Secretary", "*Also requests*", "*Further invites*", "*Approves*", "*Welcomes*", "*Takes note of*") and end with a semicolon. The last one ends with a full stop.
- Decisions are edited like the rest of the document: correct spelling, punctuation, terms and format to house style. Do not change what a decision decides, requests or approves.

## 7. Converting Markdown to an IPBES-style Word document

When converting a draft written in Markdown, map the elements to the styles above:

| Markdown | Word style |
|---|---|
| `# Title` | `BB_Title` |
| `## I. Heading` | `CH1`. Roman numeral, tab, title. |
| `### A. Heading` | `CH2` |
| `#### 1. Heading` | `CH3` |
| Plain paragraph | `Normal_number` in reports and documents with numbered paragraphs; otherwise `Normal_No_number` |
| `1.` and `(a)` lists | `Normal_number` at level 2 ((a)) and level 3 ((i)). Do not use bullets in reports. |
| `- bullet` | `Normal_No_number` with a bullet, for short items only, outside reports |
| Table plus a caption line "Table 1: Title" | `Title_table` ("Table 1", line break, title), then the headnote, then the table. Horizontal rules only (see `house-style.md` 5.5). |
| `**bold**` and `*italic*` | Keep only where house style allows: bold for headings and run-in headings; italics for publication titles, scientific names and decision verbs. |
| `[text](url)` | Hyperlinked text, with the URL in full in a footnote for masthead documents |
| `> quote` | `Normal_No_number`, indented, without quotation marks (if longer than five lines) |
| Footnote `[^1]` | Word footnote. Footnote numbers go after punctuation. |

Then apply the language checks in `house-style.md`.

The converter `scripts/md_to_ipbes_docx.py` uses the clean IPBES template in `assets/ipbes-document-template.docx` (styles, numbering and page set-up of IPBES/12/11). Without it, a script (for example python-docx) can create the styles with the formats in section 3, Times New Roman 10 pt and the page set-up in section 2. A masthead is added **only on request**, using the logos in `assets/masthead/`.
