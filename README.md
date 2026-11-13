# ipbeseditorial

Editorial tools for English IPBES documents: the IPBES house style for AI assistants (Copilot and Claude), Word macros that check documents and apply corrections as tracked changes, and a Markdown-to-Word converter that produces documents in the IPBES style.

**Editorial principle:** the UN Editorial Manual is the baseline. IPBES departs from it only where its Plenary reports (IPBES-1 to IPBES-12) show a deliberate and consistent practice. The evidence is in [evidence.md](ipbes-editorial/references/evidence.md).

## Contents

| Path | Contents |
|---|---|
| [ipbes-editorial/SKILL.md](ipbes-editorial/SKILL.md) | Instructions Claude reads first |
| [ipbes-editorial/copilot/ipbes-editorial-instructions.txt](ipbes-editorial/copilot/ipbes-editorial-instructions.txt) | The same rules, condensed for Copilot |
| `ipbes-editorial/references/` | The full rules: [house-style.md](ipbes-editorial/references/house-style.md) (voice, countries, style, format), [ipbes-terms.md](ipbes-editorial/references/ipbes-terms.md) (bodies, terms, assessments, symbols), [report-structure.md](ipbes-editorial/references/report-structure.md) (masthead, Word styles, standard wording), [inclusive-language.md](ipbes-editorial/references/inclusive-language.md), `countries.csv`, `un-spelling-list.md` (October 2026) and [evidence.md](ipbes-editorial/references/evidence.md) |
| `ipbes-editorial/scripts/` | `md_to_ipbes_docx.py` (Markdown-to-Word converter) and `apply_changes.py` (applies a change table to a Word file as tracked changes) |
| `ipbes-editorial/assets/` | Empty Word template with the IPBES styles (`ipbes-document-template.docx`), and reference copies of the masthead logos (`masthead/`) |
| `word-macro/IPBESEditorial.bas` | Word macros (see section 3) |
| `test/` | Test files for all the tools: `ipbes-test-document.docx` (a document with typical errors), `ipbes-test-change-list.docx` (a change table for that document) and `ipbes-test-draft.md` (a Markdown draft with typical errors) |
| `tools/` | Maintenance scripts (see section 4) |
| `assets/` | IPBES logos, motifs and Visual Guidelines pages. These are brand and web material, kept here for now (see section 5). |

---

## 1. Using the assistant with Copilot

The assistant can **analyse** a document against the house style, **fix** it, and **convert** a Markdown draft into a Word document in the IPBES style. Copilot is the tool the organization approves; follow the organization's rules on which tool may be used for which documents.

### Option A: a Copilot agent (recommended, set up once)

1. Open **Microsoft 365 Copilot Chat** and choose **Create agent** (it may be called **New agent**).
2. On the **Configure** tab, fill in the fields:
   - **Name:** IPBES editorial assistant
   - **Description:** Edits and standardizes English IPBES documents in the IPBES and UN house style.
   - **Instructions:** paste the whole text of `ipbes-editorial/copilot/ipbes-editorial-instructions.txt`.
3. Under **Knowledge**, upload the following files by providing the github URL: https://github.com/ipbes/ipbeseditorial/tree/main/ipbes-editorial/references:
   - `house-style.md`
   - `ipbes-terms.md`
   - `report-structure.md`
   - `inclusive-language.md`
   - `countries.csv`

   If Copilot does not accept `.md` files, use Word copies (see section 4).
4. Select **Create**. Share the agent with colleagues if that is allowed.

### Option B: no agent

Start a new Copilot chat and paste the whole text of the instructions file. Then make your request, in the same message or the next one.

### What to ask

Attach the document, then ask one of three things:

1. **Analysis:** "Analyse the attached document against the IPBES editorial criteria."

   The assistant goes through the whole document and reports where it departs from the house style and formatting requirements, grouped by category. It does not change anything.
2. **Fixing:** "Fix the attached document to match the IPBES style and formatting requirements."

   Copilot returns a **change table** (No. | Original | Revised | Rule) and a list of queries for the author. It does not return a new file, because a file created by Copilot loses the Word styles and cannot show tracked changes. To apply the table:
   1. Copy the table from Copilot and paste it into a new Word document.
   2. Open the document to correct, and run the **IPBES_ApplyChangeList** macro (section 3).

   Each change is made as a tracked change, with the rule as a comment, and the formatting is kept. You then accept or reject each change.
3. **Conversion:** "Convert the attached Markdown draft into a Word document in the IPBES style."

   The assistant applies the language rules and returns a Word file. A file created by Copilot does not have the IPBES Word styles. For official documents, ask Copilot for the corrected Markdown instead, then convert it with the converter script (section 4), which uses the IPBES template.

**Always read the result before using it.** The assistant lists anything it is unsure of under "queries for the author". It must not change figures, names, dates, quotations or decisions. Report it if it ever does.

## 2. Using the assistant with Claude

- **Claude (web or desktop app):** package the skill as `ipbes-editorial.zip` (see section 4). Upload it under **Settings → Capabilities → Skills**. Claude then uses the skill whenever you ask it to edit an IPBES document.
- **Claude Code:** copy the `ipbes-editorial` folder into `~/.claude/skills/` (`cp -r ipbes-editorial ~/.claude/skills/`).

You ask for things the same way as in section 1. The difference is that Claude can run code on the file itself:

- **Fixing:** Claude writes the change table and applies it to your original Word file with `apply_changes.py`. You get your document back with tracked changes, comments giving the rules, and all formatting kept. You don't need to run the macro yourself.
- **Conversion:** Claude uses the converter script with the IPBES template, so the Word file has the IPBES styles.

The same safety rules apply. A change is made only if its original text is found exactly once, and anything else is reported to you.

---

## 3. Word macros

The macros check a document against the house style and mark possible issues as Word comments. They work in Word itself, without AI and without an internet connection.

| Macro | What it does | Changes the text? |
|---|---|---|
| **IPBES_CheckDocument** | Adds a comment starting "[IPBES]" at each possible issue and opens a summary report. Checks the Platform's name, internal codes (BBA, NXS…), abbreviations, numbers and dates, punctuation, spelling, terms, country names, inclusive language, the masthead (logos, symbol, date, spacing) and heading numbering. | No |
| **IPBES_SafeFixes** | Makes safe corrections **with Track Changes on**: the shortened Platform name, "e.g." → "e.g.,", -ise → -ize, em dashes and spaced hyphens → en dashes, double spaces, straight → curly quotes, country names such as Türkiye and Czechia, "Indigenous Peoples", "spatial planning and connectivity assessment", and a few spellings. | Yes, tracked |
| **IPBES_CheckVoice** | Marks prescriptive wording ("must", "should", "need to", "unprecedented"…) for review. IPBES is policy-relevant, not policy-prescriptive. | No |
| **IPBES_ApplyChangeList** | Applies a change table from Copilot or Claude (columns **Original** and **Revised**, optionally **Rule**) to the open document, as tracked changes, keeping all formatting. A change is made only if its original text is found exactly once. Text found more than once gets a comment instead, and the report lists anything not found. | Yes, tracked |
| **IPBES_RemoveIPBESComments** | Deletes only the comments the macros added. Other people's comments are kept. | No |

- **Agreed text** (decisions, approved summaries for policymakers, quotations) must not be changed. Reject any tracked change there, and treat comments there as queries.
- The comments are suggestions, not errors. For example, "should" in reported speech, or "Secretariat" in a budget table, is correct.
- To keep the document readable, the macros mark at most 25 places per rule. The report shows the full count.

### Installing

**For one computer:**

1. In Word, press **Alt+F11** (on a Mac: **Tools → Macro → Visual Basic Editor**).
2. In the left pane, select **Normal**.
3. Choose **File → Import File…** and select `word-macro/IPBESEditorial.bas`.
4. Choose **Debug → Compile Normal**. If nothing happens, the macros compiled correctly.

**To share with the team as an add-in (recommended):**

1. Do the steps above in a new blank document, but select that document's project instead of Normal.
2. Save the document as a **Word Macro-Enabled Template (*.dotm)** named `IPBES Editorial.dotm`.
3. Each colleague copies the file into their Word **Startup** folder and restarts Word:
   - **Windows:** `%APPDATA%\Microsoft\Word\STARTUP` (paste this into the File Explorer address bar).
   - **Mac:** `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word`.

**Buttons:**

1. Go to **File → Options → Quick Access Toolbar**.
2. Under "Choose commands from", select **Macros**.
3. Add **IPBES_CheckDocument**, **IPBES_SafeFixes**, **IPBES_ApplyChangeList** and **IPBES_RemoveIPBESComments**, then use **Modify…** to give them icons and short names.

**Macro security:** if the macros are blocked, check **File → Options → Trust Center → Trusted Locations**; the Startup folder is normally trusted. If the organization blocks macros, ask IT to sign the `.dotm` or to trust its location. Do not lower macro security for all documents.

### Using the macros

1. Save a copy of the document.
2. Run **IPBES_CheckDocument**, either with the toolbar button or through **View → Macros**.
3. Read the report that opens, then step through the comments with **Review → Next comment**.
4. Optional: run **IPBES_SafeFixes**, then review the tracked changes and accept or reject each one.
5. To apply a change table from Copilot:
   1. Paste the table into a new Word document and keep it open.
   2. Switch back to the document to correct and run **IPBES_ApplyChangeList**.
   3. Read the report and review the tracked changes.
6. When you have finished, run **IPBES_RemoveIPBESComments**.

### First-time test (before sharing)

The macros were written without access to Word, so test them on a copy of a recent Plenary report and on the files in `test/`:

- [ ] The module compiles (**Debug → Compile**).
- [ ] On the IPBES-12 report (IPBES/12/11), there are few or no masthead comments.
- [ ] Only straight quotes are marked, not curly ones.
- [ ] **IPBES_SafeFixes** on `test/ipbes-test-document.docx` produces tracked changes only.
- [ ] **IPBES_ApplyChangeList**: open `test/ipbes-test-document.docx` and `test/ipbes-test-change-list.docx`, then run the macro from the test document. Items 1–6 should be applied as tracked changes. Item 7 should be reported as not found, item 8 as skipped, and item 9 as found twice (with a comment, not applied).
- [ ] **IPBES_RemoveIPBESComments** leaves other people's comments in place.

### Maintaining the rules

The rules are in the `BuildRules` procedure, one `AddRule` line each. A line gives the category, the text to find, wildcards, match case, whole word, the message, and optionally the replacement and whether it is a safe fix. Use `ChrW()` for non-ASCII characters, because the VBA editor does not import them reliably.

---

## 4. For the secretariat editor

Run these commands from the repository folder. The scripts need Python 3; the Word tools also need python-docx (`pip install --user python-docx`).

| Task | Command |
|---|---|
| Package the skill for Claude | `python3 tools/package_skill.py` (creates `ipbes-editorial.zip`, with the `ipbes-editorial` folder that Claude expects) |
| Convert a Markdown draft to Word | `python3 ipbes-editorial/scripts/md_to_ipbes_docx.py draft.md draft.docx --numbered --symbol IPBES/13/5` (leave out `--numbered` for documents without numbered paragraphs). Test: `python3 ipbes-editorial/scripts/md_to_ipbes_docx.py test/ipbes-test-draft.md test-draft.docx --numbered --number-headings` |
| Apply a change table to a Word file as tracked changes | `python3 ipbes-editorial/scripts/apply_changes.py document.docx changes.md corrected.docx`. The change table can be Markdown, CSV, JSON or a Word table. Test: `python3 ipbes-editorial/scripts/apply_changes.py test/ipbes-test-document.docx test/ipbes-test-change-list.docx test-corrected.docx` |
| Update the country list | Download https://ibok.ipbes.net/rest/v1/geo/data?format=application/rdf%2Bxml as `ipbes-geo.rdf`, then run `python3 tools/build_countries.py ipbes-geo.rdf ipbes-editorial/references/countries.csv` |
| Rebuild the Word template from a new Plenary report | `python3 tools/make_template.py report.docx ipbes-editorial/assets/ipbes-document-template.docx` |
| Add a new Plenary report to the evidence | `python3 tools/docx_dump.py report.docx temp/dump/2027-13.txt`, then `python3 tools/evidence_scan.py temp/dump > temp/evidence_counts.md`. Compare the result with `evidence.md` and update the rules if practice has changed. |
| Extract text from a PDF source | `python3 tools/pdf_text.py source.pdf output.txt` (or `pdftotext -layout source.pdf output.txt`) |
| Make Word copies of the reference files | Open each `.md` file in Word (File → Open, "All files") and save it as `.docx`, or run `pandoc house-style.md -o house-style.docx` |

**After any change to the rules:**

- update `ipbes-editorial/copilot/ipbes-editorial-instructions.txt` (Copilot's limit is 8,000 characters, and the file is close to it);
- re-upload the knowledge files to the Copilot agent;
- re-package the Claude skill;
- update the matching `AddRule` lines in `word-macro/IPBESEditorial.bas`.

---

## 5. Brand assets (`assets/`)

These assets were extracted from the *IPBES Logo and Visual Identity Guidelines* (2021) for web and design work, which will move to a separate repository. Official documents use only the masthead logos in `ipbes-editorial/assets/masthead/` and Times New Roman.

- **`logo/`:** the IPBES logo in SVG (preferred), PDF and PNG, with and without the strapline "Science and Policy for People and Nature":
  - **original** and **medium teal**, for light backgrounds;
  - **mono**, for single colour on light backgrounds;
  - **reversed**, for dark full-colour backgrounds;
  - **whiteout**, for dark single-colour backgrounds;
  - **clear-space**, showing the minimum clear space (the height of the ascender of the "b").

  Never distort, redraw or generate the logo.
- **`motifs/`:** the fingerprint, ecosystems, species and genetics motifs.
- **`reference-pages/`:** images of pages 01–10 of the guidelines.
- **Typography:** Helvetica Neue. If it is not available, use another version of Helvetica, or Arial.
- **Primary colours:** Teal #037473, Orange (Pumpkin) #D05606, Grey #676C5F, Light Green #BFD787. The secondary palette is on `reference-pages/05-colour-palette.png`.
- **Contrast for text on white** (WCAG, at least 4.5:1):
  - Fine for any text: Teal, Grey, True Red, Plum, Purple, Faded Purple, Navy Blue, Denim Blue, Lagoon Blue, Dark Grey, Forest Green.
  - Large text only: Orange, Medium Teal, Spring Green, Mole Grey, Foam Green.
  - Never for text: Light Green, Yellow, the secondary Orange, Ice Grey, Medium Green, Light Grey.

---

**Sources:**

- UN Editorial Manual (https://www.un.org/dgacm/en/content/editorial-manual) and UN spelling list (October 2026)
- UN *Guidelines for gender-inclusive language in English*
- UN *Disability-Inclusive Communications Guidelines* (2022)
- IPBES geography list
- IPBES assessment abbreviations list
- IPBES Plenary reports IPBES/1/12 to IPBES/12/11
- IPBES *Logo and Visual Identity Guidelines* (2021)
