# IPBES official documents editing and formatting tool

Editorial tools for English IPBES documents: the IPBES house style for MS Word macros and AI assistants (Copilot and Claude), that check documents and apply corrections as tracked changes, and a Markdown-to-Word converter that produces documents in the IPBES style.

**Editorial principle:** the UN Editorial Manual is the baseline. IPBES departs from it only where its Plenary reports (IPBES-1 to IPBES-12) show a deliberate and consistent practice. The evidence is in [evidence.md](ipbes-editorial/references/evidence.md).

## Quick start: install the IPBES tab in Word (Windows)

1. **Close Word** completely, including any open documents.
2. **Download the add-in:** [IPBES Editorial.dotm](https://github.com/ipbes/ipbeseditorial/raw/main/ipbes-editorial/IPBES%20Editorial.dotm). It goes to your **Downloads** folder.
3. **Open your Word Startup folder:** press **Windows key + R**, paste the line below, and press **Enter**. Then move `IPBES Editorial.dotm` from **Downloads** into the folder that opens.

   ```
   %APPDATA%\Microsoft\Word\STARTUP
   ```

   If Windows says the folder cannot be found, paste `%APPDATA%\Microsoft\Word` instead, create a new folder there called `STARTUP`, and open it.

   **Mac:** the Startup folder is `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word`. Skip step 4.

4. **Unblock the file:** Windows blocks macros in files downloaded from the internet. Press the **Windows key**, type **PowerShell**, open it, paste this command, and press **Enter**:

   ```powershell
   Unblock-File "$env:APPDATA\Microsoft\Word\STARTUP\IPBES Editorial.dotm"
   ```

   No message means it worked.
5. **Start Word.** An **IPBES** tab appears after the **Review** tab. Section 1 below explains what each button does.

**If the tab doesn't appear or the macros are blocked:** go to **File → Options → Add-ins**, set **Manage** to **Templates**, select **Go…**, and check that `IPBES Editorial.dotm` is listed and ticked. Also check **File → Options → Trust Center → Trusted Locations**; the Startup folder is normally trusted. If the organization blocks macros, ask IT to sign the `.dotm` or to trust its location. Do not lower macro security for all documents.

## 1. Microsoft Word VBA macros (No AI)

The macros check a document against the house style and mark possible issues as Word comments. They work in Word itself, without AI and without an internet connection.

**Typical order** (save a copy of the document first):

1. **Fix formatting**: answer whether the document has numbered paragraphs, then review the tracked changes.
2. **Check document**: read the report, then step through the comments with **Review → Next comment**.
3. **Safe fixes** (optional): accept or reject each tracked change.
4. **Remove IPBES comments** when you have finished.

### IPBES_CheckDocument
Adds a comment starting "[IPBES]" at each possible issue and opens a summary report. Checks the Platform's name, internal codes (BBA, NXS…), abbreviations, numbers and dates, punctuation, spelling, terms, country names, inclusive language, the masthead (logos, symbol, date, spacing) and heading numbering.

### IPBES_FixFormatting
Applies the IPBES layout, using the IPBES styles carried in the add-in. It asks whether the document is a report with numbered paragraphs. It then: copies the IPBES Word styles into the document; sets A4 paper and the IPBES margins; turns Title, Heading 1–4 and Caption into BB_Title, CH1–CH4 and Title_table, laying out heading numbers as tab, "I.", tab; puts body text in Normal_number (numbered) or Normal_No_number, removing typed paragraph numbers and manual fonts, sizes, colours and indents (bold and italic are kept); deletes empty spacer paragraphs (the IPBES styles provide the spacing); reduces double (or longer) spaces to one; turns "Table 1: Title" captions and "(Number)" headnotes into the IPBES table title; and formats tables (indented to the text, horizontal rules only, no shading, italic column heads, numbers right-aligned). Documents that already use the IPBES styles (such as recent Plenary reports) keep their own style definitions, headers, footers, table layout and Normal paragraphs; only other paragraphs, captions and over-wide tables are changed, so the run is quick. It ends with a short summary; the tracked changes are the detailed record, and Ctrl+Z undoes the whole run. Lists, empty paragraphs and unrecognised styles are listed for you to check. Style and text changes are tracked; page set-up and style definitions are not.

### IPBES_SafeFixes
Applies **every rule in the guidance that has a standard correction**, with Track Changes on. Where a rule has exceptions, the macro applies the usual form and you reject the exceptions when you review:

- **Platform and codes:** the full Platform name, "IPBES's" → "IPBES", internal codes (BBA, SPM, NCP, ILK, IPLCs…) → short names.
- **Titles and abbreviations:** "H.E." and "His/Her Excellency" removed, "Dr." removed, "Mrs." → "Ms.", "Mr"/"Ms" → "Mr."/"Ms.", "e.g." → "e.g.,", "U.N." → "United Nations", "US$"/"USD" → "$", "US dollars" → "United States dollars".
- **Numbers and dates:** "15%" → "15 per cent"; "3rd March", "March 3, 2026", "3-8 March", "from 3 March to 5 March" → "3 March", "3 March 2026", "from 3 to 8 March", "from 3 to 5 March"; "from 2024–2026" → "from 2024 to 2026"; "between 10-15" → "between 10 and 15"; en dashes in year, page and per cent ranges; "10.00 a.m.", "10:15 a.m.", "10am" → "10 a.m.", "10.15 a.m."; "12th session" → "twelfth session"; "21st century" → "twenty-first century"; "6(a)" → "6 (a)"; "this/next/last year" → the year, taken from today's date (check older documents).
- **Punctuation:** em dashes and spaced hyphens → spaced en dashes, double spaces, no space before ; : ? !, straight → curly quotes.
- **Spelling:** -ise → -ize, UN compounds (policymakers, decision makers, capacity-building, well-being, data set, website…), British spellings (programme, judgement, centre, colour, behaviour, analyse, ageing, towards, adviser…), "Internet".
- **Terms:** lower case for summary for policymakers, work programme, task force, technical support unit, trust fund, secretariat, scoping report, conceptual framework, contributions to people; "Indigenous Peoples", "Indigenous and local knowledge", "Kunming-Montreal", "spatial planning and connectivity assessment".
- **Countries:** Türkiye, Czechia, Viet Nam, Eswatini, Cabo Verde, Côte d’Ivoire, Republic of Korea, Democratic People’s Republic of Korea, Russian Federation, Syrian Arab Republic, Lao People’s Democratic Republic, Netherlands (Kingdom of the), North Macedonia, "UK"/"USA"/"U.S." → "United Kingdom"/"United States" (give the full name at first mention yourself).
- **Inclusive language:** Chairman/Chairwoman/Chairperson → Chair, "he/she" → "he or she" (rewording is better), mankind, manpower, man-made, spokesman, fishermen, layman, policeman, and the disability terms ("the disabled", "suffers from", "wheelchair-bound", "the mentally ill", "the blind"…) → the UN terms.

What it cannot fix by find-and-replace stays for **Check document**: prescriptive voice, "and/or", "the plenary", "Member States", "Co-chair", generic "he", first mentions of names and abbreviations, and anything that needs the sentence rewritten.

### IPBES_CheckVoice
Marks prescriptive wording ("must", "should", "need to", "unprecedented"…) for review. IPBES is policy-relevant, not policy-prescriptive.

### IPBES_ApplyChangeList
Applies a change table from Copilot or Claude (columns **Original** and **Revised**, optionally **Rule**) to the open document, as tracked changes, keeping all formatting. A change is made only if its original text is found exactly once. Text found more than once gets a comment instead, and the closing message lists the items not applied.

To apply a **change table** (see Copilot section below):
   1. Paste the table into a new Word document and keep it open.
   2. Switch back to the document to correct and run **Apply change list** (**IPBES_ApplyChangeList**).
   3. Read the closing message and review the tracked changes.

### IPBES_RemoveIPBESComments
Deletes only the comments the macros added. Other people's comments are kept.

- **Every document is open for editing**, including decisions and annexes: the secretariat edits a document and sends it clean to UNDOCS for final editing and publishing. Nothing is edited after UNDOCS has processed it, so the tools treat every document as a draft for submission. What must not change is the substance (what was decided, figures, names, dates) and the words of direct quotations. Reject any change that alters them.
- The comments are suggestions, not errors. For example, "should" in reported speech, or "Secretariat" in a budget table, is correct.
- To keep the document readable, the macros mark at most 25 places per rule. The report shows the full count.

### Updating the add-in after a new version of the macros

1. Close Word completely.
2. Start Word and open `IPBES Editorial.dotm` from the STARTUP folder (%appdata%\Microsoft\Word\Startup) with **File → Open**. A double-click would create a new document instead.
3. Press **Alt+F11**. Under the template's project, right-click **Modules → IPBESEditorial**, choose **Remove IPBESEditorial…**, and answer **No** to exporting it.
4. Choose **File → Import File…**, select the new `word-macro/IPBESEditorial.bas`, then choose **Debug → Compile**.
5. Save, close Word, and start it again.
6. In a WSL terminal in the repository folder, run:

   ```
   python3 tools/prepare_addin.py "$(wslpath "$(cmd.exe /c 'echo %APPDATA%' | tr -d '\r')")/Microsoft/Word/STARTUP/IPBES Editorial.dotm"
   ```

   The script copies the IPBES styles and numbering from `ipbes-editorial/assets/ipbes-document-template.docx` into the add-in and adds the ribbon tab from `word-macro/ipbes-ribbon.xml`. It keeps a backup (`IPBES Editorial.dotm.bak`). Run it again after any change to either file. Importing a new `.bas` into the `.dotm` doesn't remove the tab or the styles.


### Maintaining the rules

The rules are in the `BuildRules` procedure, one `AddRule` line each. A line gives the category, the text to find, wildcards, match case, whole word, the message, and optionally the replacement and whether it is a safe fix. Use `ChrW()` for non-ASCII characters, because the VBA editor does not import them reliably.

---

## 2. Using the assistant with Copilot

The assistant can **analyse** a document against the house style, **edit** it, and **convert** a Markdown draft into a Word document in the IPBES style. Copilot **cannot produce a document with tracked changes**.

### Opening the shared agent

The secretariat has a shared agent, so you don't need to set anything up:

1. Open **Microsoft 365 Copilot Chat**.
2. Search the agents for **IPBES editorial assistant** and open it.
3. Attach your document and paste one of the prompts below.

### Setting up the agent (secretariat editor only)

1. Open **Microsoft 365 Copilot Chat** and choose **Create agent** (it may be called **New agent**).
2. On the **Configure** tab, fill in the fields:
   - **Name:** IPBES editorial assistant
   - **Description:** Edits and standardizes English IPBES documents in the IPBES and UN house style.
   - **Instructions:** paste the whole text of `ipbes-editorial/copilot/ipbes-editorial-instructions.txt`.
3. Under **Skills**, select **Add** and upload `ipbes-editorial.zip`. It holds the full rules (house style, terms, country list, spelling list). To make it, run `python3 tools/package_skill.py` in the repository folder. Leave **Knowledge** empty.
4. Select **Create**, then share the agent with the secretariat.

After any change to the rules, update the instructions and upload the new zip to the same agent (see section 4), so that colleagues keep using one agent.

### What to ask

Pick the task, attach the document, and paste the prompt. Start a **new chat** for each document.

| I want… | Use | Prompt to paste |
|---|---|---|
| **A list of problems**, nothing changed | Copilot | `Analyse the attached document against the IPBES editorial rules. Do not change anything.` |
| **A Markdown draft turned into an IPBES Word document** | Copilot | `Convert the attached Markdown into an IPBES Word document with the skill's script scripts/md_to_ipbes_docx.py, using --numbered --number-headings. Return the Word file and the queries.` |
| **A clean, edited Word document** in the IPBES styles | Copilot | `Edit the attached document to the IPBES rules. Return a clean Word file made with the skill's script scripts/md_to_ipbes_docx.py, using --numbered --number-headings, plus the change table and the queries.` |
| **Tracked changes in my own document** | Copilot, then Word | `Give me a change list only for the attached document: one table with No. / Original / Revised / Rule, and the queries.` Paste the table into a new Word document, then press **Apply change list** in your document. |
| **Formatting only** (styles, numbering, tables, margins) | Word | No prompt: press **Fix formatting** on the IPBES tab. |

Leave out `--numbered` for documents without numbered paragraphs, such as notes and web text.

**After a Copilot Word file comes back:**

- **Check the styles.** Click a heading: the style box on the Home tab should say **CH1** or **CH2**, not **Heading 1**. If it says Heading 1, Copilot didn't use the script: press **Fix formatting** on the IPBES tab, then **Review → Accept → Accept All Changes**.
- **See what changed:** **Review → Compare → Compare…**, with your document as the original and Copilot's file as the revised one. Under **More**, untick **Formatting** to see only the wording changes.
- **Copy across what Copilot can't carry over:** the masthead, images and page headers.

**Use the Word macros instead of Copilot** for long reports and for scanned (OCR) documents. Copilot only receives the extracted text, and it struggles with long or damaged text. In Word, run **Fix formatting**, then **Check document** and **Safe fixes**, on the document itself.

**Always read the result before using it.** The assistant lists anything it is unsure of under "queries for the author". It must not change figures, names, dates, quotations or what was decided. Report it if it ever does.

## 3. Using the assistant with Claude

- **Claude (web or desktop app):** package the skill as `ipbes-editorial.zip` (see section 4). Upload it under **Settings → Capabilities → Skills**. Claude then uses the skill whenever you ask it to edit an IPBES document.
- **Claude Code:** copy the `ipbes-editorial` folder into `~/.claude/skills/` (`cp -r ipbes-editorial ~/.claude/skills/`).

You ask for things the same way as in section 2. The difference is that Claude can run code on the file itself:

- **Fixing:** Claude writes the change table and applies it to your original Word file with `apply_changes.py`. You get your document back with tracked changes, comments giving the rules, and all formatting kept. You don't need to run the macro yourself.
- **Conversion:** Claude uses the converter script with the IPBES template, so the Word file has the IPBES styles.

The same safety rules apply. A change is made only if its original text is found exactly once, and anything else is reported to you.

---


## 4. Maintenance

Run these commands from the repository folder. The scripts need Python 3; the Word tools also need python-docx (`pip install --user python-docx`).

| Task | Command |
|---|---|
| Add or update the IPBES styles and ribbon tab in the Word add-in | `python3 tools/prepare_addin.py "path/to/IPBES Editorial.dotm"` (close Word first; see "Updating the add-in" in section 1) |
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
- re-package the skill (`python3 tools/package_skill.py`) and upload the new zip to the Copilot agent and to Claude;
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


## Contents

| Path | Contents |
|---|---|
| [ipbes-editorial/SKILL.md](ipbes-editorial/SKILL.md) | Instructions Claude reads first |
| [ipbes-editorial/copilot/ipbes-editorial-instructions.txt](ipbes-editorial/copilot/ipbes-editorial-instructions.txt) | The same rules, condensed for Copilot |
| `ipbes-editorial/references/` | The full rules: [house-style.md](ipbes-editorial/references/house-style.md) (voice, countries, style, format), [ipbes-terms.md](ipbes-editorial/references/ipbes-terms.md) (bodies, terms, assessments, symbols), [report-structure.md](ipbes-editorial/references/report-structure.md) (masthead, Word styles, standard wording), [inclusive-language.md](ipbes-editorial/references/inclusive-language.md), `countries.csv`, `un-spelling-list.md` (October 2026) and [evidence.md](ipbes-editorial/references/evidence.md) |
| `ipbes-editorial/scripts/` | `md_to_ipbes_docx.py` (Markdown-to-Word converter) and `apply_changes.py` (applies a change table to a Word file as tracked changes) |
| `ipbes-editorial/assets/` | Empty Word template with the IPBES styles (`ipbes-document-template.docx`), and reference copies of the masthead logos (`masthead/`) |
| `word-macro/` | `IPBESEditorial.bas` (the Word macros) and `ipbes-ribbon.xml` (the IPBES ribbon tab); see section 1 |
| `test/` | Test files for all the tools: `ipbes-test-document.docx` (a document with typical errors), `ipbes-test-unformatted.docx` (a document in Word's default formatting, for **Fix formatting**), `ipbes-test-change-list.docx` (a change table for that document) and `ipbes-test-draft.md` (a Markdown draft with typical errors) |
| `tools/` | Maintenance scripts (see section 4) |
| `assets/` | IPBES logos, motifs and Visual Guidelines pages. These are brand and web material, kept here for now (see section 5). |

---

**Sources:**

- UN Editorial Manual (https://www.un.org/dgacm/en/content/editorial-manual) and UN spelling list (October 2026)
- UN *Guidelines for gender-inclusive language in English*
- UN *Disability-Inclusive Communications Guidelines* (2022)
- IPBES geography list
- IPBES assessment abbreviations list
- IPBES Plenary reports IPBES/1/12 to IPBES/12/11
- IPBES *Logo and Visual Identity Guidelines* (2021)
