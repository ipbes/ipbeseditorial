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
Applies the IPBES layout, using the IPBES styles carried in the add-in. It asks whether the document is a report with numbered paragraphs. It then: copies the IPBES Word styles into the document; sets A4 paper and the IPBES margins; turns Title, Heading 1–4 and Caption into BB_Title, CH1–CH4 and Title_table, laying out heading numbers as tab, "I.", tab; puts body text in Normal_number (numbered) or Normal_No_number, removing typed paragraph numbers and manual fonts, sizes, colours and indents (bold and italic are kept); deletes empty spacer paragraphs (the IPBES styles provide the spacing); turns "Table 1: Title" captions and "(Number)" headnotes into the IPBES table title; and formats tables (indented to the text, horizontal rules only, no shading, italic column heads, numbers right-aligned). Documents that already use the IPBES styles (such as recent Plenary reports) keep their own style definitions, headers, footers, table layout and Normal paragraphs; only other paragraphs, captions and over-wide tables are changed, so the run is quick. It ends with a short summary; the tracked changes are the detailed record, and Ctrl+Z undoes the whole run. Lists, empty paragraphs and unrecognised styles are listed for you to check. Style and text changes are tracked; page set-up and style definitions are not.

### IPBES_SafeFixes
Makes safe corrections **with Track Changes on**: the shortened Platform name, "e.g." → "e.g.,", -ise → -ize, em dashes and spaced hyphens → en dashes, double spaces, straight → curly quotes, country names such as Türkiye and Czechia, "Indigenous Peoples", "spatial planning and connectivity assessment", and a few spellings.

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

- **Agreed text** (decisions, approved summaries for policymakers, quotations) must not be changed. Reject any tracked change there, and treat comments there as queries.
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

The assistant can **analyse** a document against the house style, **fix** it, and **convert** a Markdown draft into a Word document in the IPBES style. Copilot **cannot produce a document with tracked changes**. To set up a reusable Copilot agent:

1. Open **Microsoft 365 Copilot Chat** and choose **Create agent** (it may be called **New agent**).
2. On the **Configure** tab, fill in the fields:
   - **Name:** IPBES editorial assistant
   - **Description:** Edits and standardizes English IPBES documents in the IPBES and UN house style.
   - **Instructions:** paste the whole text of `ipbes-editorial/copilot/ipbes-editorial-instructions.txt`.
3. Under **Skills**, select **Add** and upload `ipbes-editorial.zip`. It holds the full rules (house style, terms, country list, spelling list). To make it, run `python3 tools/package_skill.py` in the repository folder. Leave **Knowledge** empty.

4. Select **Create**. Share the agent with colleagues if that is allowed, so that only one person builds and uploads the zip.

If you can't or don't want to set up an agent, start a new Copilot chat and paste the whole text of the instructions file. Then make your request, in the same message or the next one.

### What to ask

Attach the document, then ask one of three things:

1. **Analysis:** "Analyse the attached document against the IPBES editorial criteria."

   The assistant goes through the whole document and reports where it departs from the house style and formatting requirements, grouped by category. It does not change anything.
2. **Fixing:** "Fix the attached document to match the IPBES style and formatting requirements."

   Copilot cannot change formatting: run **Fix formatting** (section 1) first; Copilot handles the wording. It returns a **change table** (No. | Original | Revised | Rule) and a list of queries for the author. It does not return a new file, because a file created by Copilot loses the Word styles and cannot show tracked changes. To apply the table:
   1. Copy the table from Copilot and paste it into a new Word document.
   2. Open the document to correct, and run the **IPBES_ApplyChangeList** macro (section 1).

   Each change is made as a tracked change, with the rule as a comment, and the formatting is kept. You then accept or reject each change.
3. **Conversion:** "Convert the attached Markdown draft into a Word document in the IPBES style."

   The assistant applies the language rules and returns a Word file, but without the IPBES Word styles. Open the file in Word, press **Fix formatting** on the IPBES tab, then choose **Review → Accept → Accept All Changes**. The file now has the IPBES styles, numbering and layout.

**Always read the result before using it.** The assistant lists anything it is unsure of under "queries for the author". It must not change figures, names, dates, quotations or decisions. Report it if it ever does.

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
