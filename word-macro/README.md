# IPBES editorial macros for Word

These macros check a Word document against the IPBES house style and mark possible issues as comments. A second macro makes a set of safe corrections with Track Changes on, so you can accept or reject each one.

They apply the same rules as the IPBES editorial assistant (`skills/ipbes-editorial/`) but **work without AI and without an internet connection**.

| Macro | What it does | Changes the text? |
|---|---|---|
| **IPBES_CheckDocument** | Adds a comment starting "[IPBES]" at each possible issue and opens a summary report. Checks: the Platform's name, internal codes (BBA, NXS…), abbreviations, numbers and dates, punctuation, spelling, terms, country names, inclusive language, the masthead (logos, symbol, date, spacing) and heading numbering. | No |
| **IPBES_SafeFixes** | Corrects safe items **with Track Changes on**: the shortened Platform name, "e.g." → "e.g.,", -ise → -ize, em dashes and spaced hyphens → en dashes, double spaces, straight → curly quotes, Türkiye, Czechia, Viet Nam and similar names, "Indigenous Peoples", and a few spellings. | Yes, tracked |
| **IPBES_CheckVoice** | Marks prescriptive wording ("must", "should", "need to", "unprecedented"…) for review. IPBES is policy-relevant, not policy-prescriptive. | No |
| **IPBES_RemoveIPBESComments** | Deletes only the comments the macros added. Other people's comments are kept. | No |

**Important:**

- **Agreed text** (decisions, approved summaries for policymakers, quotations) must not be changed. Reject any tracked change there and treat comments there as queries.
- Comments are suggestions, not errors. For example, "should" in reported speech or "Secretariat" in a budget table is correct.
- To keep the document readable, the macros mark at most 25 places per rule. The report shows the full count.

---

## Installing (programme assistant, once)

### Option 1: for one computer

1. In Word, press **Alt+F11** to open the VBA editor. On a Mac, use **Tools → Macro → Visual Basic Editor**.
2. In the left pane, select **Normal** (or **Project (Normal)**).
3. Choose **File → Import File…** and select `IPBESEditorial.bas`.
4. Choose **Debug → Compile Normal**. If nothing happens, the macros compiled correctly.
5. Close the editor.

### Option 2: share with the team as a Word add-in (recommended)

1. Do the steps of option 1 in a **new blank document**, but select that document's project (for example **Project (Document1)**) instead of Normal.
2. Save the document as a **Word Macro-Enabled Template (*.dotm)** named `IPBES Editorial.dotm`.
3. Each colleague copies `IPBES Editorial.dotm` into their Word **Startup** folder:
   - **Windows:** `%APPDATA%\Microsoft\Word\STARTUP`. Paste this into the File Explorer address bar.
   - **Mac:** `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word`
4. Restart Word. The macros are now available in every document.

### Adding buttons

To add buttons to the Quick Access Toolbar:

1. Go to **File → Options → Quick Access Toolbar**.
2. Under "Choose commands from", select **Macros**.
3. Add **IPBES_CheckDocument**, **IPBES_SafeFixes** and **IPBES_RemoveIPBESComments**.
4. Use **Modify…** to give each one an icon and a short name, for example "IPBES check", "IPBES fix" and "Clear IPBES comments".

### Macro security

Word may block unsigned macros. If the macros don't appear or don't run:

- Check **File → Options → Trust Center → Trust Center Settings → Trusted Locations**. The Startup folder is normally trusted.
- If the organization blocks macros, ask IT to sign `IPBES Editorial.dotm` or to trust its location. Do not lower macro security for all documents.

---

## Using the macros

1. Open the document and **save a copy first**.
2. Run **IPBES_CheckDocument**, either with the toolbar button or through **View → Macros → IPBES_CheckDocument → Run**.
3. Read the report that opens in a new window. Go through the comments with **Review → Next comment**.
4. Optional: run **IPBES_SafeFixes**. Then review the tracked changes (**Review → Next**) and accept or reject each one.
5. When you've finished, run **IPBES_RemoveIPBESComments**.

## First-time test (please do this before sharing)

The macros were written without access to Word. Before sharing them, test them on a copy of a recent Plenary report and on a deliberately "bad" test document. Check that:

- [ ] The module compiles (**Debug → Compile**).
- [ ] **IPBES_CheckDocument** on the IPBES-12 report (IPBES/12/11) gives few or no masthead comments. That report is the reference for the masthead check.
- [ ] **Straight quotes:** type `"test"` with AutoCorrect off and run the check. Only straight quotes should be marked, not curly ones. If curly quotes are marked too, tell the person maintaining the macros.
- [ ] **IPBES_SafeFixes** on a test sentence such as `The Intergovernmental Platform on Biodiversity and Ecosystem Services (e.g. in Turkey) will prioritise  well being - whilst...` produces tracked changes only.
- [ ] **IPBES_RemoveIPBESComments** leaves other people's comments in place.

## Maintaining the rules

The rules are in the `BuildRules` procedure. Each rule is one `AddRule` line with:

- category
- text to find
- wildcards (True/False)
- match case
- whole word
- the message shown in the comment
- (optional) the replacement and whether it is a safe fix

Use `ChrW()` for any non-ASCII character, as in the Türkiye rule; the VBA editor does not import them reliably. After changing the rules, update `skills/ipbes-editorial/references/house-style.md` as well, so the macros and the AI assistant stay consistent.
