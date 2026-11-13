# IPBES editorial assistant: how to use it

The editorial assistant teaches an AI assistant the IPBES house style for English documents. The house style is the UN Editorial Manual plus the practices IPBES has followed consistently in its Plenary reports since 2013. With it, the assistant can:

- **edit** a document or a passage of text and list what it changed;
- **standardize** the language of a document (names, terms, spelling, numbers, punctuation);
- **convert** a draft written in Markdown into a Word document in the IPBES style.

It works with **Microsoft Copilot**, which is the tool the organization approves, and with **Claude**. Follow the organization's rules on which tool may be used for which documents.

---

## 1. Using it with Copilot

### Option A: a Copilot agent (recommended, set up once)

1. Open **Microsoft 365 Copilot Chat** and choose **Create agent** (it may be called **New agent**).
2. Switch to the **Configure** tab and fill in the fields:
   - **Name:** IPBES editorial assistant
   - **Description:** Edits and standardizes English IPBES documents in the IPBES and UN house style.
   - **Instructions:** open the file `copilot/ipbes-editorial-instructions.txt`, copy all of its text and paste it in.
3. Under **Knowledge**, upload these files from the `references` folder. They give the agent the full rules:
   - `house-style.md`
   - `ipbes-terms.md`
   - `report-structure.md`
   - `inclusive-language.md`
   - `countries.csv`

   If Copilot does not accept `.md` files, ask the programme assistant for the `.docx` versions (see section 3).
4. Select **Create**. Share the agent with colleagues if that is allowed.

### Option B: no agent

Start a new Copilot chat, paste the whole text of `copilot/ipbes-editorial-instructions.txt`, and then make your request in the same message or the next one.

### What to ask

- "Edit this paragraph to IPBES style and list what you changed."
- "Standardize the language in the attached document. Don't rewrite, just make it consistent."
- "Check the masthead of the attached document."
- "Turn this Markdown draft into an IPBES-style Word document."

**Always read the result before using it.** The assistant flags anything it is unsure of under "queries for the author". It must not change figures, names, dates, quotations or decisions; tell the programme assistant if it ever does.

---

## 2. Using it with Claude

- **Claude (web or desktop app):** the programme assistant makes `ipbes-editorial.zip` (see section 3). Upload it in Claude under **Settings → Capabilities → Skills**. Claude then uses the skill whenever you ask it to edit an IPBES document.
- **Claude Code:** copy the `ipbes-editorial` folder into `~/.claude/skills/`.

You ask for things in the same way as in section 1.

---

## 3. For the programme assistant

Run these commands in a terminal from the repository folder. The scripts need Python 3; the Word tools also need python-docx (`pip install --user python-docx`).

| Task | Command |
|---|---|
| Package the skill for Claude | `cd skills && zip -r ipbes-editorial.zip ipbes-editorial` |
| Convert a Markdown draft to Word | `python3 skills/ipbes-editorial/scripts/md_to_ipbes_docx.py draft.md draft.docx --numbered --symbol IPBES/13/5` (leave out `--numbered` for documents without numbered paragraphs) |
| Update the country list from IPBES | Download https://ibok.ipbes.net/rest/v1/geo/data?format=application/rdf%2Bxml and save it as `ipbes-geo.rdf`, then run `python3 tools/build_countries.py ipbes-geo.rdf skills/ipbes-editorial/references/countries.csv` |
| Rebuild the Word template from a new Plenary report | `python3 tools/make_template.py report.docx skills/ipbes-editorial/assets/ipbes-document-template.docx` |
| Add a new Plenary report to the evidence | `python3 tools/docx_dump.py IPBES-13-report.docx temp/dump/2027-13.txt`, then `python3 tools/evidence_scan.py temp/dump > temp/evidence_counts.md`. Compare the result with `references/evidence.md` and update the rules if practice has changed. |
| Make Word copies of the reference files (if Copilot refuses `.md`) | Open each `.md` file in Word (File → Open, "All files") and save it as `.docx`. Alternatively, if pandoc is installed: `pandoc house-style.md -o house-style.docx`. |

**After any change to the rules,**

- update `copilot/ipbes-editorial-instructions.txt`. Copilot's limit for agent instructions is 8,000 characters, and the file is close to it.
- re-upload the knowledge files to the Copilot agent.
- re-package the Claude skill.

---

## 4. What is in this folder

| Path | Contents |
|---|---|
| `SKILL.md` | Instructions Claude reads first |
| `copilot/ipbes-editorial-instructions.txt` | The same rules, condensed for Copilot |
| `references/house-style.md` | Full rules: voice, countries, inclusive language, style, format |
| `references/ipbes-terms.md` | IPBES bodies, terms, assessment titles and short names, document symbols |
| `references/report-structure.md` | Masthead, Word styles, report structure, standard wording |
| `references/inclusive-language.md` | Gender- and disability-inclusive language, accessible documents |
| `references/countries.csv` | Country names, IPBES membership and regions (from the IPBES geography list) |
| `references/un-spelling-list.md` | The UN spelling list (October 2026) |
| `references/evidence.md` | What the 12 Plenary reports show, and why each IPBES rule exists |
| `scripts/md_to_ipbes_docx.py` | Markdown-to-Word converter |
| `assets/ipbes-document-template.docx` | Empty Word template with the IPBES styles |
| `assets/masthead/` | Reference copies of the masthead logos |

Sources: UN Editorial Manual (https://www.un.org/dgacm/en/content/editorial-manual); UN guidelines for gender-inclusive language; UN Disability-Inclusive Communications Guidelines (2022); IPBES geography list; IPBES Plenary reports IPBES/1/12 to IPBES/12/11.
