# Evidence: what the IPBES Plenary reports actually do

This file records why each **IPBES practice** in `house-style.md` differs from, or adds to, the UN Editorial Manual. The rule is to follow the UN Editorial Manual unless the reports show a **deliberate and consistent** different practice.

## Corpus

The corpus is the 12 reports of the Plenary on the work of its sessions:

| Session | Symbol | Year | Venue |
|---|---|---|---|
| IPBES-1 | IPBES/1/12 | 2013 | Bonn |
| IPBES-2 | IPBES/2/17 | 2013 | Antalya |
| IPBES-3 | IPBES/3/18 | 2015 | Bonn |
| IPBES-4 | IPBES/4/19 | 2016 | Kuala Lumpur |
| IPBES-5 | IPBES/5/15 | 2017 | Bonn |
| IPBES-6 | IPBES/6/15 | 2018 | Medellin |
| IPBES-7 | IPBES/7/10 | 2019 | Paris |
| IPBES-8 | IPBES/8/11 | 2021 | Online |
| IPBES-9 | IPBES/9/14 | 2022 | Bonn |
| IPBES-10 | IPBES/10/12 | 2023 | Bonn |
| IPBES-11 | IPBES/11/12 | 2024 | Windhoek |
| IPBES-12 | IPBES/12/11 | 2026 | Manchester |

**Method.**

1. `tools/docx_dump.py` converts each report to text, keeping the Word style of each paragraph.
2. `tools/evidence_scan.py` counts competing forms in the body text, including annexed decisions.
3. Each count was then checked in context.

To repeat the scan on new reports:

```
python3 tools/docx_dump.py report.docx dump/2027-13.txt
python3 tools/evidence_scan.py dump/ > evidence_counts.md
```

**Test applied.**

- **IPBES practice:** the form is used in most reports, including the recent ones, and its rival is rare or absent.
- **UN rule:** the reports are split, or the rival form appears only in older reports.
- **Current IPBES practice:** practice changed at a clear point and has been stable since. The newer form is used.

The annexes contain agreed text, including decisions, scoping reports and terms of reference. A form that appears only there is noted. This describes the evidence only: before submission to UNDOCS, the secretariat edits such texts like any other (see `house-style.md`, "Every document is open for editing").

---

## 1. Confirmed IPBES practice (differs from or adds to the UN Editorial Manual)

| Practice | Counts (reports IPBES-1 → IPBES-12) | Conclusion |
|---|---|---|
| **"e.g.," and "i.e.," with a comma** | "e.g.," 158 against "e.g. " 3, in 10 of 12 reports. "i.e.," 27 against 0. | IPBES practice: keep the comma. The UN Manual has no comma. |
| **"Plenary" capitalized** as the name of the governing body | 2,333 against 22 lower case. The lower-case uses are "plenary session" and similar. | IPBES practice. |
| **"per cent" in text; no "%" in body text** | "per cent" 146; "%" 0 in body text; "percent" 1 (IPBES-2) | Confirms the UN rule. Noted because drafts often use "%". |
| **Nature's contributions to people**, lower case | 171 against 0 capitalized | IPBES term, lower case. |
| **Summary for policymakers**, lower case | 442 against 3 capitalized (IPBES-9, in a publication title) | IPBES term, lower case. |
| **"policymakers"**, one word | 551 against 3 "policy makers" (IPBES-4) | IPBES practice; the UN list agrees. |
| **"decision makers"** (noun, open); **"decision-making"** | 52 open against 6 "decision-makers" (IPBES-8 only, in an annex). "decision-making" 149. | UN list form, confirmed. |
| **Indigenous** capitalized | IPBES-1 to IPBES-9: lower case dominant (730 lower, 45 capitalized). IPBES-10 to IPBES-12: 294 capitalized, 1 lower. "Indigenous Peoples" 88 against 0 since IPBES-10. | Current IPBES practice since IPBES-10, in line with current UN usage. |
| **Türkiye** | "Turkey" until IPBES-8; "Türkiye" in IPBES-9 to IPBES-12 (7 against 0) | UN name change of 2022, followed. |
| **Czechia** | "Czech Republic" until IPBES-5; "Czechia" since IPBES-6 | UN short name, followed. |
| **"Netherlands (Kingdom of the)"** in lists | Since IPBES-10 (9 uses) | UNTERM change, followed. |
| **Names and courtesy titles** | IPBES-1 to IPBES-9: "Ms. Anne Larigauderie" at first mention. IPBES-10 to IPBES-12: full name with no title at first mention ("Anne Larigauderie, IPBES Executive Secretary"; "the Chair of IPBES, David Obura"), then "Ms. Larigauderie". "Mrs.", "Dr." and "H.E." never occur. | Current IPBES practice since IPBES-10. |
| **Co-Chair** | Assessment and task force co-chairs are lower case through IPBES-10 ("co-chairs", 311 uses). IPBES-11 and IPBES-12 capitalize specific Co-Chairs ("the Co-Chairs of the assessment", "the Co-Chairs of working group II"; 15 uses) but keep lower case in generic lists in annexes ("with the co-chairs, coordinating lead authors and lead authors"). The verb is always "co-chaired". | Current practice: capitalize when referring to specific people, lower case when generic. This is consistent with the UN rule for titles. |
| **Assessment short names** in lower case in parentheses after the full title | For example "(business and biodiversity assessment)" 20, "(nexus assessment)" 20, "(transformative change assessment)" 24 | IPBES practice. The internal three-letter codes (BBA, NXS, TCA…) never appear in any report. |
| **Decision text** | The first word or words of preambular and operative paragraphs are in italics: 57 of 64 in IPBES-12 and 41 of 45 in IPBES-3. The exceptions are headings and agenda items that begin with the same word. | UN practice for decisions, confirmed. |
| **"the Platform" and "IPBES"** | Both are used in every report. "the Platform" dominated until IPBES-4; "IPBES" has dominated since IPBES-5. | Both are acceptable. Prefer "IPBES" in narrative and "the Platform" in formulaic and agreed text ("members of the Platform"). |
| **Full name of IPBES** | "Intergovernmental Science-Policy Platform on Biodiversity and Ecosystem Services": 225 uses. The shortened form "Intergovernmental Platform on Biodiversity and Ecosystem Services" occurs **once** (IPBES-6, in an event title). | Never shorten the name. That single occurrence is an error. |

## 2. Inconsistent in the reports, so the UN rule applies

| Item | Counts | Rule applied |
|---|---|---|
| "and/or" | 69 in total. Recent reports have it almost only in annexed scoping reports and terms of reference (agreed text). | UN rule: avoid it. Replace it where the meaning is clear; otherwise query it. |
| "well-being" against "wellbeing" | 59 against 5. "wellbeing" appears in IPBES-3, 6, 8 and 10, mostly inside titles of other bodies' documents. | UN list: "well-being". |
| "Secretariat" capitalized | 147 uses, almost all in budget table line items ("3. Secretariat", "3.1 Secretariat personnel") or other bodies' names ("Secretariat of the Pacific Community") | Narrative: "the secretariat" (lower case, 902 uses). Leave budget tables as they are. |
| Bare "IPBES’" possessive | 11, IPBES-5 only | No possessive on acronyms (UN rule). |
| Straight quotation marks and apostrophes | Straight double quotes 23 (IPBES-4 and IPBES-5), straight apostrophes 20. Curly forms appear thousands of times. | Curly, always. |
| Date ranges with a hyphen ("21-26 January") | 1 (IPBES-1) against 11 with an en dash | En dash. |

## 3. Forms that never occur (useful negative checks)

These forms appear 0 times in all 12 reports, except where a count is given. If one turns up in a draft, it is almost certainly wrong:

- -ise spellings (recognise, organise, emphasise…). The -ize forms appear 1,191 times.
- "program" (as against "programme"), "whilst", "amongst", "co-operation", "inter-sessional", "sub-region", "on-line", "e-mail", "nongovernmental", "multistakeholder"
- "Chairman", "Chairperson", "Chairwoman", "Mrs.", "Dr.", "H.E.", "His Excellency"
- "US$" or "USD" (almost never: 3 uses in IPBES-1 and 3 in IPBES-4; "$" is used in about 240 places)
- times with a colon ("10:15"). The form used is "10.15 a.m."
- em dash (—). The spaced en dash is used for asides; an unspaced en dash is used in tables as the "nil" symbol.
- "IPBES’s" or "IPBES's"

## 4. Word template evidence (IPBES-10 to IPBES-12)

The paragraph styles have been the same since IPBES-2. The masthead layout has been stable since IPBES-10. Details are in `report-structure.md`.

| Style | Use | Count in IPBES-12 |
|---|---|---|
| Normal_number | numbered paragraphs | 258 |
| Normal_No_number | unnumbered text and agenda lists | 101 |
| CH1, CH2 | chapter and section headings | 45, 28 |
| Normal-pool | decision lists and annexes | 74 |
| Title_table | table number, title and headnote | 14 |
| BB_Title | report title | 1 |

IPBES-1 (2013) used a different template (`__Single Txt`, `_ H _Ch`, `_ H_1`). It is not used as evidence for format.

## 5. Discrepancies with other IPBES sources

- **Spatial planning assessment.** The IPBES assessment abbreviations list (IPBES_Assessment_Abbreviations.pdf) gives the short name "Spatial planning assessment". The reports use "spatial planning and connectivity assessment" (IPBES-11: 9; IPBES-12: 16). "Spatial planning assessment" appears 3 times, all in IPBES-11. **Confirmed by the secretariat (13 November 2026): "spatial planning and connectivity assessment" is correct**; the PDF is outdated on this point.
- **Transformative change assessment.** The full title in the PDF is cut short. The reports give: "thematic assessment of the underlying causes of biodiversity loss and the determinants of transformative change and options for achieving the 2050 Vision for Biodiversity".
- **Business and biodiversity assessment.** The PDF capitalizes the full title ("Methodological Assessment of the Impact…"). The reports use lower case in running text.
