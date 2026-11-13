"""Count competing spellings/forms across the IPBES Plenary reports.

Usage:
    python3 tools/evidence_scan.py DUMP_DIR > evidence_counts.md

DUMP_DIR holds the text dumps produced by tools/docx_dump.py (one per report,
named YYYY-N.txt). Prints a Markdown table: one row per pattern, one column per
report. Used to build skills/ipbes-editorial/references/evidence.md.
"""
import glob
import os
import re
import sys

# (label, regex) -- case-sensitive unless (?i) is given
PATTERNS = [
    # Platform name
    ('Intergovernmental Science-Policy Platform (full name)', r'Intergovernmental Science[-‑‐]Policy Platform'),
    ('"Intergovernmental Platform on Biodiversity" (wrong short form)', r'Intergovernmental Platform on Biodiversity'),
    ('"the Platform"', r'\bthe Platform\b'),
    ('"IPBES" (any)', r'\bIPBES\b'),
    ("IPBES’s / IPBES's", r'IPBES[’\']s\b'),
    ('IPBES’ (bare apostrophe)', r'IPBES[’\'](?!s)'),
    # Bodies and posts
    ('Plenary (capital)', r'\bPlenary\b'),
    ('plenary (lower case)', r'\bplenary\b'),
    ('secretariat (lower case)', r'\bsecretariat\b'),
    ('Secretariat (capital)', r'\bSecretariat\b'),
    ('Chair', r'\bChair\b'),
    ('Chairman / Chairperson', r'\bChair(?:man|person|woman)\b'),
    ('Co-Chair(s) (capital)', r'\bCo[-‑]Chairs?\b'),
    ('co-chair(s) (lower case noun)', r'\bco[-‑]chairs?\b'),
    ('Member States', r'\bMember States\b'),
    ('members of the Platform / IPBES members', r'members of (?:the Platform|IPBES)|IPBES members'),
    ('task force (lower case)', r'\btask forces?\b'),
    ('Task Force (capital)', r'\bTask Forces?\b'),
    ('work programme (lower case)', r'\bwork programme'),
    ('Work Programme (capital)', r'\bWork Programme'),
    ('summary/summaries for policymakers (lower)', r'\bsummar(?:y|ies) for policymakers'),
    ('Summary for Policymakers (capital)', r'\bSummar(?:y|ies) for Policymakers'),
    ('nature’s contributions to people', r'nature[’\']s contributions to people'),
    ('Nature’s Contributions to People (capital)', r'Nature[’\']s Contributions to People'),
    # Indigenous
    ('Indigenous (capital)', r'\bIndigenous\b'),
    ('indigenous (lower case)', r'\bindigenous\b'),
    ('Indigenous Peoples', r'\bIndigenous Peoples\b'),
    ('indigenous peoples', r'\bindigenous peoples\b'),
    # Titles of persons
    ('Ms.', r'\bMs\.'),
    ('Mrs.', r'\bMrs\.'),
    ('Mr.', r'\bMr\.'),
    ('Dr.', r'\bDr\.'),
    ('H.E. / His/Her Excellency', r'\bH\.E\.|\bH(?:is|er) Excellency'),
    # Abbreviations
    ('(e.g., ', r'e\.g\.,'),
    ('(e.g. ', r'e\.g\. '),
    ('i.e.,', r'i\.e\.,'),
    ('i.e. ', r'i\.e\. '),
    ('and/or', r'\band/or\b'),
    # Numbers
    ('per cent', r'\bper cent\b'),
    ('% (sign)', r'\d\s?%'),
    ('percent (one word)', r'\bpercent\b'),
    ('$ figure', r'\$\s?\d'),
    ('US$ / USD', r'US\$|\bUSD\b'),
    ('a.m./p.m. with dot (10.15 a.m.)', r'\b\d{1,2}\.\d\d [ap]\.m\.'),
    ('time with colon (10:15)', r'\b\d{1,2}:\d\d\b'),
    ('date range with hyphen (21-26 January)', r'\b\d{1,2}-\d{1,2} (?:January|February|March|April|May|June|July|August|September|October|November|December)'),
    ('date range with en dash (21–26 January)', r'\b\d{1,2}–\d{1,2} (?:January|February|March|April|May|June|July|August|September|October|November|December)'),
    ('"from X to Y" date', r'\bfrom \d{1,2} to \d{1,2} (?:January|February|March|April|May|June|July|August|September|October|November|December)'),
    # Spelling and compounds
    ('decision makers (open)', r'\bdecision makers?\b'),
    ('decision-makers (hyphen noun)', r'\bdecision[-‑]makers?\b'),
    ('decision-making', r'\bdecision[-‑]making\b'),
    ('policymakers', r'\bpolicymakers?\b'),
    ('policy makers / policy-makers', r'\bpolicy[ -‑]makers?\b'),
    ('capacity-building', r'\bcapacity[-‑]building\b'),
    ('capacity building (open)', r'\bcapacity building\b'),
    ('well-being', r'\bwell[-‑]being\b'),
    ('wellbeing', r'\bwellbeing\b'),
    ('intersessional', r'\bintersessional\b'),
    ('inter-sessional', r'\binter[-‑]sessional\b'),
    ('subregion(al)', r'\bsubregion'),
    ('sub-region(al)', r'\bsub[-‑]region'),
    ('cooperation', r'\bcooperation\b'),
    ('co-operation', r'\bco[-‑]operation\b'),
    ('email', r'\bemail'),
    ('e-mail', r'\be[-‑]mail'),
    ('online', r'\bonline\b'),
    ('on-line', r'\bon[-‑]line\b'),
    ('non-governmental', r'\bnon[-‑]governmental\b'),
    ('nongovernmental', r'\bnongovernmental\b'),
    ('multi-stakeholder', r'\bmulti[-‑]stakeholder'),
    ('multistakeholder', r'\bmultistakeholder'),
    ('-ize verbs (recognize, organize, emphasize…)', r'\b(?:recogni|organi|emphasi|mobili|synthesi|prioriti|summari|utili|finali|operationali|characteri)z(?:e|ed|es|ing|ation)'),
    ('-ise verbs (recognise, organise…)', r'\b(?:recogni|organi|emphasi|mobili|synthesi|prioriti|summari|utili|finali|operationali|characteri)s(?:e|ed|es|ing|ation)'),
    ('programme', r'\bprogrammes?\b'),
    ('program (not computer)', r'\bprograms?\b'),
    ('whilst', r'\bwhilst\b'),
    ('amongst', r'\bamongst\b'),
    ('judgement', r'\bjudgements?\b'),
    ('judgment', r'\bjudgments?\b'),
    # Countries
    ('Türkiye', r'Türkiye'),
    ('Turkey', r'\bTurkey\b'),
    ('Czechia', r'\bCzechia\b'),
    ('Czech Republic', r'\bCzech Republic\b'),
    ('Viet Nam', r'\bViet Nam\b'),
    ('Vietnam', r'\bVietnam\b'),
    ('United Kingdom of Great Britain and Northern Ireland', r'United Kingdom of Great Britain and Northern Ireland'),
    ('United States of America', r'United States of America'),
    ('Netherlands (Kingdom of the)', r'Netherlands \(Kingdom of the\)'),
    # Punctuation
    ('straight apostrophe in word', r"[A-Za-z]'[A-Za-z]"),
    ('curly apostrophe in word', r'[A-Za-z]’[A-Za-z]'),
    ('straight double quote', r'"'),
    ('curly double quotes', r'[“”]'),
    ('em dash —', r'—'),
    ('spaced en dash – ', r' – '),
    ('double space in text', r'[^\s]  +[^\s]'),
    # Decisions
    ('decision IPBES-n/n', r'decision IPBES[-‑]\d+/\d+'),
    ('IPBES-n session shorthand', r'\bIPBES[-‑]\d+\b(?!/)'),
    # Prescriptive words in report text (context needed)
    ('should', r'\bshould\b'),
    ('must', r'\bmust\b'),
]


def body_text(path):
    lines = []
    in_body = False
    for line in open(path, encoding='utf-8'):
        if line.startswith('[BODY]'):
            in_body = True
            continue
        if not in_body:
            continue
        # drop the [Style] prefix
        lines.append(re.sub(r'^\s*\[[^\]]*\]\s?', '', line))
    return ''.join(lines)


def main():
    d = sys.argv[1]
    files = sorted(f for f in glob.glob(os.path.join(d, '20*.txt')))
    texts = {os.path.basename(f)[:-4]: body_text(f) for f in files}
    cols = list(texts)
    print('| Pattern | ' + ' | '.join('IPBES-' + c.split('-')[1] for c in cols) + ' |')
    print('|---|' + '---:|' * len(cols))
    for label, rx in PATTERNS:
        r = re.compile(rx)
        counts = [len(r.findall(texts[c])) for c in cols]
        print('| ' + label.replace('|', '/') + ' | ' + ' | '.join(str(n) for n in counts) + ' |')


if __name__ == '__main__':
    main()
