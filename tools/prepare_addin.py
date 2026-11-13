"""Prepare the Word add-in "IPBES Editorial.dotm": IPBES styles + IPBES ribbon tab.

The .dotm is a zip package. This script
  1. copies the IPBES Word styles and paragraph numbering (styles.xml and
     numbering.xml) from ipbes-editorial/assets/ipbes-document-template.docx
     into the add-in, so that the "Fix formatting" macro can apply them in any
     document;
  2. adds the IPBES ribbon tab (word-macro/ipbes-ribbon.xml).
Everything else, including the macros, is left untouched. Running it again
replaces both with the current versions. A backup copy (.bak) is kept.

Close Word before running it. Standard library only.

Usage (from the repository folder, in WSL):
    python3 tools/prepare_addin.py "/mnt/c/Users/<you>/AppData/Roaming/Microsoft/Word/STARTUP/IPBES Editorial.dotm"
Options:
    --no-styles     only add the ribbon
    --no-ribbon     only add the styles
"""
import os
import re
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
RIBBON = os.path.join(ROOT, 'word-macro', 'ipbes-ribbon.xml')
TEMPLATE = os.path.join(ROOT, 'ipbes-editorial', 'assets', 'ipbes-document-template.docx')

RIBBON_PART = 'customUI/customUI14.xml'
RIBBON_REL = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
NUMBERING_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
NUMBERING_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'


def add_ribbon(parts):
    parts[RIBBON_PART] = open(RIBBON, 'rb').read()
    rels = parts['_rels/.rels'].decode('utf-8')
    rels = re.sub(r'<Relationship [^>]*Type="%s"[^>]*/>' % re.escape(RIBBON_REL), '', rels)
    rels = rels.replace('</Relationships>',
                        '<Relationship Id="rIPBESRibbon" Type="%s" Target="%s"/></Relationships>' % (RIBBON_REL, RIBBON_PART))
    parts['_rels/.rels'] = rels.encode('utf-8')
    ct = parts['[Content_Types].xml'].decode('utf-8')
    if not re.search(r'<Default [^>]*Extension="xml"', ct):
        ct = re.sub(r'(<Types[^>]*>)', r'\1<Default Extension="xml" ContentType="application/xml"/>', ct, count=1)
        parts['[Content_Types].xml'] = ct.encode('utf-8')


def add_styles(parts):
    with zipfile.ZipFile(TEMPLATE) as t:
        parts['word/styles.xml'] = t.read('word/styles.xml')
        parts['word/numbering.xml'] = t.read('word/numbering.xml')
    # relationship from the main document to numbering.xml
    rp = 'word/_rels/document.xml.rels'
    rels = parts[rp].decode('utf-8')
    if NUMBERING_REL not in rels:
        rels = rels.replace('</Relationships>',
                            '<Relationship Id="rIPBESNumbering" Type="%s" Target="numbering.xml"/></Relationships>' % NUMBERING_REL)
        parts[rp] = rels.encode('utf-8')
    # content type for numbering.xml
    ct = parts['[Content_Types].xml'].decode('utf-8')
    if '/word/numbering.xml' not in ct:
        ct = ct.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="%s"/></Types>' % NUMBERING_CT)
        parts['[Content_Types].xml'] = ct.encode('utf-8')
    # stylesWithEffects.xml (older Word) would keep the old styles: drop it
    if 'word/stylesWithEffects.xml' in parts:
        del parts['word/stylesWithEffects.xml']
        rels = parts[rp].decode('utf-8')
        rels = re.sub(r'<Relationship [^>]*Target="stylesWithEffects.xml"[^>]*/>', '', rels)
        parts[rp] = rels.encode('utf-8')
        ct = parts['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override PartName="/word/stylesWithEffects.xml"[^>]*/>', '', ct)
        parts['[Content_Types].xml'] = ct.encode('utf-8')


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    flags = [a for a in sys.argv[1:] if a.startswith('--')]
    if not args:
        sys.exit(__doc__)
    dotm = args[0]
    if not os.path.exists(dotm):
        sys.exit('Not found: %s' % dotm)
    if not zipfile.is_zipfile(dotm):
        sys.exit('Not a Word package: %s' % dotm)
    if not dotm.lower().endswith('.dotm'):
        print('Warning: the file does not end in .dotm. Word loads only templates (.dotm) from the Startup folder.')

    with zipfile.ZipFile(dotm) as z:
        names = z.namelist()
        if 'word/vbaProject.bin' not in names:
            print('Warning: no macros found in this file (word/vbaProject.bin missing). '
                  'Import IPBESEditorial.bas into it first, or the buttons will not work.')
        parts = {n: z.read(n) for n in names}
        infos = {n: z.getinfo(n) for n in names}

    done = []
    if '--no-styles' not in flags:
        add_styles(parts); done.append('IPBES styles and numbering')
    if '--no-ribbon' not in flags:
        add_ribbon(parts); done.append('IPBES ribbon tab')

    backup = dotm + '.bak'
    shutil.copy2(dotm, backup)
    tmp = dotm + '.tmp'
    order = ['[Content_Types].xml'] + [n for n in parts if n != '[Content_Types].xml']
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in order:
            info = infos.get(n)
            zi = zipfile.ZipInfo(n, date_time=info.date_time if info else (2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, parts[n])
    os.replace(tmp, dotm)
    print('Added to %s: %s' % (dotm, ', '.join(done)))
    print('Backup of the previous file: %s' % backup)
    print('Start Word: the IPBES tab appears after the Review tab.')


if __name__ == '__main__':
    main()
