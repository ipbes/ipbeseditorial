"""Add the IPBES ribbon tab to the Word add-in "IPBES Editorial.dotm".

The .dotm is a zip package. This script adds word-macro/ipbes-ribbon.xml as
customUI/customUI14.xml and registers it, leaving everything else (including
the macros) untouched. Running it again replaces the ribbon with the current
version of ipbes-ribbon.xml. A backup copy (.bak) is kept.

Close Word before running it. Standard library only.

Usage (from the repository folder, in WSL):
    python3 tools/add_ribbon.py "/mnt/c/Users/<you>/AppData/Roaming/Microsoft/Word/STARTUP/IPBES Editorial.dotm"

On Windows with Python installed:
    python tools\\add_ribbon.py "%APPDATA%\\Microsoft\\Word\\STARTUP\\IPBES Editorial.dotm"
"""
import os
import re
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
RIBBON = os.path.join(HERE, '..', 'word-macro', 'ipbes-ribbon.xml')
PART = 'customUI/customUI14.xml'
REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
REL_ID = 'rIPBESRibbon'


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    dotm = sys.argv[1]
    if not os.path.exists(dotm):
        sys.exit('Not found: %s' % dotm)
    if not zipfile.is_zipfile(dotm):
        sys.exit('Not a Word package: %s' % dotm)
    ribbon = open(RIBBON, 'rb').read()

    with zipfile.ZipFile(dotm) as z:
        names = z.namelist()
        if 'word/vbaProject.bin' not in names:
            print('Warning: no macros found in this file (word/vbaProject.bin missing). '
                  'Import IPBESEditorial.bas into it first, or the buttons will not work.')
        parts = {n: z.read(n) for n in names}
        infos = {n: z.getinfo(n) for n in names}

    # 1. ribbon part (replace any earlier version)
    parts[PART] = ribbon

    # 2. package relationship
    rels = parts['_rels/.rels'].decode('utf-8')
    rels = re.sub(r'<Relationship [^>]*Type="%s"[^>]*/>' % re.escape(REL_TYPE), '', rels)
    rels = rels.replace('</Relationships>',
                        '<Relationship Id="%s" Type="%s" Target="%s"/></Relationships>' % (REL_ID, REL_TYPE, PART))
    parts['_rels/.rels'] = rels.encode('utf-8')

    # 3. content type for .xml parts (normally already present)
    ct = parts['[Content_Types].xml'].decode('utf-8')
    if not re.search(r'<Default [^>]*Extension="xml"', ct):
        ct = re.sub(r'(<Types[^>]*>)', r'\1<Default Extension="xml" ContentType="application/xml"/>', ct, count=1)
        parts['[Content_Types].xml'] = ct.encode('utf-8')

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
    print('IPBES ribbon added to %s' % dotm)
    print('Backup of the previous file: %s' % backup)
    print('Start Word: the IPBES tab appears after the Review tab.')


if __name__ == '__main__':
    main()
