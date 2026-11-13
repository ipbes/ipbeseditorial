"""Package the ipbes-editorial/ folder as a Claude skill: ipbes-editorial.zip.

Claude expects the zip to contain one folder named after the skill
(ipbes-editorial/, with SKILL.md inside), which is the folder name used in
this repository. Standard library only.

Usage (from the repository folder):
    python3 tools/package_skill.py [output.zip]
"""
import os
import sys
import zipfile

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
SRC = os.path.join(ROOT, 'ipbes-editorial')
NAME = 'ipbes-editorial'


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, NAME + '.zip')
    n = 0
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for folder, dirs, files in os.walk(SRC):
            dirs[:] = [d for d in dirs if d != '__pycache__' and not d.startswith('.')]
            for f in files:
                if f.endswith('.pyc') or f.startswith('.'):
                    continue
                path = os.path.join(folder, f)
                z.write(path, os.path.join(NAME, os.path.relpath(path, SRC)))
                n += 1
    print('%s written (%d files)' % (os.path.normpath(out), n))


if __name__ == '__main__':
    main()
