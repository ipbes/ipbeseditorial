"""Minimal stdlib PDF text extractor (FlateDecode streams, ToUnicode CMaps)."""
import re, sys, zlib

data = open(sys.argv[1], 'rb').read()

objs = {}
for m in re.finditer(rb'(\d+)\s+0\s+obj(.*?)endobj', data, re.S):
    objs[int(m.group(1))] = m.group(2)

def stream(body):
    m = re.search(rb'stream\r?\n(.*?)\r?\nendstream', body, re.S)
    if not m:
        return None
    raw = m.group(1)
    if b'FlateDecode' in body[:m.start()]:
        try:
            return zlib.decompress(raw)
        except zlib.error:
            try:
                return zlib.decompressobj().decompress(raw)
            except zlib.error:
                return None
    return raw

# Expand object streams
for n, body in list(objs.items()):
    if b'/ObjStm' in body[:300]:
        s = stream(body)
        if not s:
            continue
        first = int(re.search(rb'/First\s+(\d+)', body).group(1))
        hdr = s[:first].split()
        pairs = [(int(hdr[i]), int(hdr[i + 1])) for i in range(0, len(hdr) - 1, 2)]
        for i, (on, off) in enumerate(pairs):
            end = pairs[i + 1][1] if i + 1 < len(pairs) else len(s) - first
            objs.setdefault(on, s[first + off:first + end])

def parse_cmap(s):
    cmap = {}
    for blk in re.findall(rb'beginbfchar(.*?)endbfchar', s, re.S):
        for a, b in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', blk):
            cmap[int(a, 16)] = bytes.fromhex(b.decode()).decode('utf-16-be', 'ignore')
    for blk in re.findall(rb'beginbfrange(.*?)endbfrange', s, re.S):
        for a, b, c in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[^\]]*\])', blk):
            lo, hi = int(a, 16), int(b, 16)
            if c.startswith(b'<'):
                base = bytes.fromhex(c[1:-1].decode()).decode('utf-16-be', 'ignore') or ' '
                for k in range(lo, hi + 1):
                    cmap[k] = base[:-1] + chr(min(ord(base[-1]) + k - lo, 0x10FFFF))
            else:
                for k, h in zip(range(lo, hi + 1), re.findall(rb'<([0-9A-Fa-f]+)>', c)):
                    cmap[k] = bytes.fromhex(h.decode()).decode('utf-16-be', 'ignore')
    return cmap

def font_cmap(fobj):
    body = objs.get(fobj, b'')
    m = re.search(rb'/ToUnicode\s+(\d+)\s+0\s+R', body)
    if not m:
        return None, 1
    s = stream(objs.get(int(m.group(1)), b'')) or b''
    cm = parse_cmap(s)
    width = 2 if b'Type0' in body else 1
    return cm, width

def resolve_dict(body, key):
    m = re.search(rb'/' + key + rb'\s*(<<(?:[^<>]|<<[^<>]*>>)*>>|\d+\s+0\s+R)', body, re.S)
    if not m:
        return b''
    v = m.group(1)
    if v.endswith(b'R'):
        return objs.get(int(v.split()[0]), b'')
    return v

def page_fonts(page):
    res = resolve_dict(page, b'Resources')
    fd = resolve_dict(res, b'Font')
    return {name: font_cmap(int(n)) for name, n in re.findall(rb'/([\w.+-]+)\s+(\d+)\s+0\s+R', fd)}

def decode_str(raw, cm, width):
    if cm is None:
        return raw.decode('latin-1')
    out = []
    if width == 2:
        for i in range(0, len(raw) - 1, 2):
            out.append(cm.get(raw[i] << 8 | raw[i + 1], ''))
    else:
        for c in raw:
            out.append(cm.get(c, chr(c)))
    return ''.join(out)

def lit(s):
    out = bytearray(); i = 0
    while i < len(s):
        c = s[i]
        if c == 0x5c and i + 1 < len(s):
            n = s[i + 1]
            esc = {ord('n'): 10, ord('r'): 13, ord('t'): 9, ord('b'): 8, ord('f'): 12}
            if n in esc:
                out.append(esc[n]); i += 2
            elif 0x30 <= n <= 0x37:
                j = i + 1
                while j < len(s) and j < i + 4 and 0x30 <= s[j] <= 0x37:
                    j += 1
                out.append(int(s[i + 1:j], 8) & 255); i = j
            elif n in (10, 13):
                i += 2
            else:
                out.append(n); i += 2
        else:
            out.append(c); i += 1
    return bytes(out)

tok = re.compile(rb'\((?:\\.|[^\\()]|\((?:\\.|[^\\()])*\))*\)|<[0-9A-Fa-f\s]*>|\[|\]|/[^\s/\[\]()<>]+|-?[\d.]+|[A-Za-z\'"*]+')

def page_text(content, fonts):
    out = []; stack = []; cm, w = None, 1; arr = None
    for t in tok.findall(content):
        if t == b'[':
            arr = []
        elif t == b']':
            stack.append(arr); arr = None
        elif t[:1] in b'(<' and t[:2] != b'<<':
            raw = lit(t[1:-1]) if t[:1] == b'(' else bytes.fromhex(re.sub(rb'\s', b'', t[1:-1]).decode() + ('0' if len(re.sub(rb'\s', b'', t[1:-1])) % 2 else ''))
            s = decode_str(raw, cm, w)
            (arr if arr is not None else stack).append(s)
        elif arr is not None:
            try:
                if float(t) < -200:
                    arr.append(' ')
            except ValueError:
                pass
        elif t == b'Tf':
            name = stack[-2] if len(stack) >= 2 else b''
            if isinstance(name, bytes):
                cm, w = fonts.get(name[1:], (None, 1))
            stack = []
        elif t in (b'Tj', b"'", b'"'):
            if stack and isinstance(stack[-1], str):
                out.append(stack[-1])
            if t != b'Tj':
                out.append('\n')
            stack = []
        elif t == b'TJ':
            if stack and isinstance(stack[-1], list):
                out.append(''.join(x for x in stack[-1] if isinstance(x, str)))
            stack = []
        elif t in (b'Td', b'TD', b'T*', b'Tm'):
            if t == b'T*' or (t in (b'Td', b'TD') and len(stack) >= 2 and isinstance(stack[-1], bytes) and stack[-1] not in (b'0', b'0.0', b'-0')):
                out.append('\n')
            elif t == b'Tm':
                out.append('\n')
            else:
                out.append(' ')
            stack = []
        elif t == b'ET':
            out.append('\n'); stack = []
        else:
            stack.append(t)
    return ''.join(out)

# Pages in order via Kids traversal
root = re.search(rb'/Root\s+(\d+)\s+0\s+R', data)
cat = objs[int(root.group(1))] if root else next(b for b in objs.values() if b'/Catalog' in b)
pages_root = int(re.search(rb'/Pages\s+(\d+)\s+0\s+R', cat).group(1))

def walk(n):
    b = objs.get(n, b'')
    if re.search(rb'/Type\s*/Pages', b):
        kids = re.search(rb'/Kids\s*\[(.*?)\]', b, re.S).group(1)
        for k in re.findall(rb'(\d+)\s+0\s+R', kids):
            yield from walk(int(k))
    else:
        yield b

result = []
for i, page in enumerate(walk(pages_root), 1):
    fonts = page_fonts(page)
    c = re.search(rb'/Contents\s*(\[[^\]]*\]|\d+\s+0\s+R)', page)
    content = b''
    if c:
        for n in re.findall(rb'(\d+)\s+0\s+R', c.group(1)):
            content += (stream(objs.get(int(n), b'')) or b'') + b'\n'
    result.append(f'\n=== page {i} ===\n' + page_text(content, fonts))
open(sys.argv[2], 'w').write(''.join(result))
print('pages', len(result))
