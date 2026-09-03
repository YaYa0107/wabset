import zipfile, re, os, json, shutil
from xml.etree import ElementTree as ET

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
OUT = r"C:\Users\lyq15\WorkBuddy\2026-09-03-11-20-34\ppt_extract"
SITE = r"C:\Users\lyq15\WorkBuddy\2026-09-03-11-20-34\website"

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'a': A, 'p': P, 'r': R}

z = zipfile.ZipFile(SRC)

# --- presentation size ---
pres = ET.fromstring(z.read('ppt/presentation.xml'))
sz = pres.find('.//p:sldSz', NS)
print('slide size:', sz.attrib if sz is not None else None)

# --- theme colors ---
try:
    theme = ET.fromstring(z.read('ppt/theme/theme1.xml'))
    clr = theme.find('.//a:clrScheme', NS)
    print('--- theme color scheme:', clr.attrib.get('name'))
    for c in clr:
        nm = c.tag.split('}')[-1]
        for ch in c:
            print('   ', nm, ch.attrib.get('val') or ch.attrib.get('lastClr'))
except Exception as e:
    print('theme err', e)

# --- copy media ---
mg = os.path.join(SITE, 'assets', 'img')
os.makedirs(mg, exist_ok=True)
media = [n for n in z.namelist() if n.startswith('ppt/media/')]
for n in media:
    base = n.split('/')[-1]
    if not base:
        continue
    with open(os.path.join(mg, base), 'wb') as f:
        f.write(z.read(n))
print('media copied:', len(media))

# --- slide -> rels mapping ---
def slide_images(idx):
    rels = 'ppt/slides/_rels/slide%d.xml.rels' % idx
    if rels not in z.namelist():
        return []
    r = ET.fromstring(z.read(rels))
    out = []
    for rel in r:
        t = rel.attrib.get('Target', '')
        if 'media/' in t:
            out.append(t.split('/')[-1])
    return out

def paras_text(el):
    res = []
    for p in el.iter('{%s}p' % A):
        buf = []
        for t in p.iter('{%s}t' % A):
            buf.append(t.text or '')
        line = ''.join(buf).strip()
        if line:
            res.append(line)
    return res

def table_data(tbl):
    rows = []
    for tr in tbl.findall('a:tr', NS):
        cells = []
        for tc in tr.findall('a:tc', NS):
            txt = ' '.join(paras_text(tc)).strip()
            cells.append(txt)
        rows.append(cells)
    return rows

slides = []
names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
               key=lambda x: int(re.search(r'(\d+)', x.split('/')[-1]).group(1)))
for idx, sn in enumerate(names, start=1):
    root = ET.fromstring(z.read(sn))
    tree = root.find('.//p:cSld/p:spTree', NS)
    items = []
    if tree is not None:
        for child in tree:
            tag = child.tag.split('}')[-1]
            if tag == 'sp':
                tbl = child.find('.//a:tbl', NS)
                if tbl is not None:
                    items.append({'kind': 'table', 'rows': table_data(tbl)})
                    continue
                texts = paras_text(child)
                if texts:
                    items.append({'kind': 'text', 'texts': texts})
            elif tag == 'pic':
                items.append({'kind': 'pic', 'img': None})  # order-based below
            elif tag == 'graphicFrame':
                tbl = child.find('.//a:tbl', NS)
                if tbl is not None:
                    items.append({'kind': 'table', 'rows': table_data(tbl)})
                else:
                    t = paras_text(child)
                    if t:
                        items.append({'kind': 'text', 'texts': t})
    imgs = slide_images(idx)
    # assign images to pic items in order
    pi = 0
    for it in items:
        if it['kind'] == 'pic':
            it['img'] = imgs[pi] if pi < len(imgs) else None
            pi += 1
    slides.append({'n': idx, 'items': items, 'imgs': imgs})

with open(os.path.join(OUT, 'slides2.json'), 'w', encoding='utf-8') as f:
    json.dump(slides, f, ensure_ascii=False, indent=1)

for s in slides:
    print('=' * 50)
    print('SLIDE', s['n'], '| images:', ','.join(s['imgs']))
    for it in s['items']:
        if it['kind'] == 'table':
            print('  [TABLE]')
            for row in it['rows']:
                print('     ', row)
        elif it['kind'] == 'pic':
            print('  [PIC]', it['img'])
        else:
            for t in it['texts']:
                print('  [T]', t[:200])
