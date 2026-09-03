import zipfile, re, os, json, sys
from xml.etree import ElementTree as ET

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
OUT = r"C:\Users\lyq15\WorkBuddy\2026-09-03-11-20-34\ppt_extract"

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

z = zipfile.ZipFile(SRC)

# list slides
slide_names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
                     key=lambda x: int(re.search(r'(\d+)', x.split('/')[-1]).group(1)))
print("slides:", len(slide_names))

# media inventory
media = [n for n in z.namelist() if n.startswith('ppt/media/')]
sizes = [(n, z.getinfo(n).file_size) for n in media]
sizes.sort(key=lambda x: -x[1])
print("media count:", len(media))
for n, s in sizes[:30]:
    print("  ", n, round(s/1024), "KB")

os.makedirs(os.path.join(OUT, 'media'), exist_ok=True)

def shape_text(sp):
    paras = []
    for p in sp.iter('{%s}p' % NS['a']):
        buf = []
        for t in p.iter('{%s}t' % NS['a']):
            buf.append(t.text or '')
        line = ''.join(buf).strip()
        if line:
            paras.append(line)
    return paras

slides_data = []
for sn in slide_names:
    root = ET.fromstring(z.read(sn))
    idx = int(re.search(r'(\d+)', sn.split('/')[-1]).group(1))
    items = []
    # spTree children
    tree = root.find('.//p:cSld/p:spTree', NS)
    if tree is None:
        tree = root.find('.//{%s}spTree' % NS['p'])
    if tree is not None:
        for child in tree:
            tag = child.tag.split('}')[-1]
            if tag in ('sp', 'graphicFrame', 'pic', 'cxnSp', 'grpSp'):
                texts = []
                for p in child.iter('{%s}p' % NS['a']):
                    buf = []
                    for t in p.iter('{%s}t' % NS['a']):
                        buf.append(t.text or '')
                    line = ''.join(buf).strip()
                    if line:
                        texts.append(line)
                # picture?
                ispic = child.find('.//{%s}blip' % NS['a']) is not None
                # table?
                has_table = child.find('.//a:tbl', NS) is not None
                if texts or ispic or has_table:
                    items.append({
                        'type': 'pic' if ispic and not texts and not has_table else ('table' if has_table else tag),
                        'texts': texts,
                    })
    # notes
    notes = None
    slides_data.append({'idx': idx, 'items': items})

with open(os.path.join(OUT, 'slides.json'), 'w', encoding='utf-8') as f:
    json.dump(slides_data, f, ensure_ascii=False, indent=1)

for s in slides_data:
    print('=' * 60)
    print('SLIDE', s['idx'])
    for it in s['items']:
        print('  [%s]' % it['type'], ' | '.join(it['texts'])[:300])
