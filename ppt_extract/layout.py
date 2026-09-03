import zipfile, re
from xml.etree import ElementTree as ET

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P, 'a': A, 'r': R}
z = zipfile.ZipFile(SRC)

names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
               key=lambda x: int(re.search(r'(\d+)', x.split('/')[-1]).group(1)))

EMU = 914400.0

def emu(v):
    return int(v) / EMU

for idx, sn in enumerate(names, 1):
    raw = z.read(sn).decode('utf-8')
    root = ET.fromstring(raw)
    # rels
    relmap = {}
    rels = 'ppt/slides/_rels/slide%d.xml.rels' % idx
    if rels in z.namelist():
        rr = ET.fromstring(z.read(rels))
        for rel in rr:
            relmap[rel.attrib['Id']] = rel.attrib.get('Target', '').split('/')[-1]
    tree = root.find('.//p:cSld/p:spTree', NS)
    if tree is None:
        continue
    out = []
    for child in tree:
        tag = child.tag.split('}')[-1]
        if tag != 'pic':
            continue
        blip = child.find('.//a:blip', NS)
        rid = blip.attrib.get('{%s}embed' % R) if blip is not None else None
        img = relmap.get(rid, '?')
        xf = child.find('.//a:xfrm', NS)
        if xf is not None:
            off = xf.find('a:off', NS)
            ext = xf.find('a:ext', NS)
            x, y = emu(off.attrib['x']), emu(off.attrib['y'])
            w, h = emu(ext.attrib['cx']), emu(ext.attrib['cy'])
        else:
            x = y = w = h = 0
        out.append((round(y, 2), round(x, 2), round(w, 2), round(h, 2), img))
    out.sort()
    print('SLIDE %d' % idx)
    for y, x, w, h, img in out:
        print('   y=%-7s x=%-7s w=%-6s h=%-6s %s' % (y, x, w, h, img))
