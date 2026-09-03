import zipfile, re
from xml.etree import ElementTree as ET
from collections import Counter

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS = {'a': A, 'p': P}
z = zipfile.ZipFile(SRC)

names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
               key=lambda x: int(re.search(r'(\d+)', x.split('/')[-1]).group(1)))

bg = Counter()
solid = Counter()
fonts = Counter()
sizes = Counter()

for sn in names:
    raw = z.read(sn).decode('utf-8')
    root = ET.fromstring(raw)
    # background
    b = root.find('.//p:cSld/p:bg/p:bgPr', NS)
    if b is not None:
        for m in re.finditer(r'val="([0-9A-Fa-f]{6})"', ET.tostring(b).decode()):
            bg[m.group(1)] += 1
        # gradient stops in bg
        for gs in b.iter('{%s}gs' % A):
            pass
    # all srgbClr in spPr (shape fills)
    for sp in root.iter('{%s}sp' % P):
        spPr = sp.find('a:spPr', NS)
        if spPr is None:
            continue
        f = spPr.find('a:solidFill', NS)
        if f is not None:
            c = f.find('a:srgbClr', NS)
            if c is not None:
                solid[c.attrib['val']] += 1
    # fonts
    for m in re.finditer(r'<a:latin typeface="([^"]+)"', raw):
        fonts[m.group(1)] += 1
    for m in re.finditer(r'sz="(\d+)"', raw):
        sizes[int(m.group(1))//100] += 1

print('BG colors:', bg.most_common(12))
print('Solid fills:', solid.most_common(25))
print('Fonts:', fonts.most_common(8))
print('Font sizes(pt):', sorted(sizes.items(), key=lambda x: -x[1])[:15])

# sample slide1 & 2 raw snippet for title formatting
for sn in names[:3]:
    raw = z.read(sn).decode('utf-8')
    print('=====', sn)
    for m in re.finditer(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw):
        pass
    print(Counter(re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw)).most_common(10))
    print(Counter(re.findall(r'sz="(\d+)"', raw)).most_common(10))
    print(Counter(re.findall(r'typeface="([^"]+)"', raw)).most_common(6))
