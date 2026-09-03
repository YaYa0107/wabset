import zipfile, re, json
from xml.etree import ElementTree as ET

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'p': P, 'a': A}
z = zipfile.ZipFile(SRC)

names = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
               key=lambda x: int(re.search(r'(\d+)', x.split('/')[-1]).group(1)))

for idx, sn in enumerate(names, 1):
    raw = z.read(sn).decode('utf-8')
    root = ET.fromstring(raw)
    timing = root.find('.//p:timing/p:tnLst', NS)
    info = []
    if timing is not None:
        # build id->shape text map
        idmap = {}
        for sp in root.iter('{%s}sp' % P):
            cNv = sp.find('.//p:nvSpPr/p:cNvPr', NS)
            if cNv is None:
                cNv = sp.find('.//p:cNvPr', NS)
            txt = ' '.join(''.join(t.text or '' for t in p.iter('{%s}t' % A))
                           for p in sp.iter('{%s}p' % A)).strip()
            if cNv is not None:
                idmap[cNv.attrib.get('id')] = (cNv.attrib.get('name', ''), txt[:40])
        for par in timing.iter('{%s}par' % P):
            cTn = par.find('p:cTn', NS)
            # find innermost
            for inner in par.iter('{%s}cTn' % P):
                pass
            # iterate child animEffect / animMotion / set
            for el in par.iter():
                tag = el.tag.split('}')[-1]
                if tag in ('animEffect', 'animMotion', 'animRot', 'animScale'):
                    tr = el.get('transition') or ''
                    # target
                    tgt = ''
                    for c in par.iter():
                        if c.tag.endswith('}spTgt'):
                            tgt = c.attrib.get('spid')
                    info.append((tag, tr, tgt, idmap.get(tgt, ('', ''))[1] or idmap.get(tgt, ('', ''))[0]))
    if info:
        print('SLIDE', idx)
        for t in info:
            print('   ', t)
