"""校验：网页里用到的图片 = PPT 各页真实存在的图片，且顺序与坐标一致。"""
import zipfile, re, os
from xml.etree import ElementTree as ET

SRC = r"C:\Users\lyq15\Desktop\网站任务\网站部署全流程解析.pptx"
SITE = r"C:\Users\lyq15\WorkBuddy\2026-09-03-11-20-34\website"
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': P, 'a': A}
z = zipfile.ZipFile(SRC)

# 1) PPT 每页真实图片（按 y, x 坐标排序）
truth = {}
for i in range(1, 31):
    sn = 'ppt/slides/slide%d.xml' % i
    if sn not in z.namelist():
        continue
    root = ET.fromstring(z.read(sn))
    relmap = {}
    rp = 'ppt/slides/_rels/slide%d.xml.rels' % i
    if rp in z.namelist():
        for rel in ET.fromstring(z.read(rp)):
            relmap[rel.attrib['Id']] = rel.attrib['Target'].split('/')[-1]
    items = []
    for pic in root.iter('{%s}pic' % P):
        blip = pic.find('.//a:blip', NS)
        rid = blip.attrib.get('{%s}embed' % R) if blip is not None else None
        img = relmap.get(rid, '?')
        xf = pic.find('.//a:xfrm', NS)
        if xf is None:
            continue
        off, ext = xf.find('a:off', NS), xf.find('a:ext', NS)
        items.append((int(off.attrib['y']), int(off.attrib['x']), img))
    # 几何判断主方向：x 跨度 > y 跨度 → 层叠/并排（按 x 排）；否则上下排列（按 y 排）
    if len(items) > 1:
        xs = [t[1] for t in items]; ys = [t[0] for t in items]
        span_x, span_y = max(xs) - min(xs), max(ys) - min(ys)
        items.sort(key=lambda t: t[1] if span_x >= span_y else t[0])
    truth[i] = [t[2] for t in items]

# 2) 网页实际引用：解析 HTML + JS 图集数组
html = open(os.path.join(SITE, 'index.html'), encoding='utf-8').read()

# JS 图集：  14:['image16','image17',...]
gal = {}
for m in re.finditer(r"(\d+):\[([^\]]+)\]", html):
    page = int(m.group(1))
    gal[page] = re.findall(r"'(image\d+)'", m.group(2))

# HTML 中的静态引用：assets/img/imageN.jpg，并找它前面最近的 P 标注
used = {}   # page -> [images]
body = html
for m in re.finditer(r'assets/img/(image\d+)\.jpg', body):
    name = m.group(1)
    # 图片标注写在 img 之后的 <figcaption> 里 → 向后查找
    ctx = body[m.end():m.end() + 400]
    pm = re.search(r'P(\d+)-(\d+)', ctx)
    if pm:
        page = int(pm.group(1))
        used.setdefault(page, [])
        if name not in used[page]:
            used[page].append(name)

for p in gal:
    used[p] = gal[p]

print('%-6s %-46s %s' % ('页码', 'PPT 真实（按坐标排序）', '网页使用'))
ok = True
for i in range(1, 31):
    t = truth.get(i, [])
    u = used.get(i, [])
    tn = [x.rsplit('.', 1)[0] for x in t]
    flag = ''
    if tn != u:
        flag = '  ❌ 不一致'
        ok = False
    print('%-6s %-46s %s%s' % ('P' + str(i), ','.join(tn) or '（无图）', ','.join(u) or '（无图）', flag))

# 3) 遗漏检查：PPT 有但网页没用
allt = set(x.rsplit('.', 1)[0] for v in truth.values() for x in v)
allu = set(x for v in used.values() for x in v)
print('\nPPT 有的图:', len(allt), '| 网页用了:', len(allu))
print('PPT 有、网页未用:', sorted(allt - allu))
print('网页用、PPT 没有:', sorted(allu - allt) or '（无）')

# 4) 文件存在性
missing = [f for f in allu if not os.path.exists(os.path.join(SITE, 'assets', 'img', f + '.jpg'))]
print('缺失文件:', missing or '（无）')
print('\n结果:', '✅ 全部一致' if ok and not missing else '⚠️ 存在差异')
