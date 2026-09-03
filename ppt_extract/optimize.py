from PIL import Image
import os, glob

IMG = r"C:\Users\lyq15\WorkBuddy\2026-09-03-11-20-34\website\assets\img"
KEEP_BIG = {'image1.jpeg', 'image2.jpeg', 'image43.jpeg'}  # 背景图保留 1600px

total_before = total_after = 0
for path in glob.glob(os.path.join(IMG, '*.*')):
    base = os.path.basename(path)
    total_before += os.path.getsize(path)
    im = Image.open(path)
    big = base in KEEP_BIG
    maxw = 1600 if big else 1100
    if im.width > maxw:
        im = im.resize((maxw, round(im.height * maxw / im.width)), Image.LANCZOS)
    im = im.convert('RGB')
    out = os.path.splitext(path)[0] + '.jpg'
    q = 85 if big else 80
    im.save(out, 'JPEG', quality=q, optimize=True, progressive=True)
    if out != path:
        os.remove(path)
    total_after += os.path.getsize(out)
    print(f"{base}: {im.width}x{im.height}")

print(f"total: {total_before/1048576:.1f} MB -> {total_after/1048576:.1f} MB")
