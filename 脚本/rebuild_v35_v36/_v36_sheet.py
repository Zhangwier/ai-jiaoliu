# -*- coding: utf-8 -*-
"""生成 2x2 联页检查图：python sheet.py r15 1"""
import sys, os, glob
from PIL import Image, ImageDraw

d = os.path.join(r"D:\MG\program\ai交流\_v36", sys.argv[1])
outdir = os.path.join(d, "sheet")
os.makedirs(outdir, exist_ok=True)
files = sorted(glob.glob(os.path.join(d, "page-*.png")))
W, H = 733, 412
for i in range(0, len(files), 4):
    grp = files[i:i + 4]
    canvas = Image.new("RGB", (W * 2 + 12, H * 2 + 12), (60, 60, 60))
    for j, f in enumerate(grp):
        im = Image.open(f).convert("RGB").resize((W, H), Image.LANCZOS)
        x = (j % 2) * (W + 12)
        y = (j // 2) * (H + 12)
        canvas.paste(im, (x, y))
        dr = ImageDraw.Draw(canvas)
        dr.text((x + 6, y + 5), os.path.basename(f), fill=(255, 60, 60))
    n = i // 4 + 1
    canvas.save(os.path.join(outdir, "s%02d.png" % n))
print("sheets", (len(files) + 3) // 4)
