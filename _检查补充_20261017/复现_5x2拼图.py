# -*- coding: utf-8 -*-
import os, glob
from PIL import Image, ImageDraw, ImageFont

OUT = r"D:\MG\program\ai交流\_检查补充_20261017"
SHOT = os.path.join(OUT, "鹈鹕统一视口截图")
codes = list("ABCDEFGHIJ")
cell_w, cell_h, pad, label_h = 480, 360, 12, 26
W = 5 * (cell_w + pad) + pad
H = 2 * (cell_h + label_h + pad) + pad

sheet = Image.new("RGB", (W, H), (255, 255, 255))
d = ImageDraw.Draw(sheet)
font = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 22)

for i, c in enumerate(codes):
    hits = glob.glob(os.path.join(SHOT, f"{c}_*_1200x900.png"))
    if not hits:
        print(c, "MISS"); continue
    im = Image.open(hits[0]).convert("RGB")
    im.thumbnail((cell_w, cell_h))
    col, row = i % 5, i // 5
    x = pad + col * (cell_w + pad)
    y = pad + row * (cell_h + label_h + pad)
    sheet.paste(im, (x, y + label_h))
    d.rectangle([x, y, x + cell_w, y + cell_h + label_h], outline=(190, 196, 205), width=2)
    d.text((x + 8, y + 2), c, fill=(16, 34, 60), font=font)

p = os.path.join(OUT, "鹈鹕盲测_5x2拼图.png")
sheet.save(p)
print(p, sheet.size, round(os.path.getsize(p) / 1024), "KB")
