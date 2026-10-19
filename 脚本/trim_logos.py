# -*- coding: utf-8 -*-
"""Trim transparent margins of rasterized brand marks -> work_e/logo_final."""
import os

from PIL import Image

SRC = "work_e/logo_png"
OUT = "work_e/logo_final"
os.makedirs(OUT, exist_ok=True)

for name in sorted(os.listdir(SRC)):
    if not name.endswith(".png"):
        continue
    key = name[:-4]
    im = Image.open(os.path.join(SRC, name)).convert("RGBA")
    bbox = im.split()[3].getbbox()
    if bbox:
        im = im.crop(bbox)
    # pad to square canvas with 6% margin so badges look even
    side = max(im.size)
    pad = int(side * 0.06)
    canvas = Image.new("RGBA", (side + 2 * pad, side + 2 * pad), (0, 0, 0, 0))
    canvas.alpha_composite(im, ((side + 2 * pad - im.width) // 2,
                                (side + 2 * pad - im.height) // 2))
    canvas = canvas.resize((512, 512), Image.LANCZOS)
    canvas.save(os.path.join(OUT, key + ".png"))
    print(key, im.size, "->512")
