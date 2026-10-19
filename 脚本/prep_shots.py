# -*- coding: utf-8 -*-
"""Redact the two magpie screenshots for the D2 pages.

Reads the untouched originals in the repo root, writes assets/*.png.
Idempotent: always starts from the original, never from a previous output.
"""
import os
from PIL import Image, ImageDraw, ImageFont

MS = "C:/Windows/Fonts/msyh.ttc"
GRAY = (118, 126, 140)
DARK = (36, 40, 50)
CHIP = (249, 249, 251)
WHITE = (255, 255, 255)

os.makedirs("assets", exist_ok=True)


def f(size):
    return ImageFont.truetype(MS, size)


# ------------------------------------------------------------------ 1. account list
def accounts():
    im = Image.open("shot_agents.png").convert("RGB")
    d = ImageDraw.Draw(im)
    seps = [12, 102, 192, 281, 371, 461, 551, 640, 730, 820, 910, 1006]
    models = [1, 3, 3, 3, 1, 1, 1, 1, 1, 2, 1]
    for i in range(11):
        top, bot = seps[i], seps[i + 1]
        d.rectangle([96, top + 48, 1560, bot - 2], fill=WHITE)
        d.text((104, top + 60), "已登录 账号 A%02d · %d 个模型" % (i + 1, models[i]),
               font=f(20), fill=GRAY)
    # API-key tail on the MP row
    d.rectangle([1792, 936, 1992, 992], fill=WHITE)
    d.text((1797, 950), "sk-F••••BW8", font=f(19), fill=GRAY)
    out = im.crop((8, 4, 2189, 1006))
    out.save("assets/shot_accounts.png", optimize=True)
    return out.size


# ------------------------------------------------------------------ 2. routing graph
def route():
    im = Image.open("shot_route.png").convert("RGB")
    d = ImageDraw.Draw(im)
    px = im.load()

    # find every text band in the name column (x 700..1000)
    bands, cur = [], None
    for y in range(im.size[1]):
        dark = any(sum(px[x, y]) < 700 for x in range(700, 1000))
        if dark and cur is None:
            cur = y
        if not dark and cur is not None:
            if y - cur >= 8:
                bands.append((cur, y))
            cur = None
    # first band is the group header ("WorkBuddy AI 5 个账号") — keep it
    bands = bands[1:]

    # the boxed group row (the one that also carries the live summary) keeps
    # its "WorkBuddy 5 个账号" label; only its numeric id gets masked
    group_row = None
    for y0, y1 in bands:
        if 690 <= y0 <= 715:
            group_row = (y0, y1)

    n = 0
    for y0, y1 in bands:
        if group_row == (y0, y1):
            continue                      # handled below
        n += 1
        d.rectangle([700, y0 - 8, 1424, y1 + 3], fill=WHITE)
        d.text((716, y0 - 2), "账号 A%02d" % n, font=f(22), fill=GRAY)
        d.rounded_rectangle([1130, y0 - 10, 1378, y1 + 8], radius=12, fill=CHIP)
        d.text((1148, y0 - 6), "deepseek-v4.1-flash", font=f(19), fill=(58, 62, 72))

    # left rail: wipe the whole node (it carries a group name + an account label)
    d.rectangle([4, 570, 250, 780], fill=WHITE)
    d.text((14, 592), "本地网关", font=f(20), fill=GRAY)
    d.text((14, 624), "路由组 auto", font=f(18), fill=GRAY)
    d.text((14, 656), "group/auto", font=f(17), fill=GRAY)
    d.text((14, 700), "下面这一列都是它", font=f(15), fill=GRAY)
    d.text((14, 724), "可以调用的账号", font=f(15), fill=GRAY)

    # active-row summary started with a numeric account id
    d.rectangle([995, 672, 2085, 740], fill=WHITE)
    d.text((1052, 688), "本次由账号 A10 回答 · 3 个可用 · 2 个休息中", font=f(18), fill=DARK)

    out = im.crop((0, 20, 2103, 1245))
    out.save("assets/shot_route.png", optimize=True)
    out.save("assets/shot_route_slide.png", optimize=True)
    return out.size


if __name__ == "__main__":
    print("accounts", accounts())
    print("route   ", route())
