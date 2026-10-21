# -*- coding: utf-8 -*-
"""Rasterize brand SVGs (simple-icons) into transparent PNGs via Playwright/Edge."""
import os
from playwright.sync_api import sync_playwright

SVGS = {
    "openai": ("openai", "000000"),
    "anthropic": ("anthropic", "000000"),
    "gemini": ("googlegemini", "000000"),
    "grok": ("x", "000000"),
    "meta": ("meta", "0064E0"),
    "mistral": ("mistral", None),
    "deepseek": ("deepseek", "4D6BFE"),
    "qwen": ("qwen", "615CED"),
    "glm": ("zhipu", None),
    "kimi": ("kimi", "000000"),
    "minimax": ("minimax", "E73562"),
    "hy": ("tencenthy", "0052D9"),
    "baidu": ("baidu", "2932E1"),
    "dots": ("xiaohongshu", "FF2442"),
}

OUT = "work_e/logo_png"
os.makedirs(OUT, exist_ok=True)

html = '<html><body style="margin:0;background:#fff">'
used = []
for key, (slug, color) in SVGS.items():
    path = os.path.join("work_e/logo", slug + ".svg")
    if not os.path.exists(path):
        print("missing svg", key, slug)
        continue
    svg = open(path, encoding="utf-8").read()
    attrs = 'width="512" height="512"'
    if color:
        attrs += ' fill="#%s"' % color
    svg = svg.replace("<svg ", "<svg " + attrs + " ", 1)
    html += '<div id="%s" style="width:512px;height:512px">%s</div>' % (key, svg)
    used.append(key)
html += "</body></html>"

hp = os.path.abspath("work_e/logo/render.html")
open(hp, "w", encoding="utf-8").write(html)

with sync_playwright() as pw:
    b = pw.chromium.launch(channel="msedge")
    pg = b.new_page(viewport={"width": 600, "height": 600})
    pg.goto("file:///" + hp.replace("\\", "/"))
    for key in used:
        el = pg.query_selector("#" + key)
        if not el:
            print("no el", key)
            continue
        el.screenshot(path=os.path.join(OUT, key + ".png"), omit_background=True)
        print("shot", key)
    b.close()
