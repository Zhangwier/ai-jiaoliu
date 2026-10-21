# -*- coding: utf-8 -*-
"""Rasterize arbitrary brand SVGs (Z.ai, Mistral) to transparent PNG via Edge."""
import os
import urllib.request

from playwright.sync_api import sync_playwright

HDR = {"User-Agent": "Mozilla/5.0"}
GET = {
    "glm": "https://z-cdn.chatglm.cn/z-ai/static/logo.svg",
    "mistral": "https://mistral.ai/favicon.svg",
}
os.makedirs("work_e/logo", exist_ok=True)
os.makedirs("work_e/logo_png", exist_ok=True)

for key, url in GET.items():
    d = urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=30).read()
    open("work_e/logo/%s.svg" % key, "wb").write(d)
    print(key, len(d), d[:120])

html = '<html><body style="margin:0;background:#fff">'
for key in GET:
    svg = open("work_e/logo/%s.svg" % key, encoding="utf-8").read()
    svg = svg.replace("<svg ", '<svg width="512" height="512" style="max-width:512px;max-height:512px" ', 1)
    html += '<div id="%s" style="width:512px;height:512px;display:flex;align-items:center;justify-content:center">%s</div>' % (key, svg)
html += "</body></html>"

hp = os.path.abspath("work_e/logo/render2.html")
open(hp, "w", encoding="utf-8").write(html)

with sync_playwright() as pw:
    b = pw.chromium.launch(channel="msedge")
    pg = b.new_page(viewport={"width": 600, "height": 600})
    pg.goto("file:///" + hp.replace("\\", "/"))
    for key in GET:
        el = pg.query_selector("#" + key)
        el.screenshot(path="work_e/logo_png/%s.png" % key, omit_background=True)
        print("shot", key)
    b.close()
