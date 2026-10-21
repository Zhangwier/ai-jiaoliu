# -*- coding: utf-8 -*-
import re
import urllib.request

HDR = {"User-Agent": "Mozilla/5.0"}
for u in ["https://z.ai/", "https://chat.z.ai/", "https://docs.z.ai/", "https://www.zhipuai.cn/"]:
    try:
        r = urllib.request.urlopen(urllib.request.Request(u, headers=HDR), timeout=25)
        h = r.read(60000).decode("utf-8", "ignore")
        print(u, r.status, re.findall(r'href="([^"]*favicon[^"]*)"', h)[:5],
              re.findall(r'href="([^"]*\.svg[^"]*)"', h)[:5])
    except Exception as e:
        print(u, "FAIL", type(e).__name__, str(e)[:80])
