# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs(); s=prs.slides[108]
FIX = {
 "并发与吞吐最高，服务器上跑开源权重的默认选择": "并发吞吐最高，服务器上跑开源权重首选",
 "表格与扫描件解析强，检索质量靠它兜底": "表格与扫描件解析强，检索靠它兜底",
 "每次调用的输入输出、耗时、成本都留痕": "每次调用都留输入输出、耗时与成本",
}
for sh in s.shapes:
    if not sh.has_text_frame: continue
    t=sh.text_frame.text.strip()
    if t in FIX:
        settext(sh, FIX[t], size=11); print("fix:", FIX[t])
    if sh.name=="TextBox 7":
        settext(sh, "来源：各项目 GitHub README；2026-10-20 查询；本稿为选型清单，未做部署实测", size=9)
        sh.width=Cm(25.4); print("src short")
save(prs)
