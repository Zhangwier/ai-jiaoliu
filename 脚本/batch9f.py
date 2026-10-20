# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs()

SUB = {
 54:  "模型本身 · 怎么交代任务 · Agent 与工具 · 常见误解",
 59:  "模型线 · 可用入口 · 强项 · 边界；OpenAI、Anthropic、Meta（Muse）、DeepSeek、Qwen 另各有专题页；降智与封号紧跟在对应公司之后",
 80:  "四件事自查 · 照着做七步 · 地区与支付 · 注册 · 价格 · 省钱的路子 · 省流版；价格与地区政策以官方页面为准",
 88:  "社区看风向 · GitHub 看版本 · B 站看演示 · X 看一手",
 93:  "降智测试（附两个自查工具）· 中转站与代理池 · 护栏与破限 · 开源落地",
 110: "发展史 · 神经网络 · token 与向量 · Transformer 与注意力 · 架构 · 规模定律 · 对齐与推理 · 多模态",
 121: "按编号归档；另有可直接点开的原链接清单，正文页脚都指到这里",
}
for n,txt in SUB.items():
    s=prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/EMU<22 and 12.0 < sh.top/EMU < 14.6 and sh.height/EMU>1.0:
            settext(sh, txt, size=14.5)
# 统一 D / Z 分隔页标题
for n,want in ((88,"附录\n资讯渠道"),(121,"附录\n参考资料与来源")):
    s=prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/EMU<22 and 8<sh.top/EMU<9.6 and sh.height/EMU>2.5:
            settext(sh, want, size=32)
            print(n,"title ->",want.replace("\n","/"))
# s52 索引行去掉页范围
s=prs.slides[51]
for sh in s.shapes:
    if sh.has_text_frame and "附录" in sh.text_frame.text and "A1" in sh.text_frame.text:
        settext(sh,"附录  常用术语 · 公司档案 · 订阅与省流版 · 资讯渠道 · 分享与工具 · 原理与历史 · 参考资料来源", size=14)
        print("s52 simplified")
save(prs)
