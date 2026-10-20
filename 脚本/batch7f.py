import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[61]
# 修正三卡标题与内容错位
m = [
 ("TextBox 12", "① 先查自己", 17),
 ("TextBox 13", "对话是不是太长、换了客户端或档位、只试了一次；先自己排除这三样。", 14),
 ("TextBox 16", "② 对不对得上指纹", 17),
 ("TextBox 17", "lm.ikale.io 用“随机数指纹”猜背后是哪个模型；结果只是参考，Key 会经过它的代理。", 14),
 ("TextBox 20", "③ 再看长期曲线", 17),
 ("TextBox 21", "Margin Lab 每天跑 50 题：单日波动 ±12.8 个点内都算正常，要看 7 日、30 日。", 14),
]
for sh in s.shapes:
    for nm, tx, sz in m:
        if sh.name == nm:
            settext(sh, tx, size=sz)
# 副标覆盖两家
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("先别急着怀疑模型"):
        settext(sh, "OpenAI 与 Anthropic 都出过：真 bug、被换模型、上下文太长、单次抽卡，都会让体感变差", size=17)
# 结论条强化提醒
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("三步走完再决定换不换"):
        settext(sh, "三步走完再决定换不换；重要交付别押在一次生成上——体感会骗人，核过的结果不会", size=18)
save(prs); print("ok")
