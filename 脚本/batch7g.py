import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[70]
EMU=360000
for sh in s.shapes:
    if sh.has_text_frame:
        t=sh.text_frame.text
        if t.startswith("① 开源不只放权重"):
            print("强项", sh.top/EMU, sh.height/EMU, sh.width/EMU)
            settext(sh, "① 开源不只放权重：技术报告写清架构与训练策略，开源周放出 FlashMLA、DeepEP、DeepGEMM 等底层算子，评测代码也公开（见 B7b）。② 价格与速度仍是最低一档：输出 8 元/百万 token、闲时半价，约 217 token/s。③ Flash 1017 已贴到 5.6 Sol 那一档：官方自报 DeepSWE 77.0，第三方复测尚未出。", size=11.5)
        if t.startswith("识图与图内文字弱"):
            print("短板", sh.top/EMU, sh.height/EMU)
            settext(sh, "识图与图内文字弱（图纸 0/2）；提示词写得粗就明显掉档；工作日高峰限速；海外合规与企业治理弱于 OpenAI / Anthropic。", size=11.5)
save(prs); print("ok")
