# -*- coding: utf-8 -*-
"""step33: 编号改 B1c/B2c、DeepSeek 改公司介绍、D3 改个人推荐、降智页加序号"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *

prs = openprs()
S = lambda n: prs.slides[n-1]
def sh(slide, name):
    r = find(slide, name=name)
    if r is None: print("!! not found", name)
    return r

# ---------- s60 降智 → B1c ----------
s = S(60)
settext(sh(s,"TextBox 4"), "B1c  “降智”：官方承认过 bug，没承认按负载降智", size=26)
settext(sh(s,"TextBox 5"), "先别急着怀疑模型：真 bug、被换模型、上下文太长、单次抽卡，都会让体感变差", size=17)
settext(sh(s,"TextBox 12"), "① 先查自己", size=17)
settext(sh(s,"TextBox 16"), "② 对不对得上指纹", size=17)
settext(sh(s,"TextBox 20"), "③ 再看长期曲线", size=17)
settext(sh(s,"TextBox 25"), "三步走完再决定换不换：只凭一次体感，换哪家都会觉得“它又变笨了”", size=18)

# ---------- s63 封号 → B2c ----------
s = S(63)
settext(sh(s,"TextBox 4"), "B2c  Anthropic 的限制：中资实体、地区与封号", size=26)
settext(sh(s,"TextBox 5"), "按时间顺序；数字来自官方公告和报道，2026-10-18 整理；这一页会直接决定你能不能长期用", size=17)
settext(sh(s,"TextBox 32"), "三条红线：中资控股、地区名单、借卡借号——采购和自用都先按这条排除", size=18)

# ---------- s68 DeepSeek 改公司介绍 ----------
s = S(68)
settext(sh(s,"TextBox 4"), "B7  DeepSeek：把开源当方法、把低价当杠杆", size=26)
settext(sh(s,"TextBox 5"), "不只放权重：报告、算子、代码一起开；Flash 1017 的能力已贴到 5.6 Sol 那一档", size=17)
settext(sh(s,"TextBox 13"), "公司 / 模型线", size=11.5)
settext(sh(s,"TextBox 14"),
    "DeepSeek（深度求索）2023 年由幻方量化孵化，一路把“低价 + 开源”做成标签：V2 打价格，V3 / R1 打开源与推理，V4 系列打长上下文。当前主力是 V4.1 Flash 1017（2026-10-17，1M 上下文）。", size=11.5)
settext(sh(s,"TextBox 16"), "产品与接入", size=11.5)
settext(sh(s,"TextBox 17"),
    "网页 / App；API 完全兼容 OpenAI（改个 base_url 就能接）；第三方客户端接入最广（WorkBuddy、Cherry Studio、Cursor 类都支持）；权重可自建、可私有化——这是它能进企业内网的前提。", size=11.5)
settext(sh(s,"TextBox 19"), "强项", size=11.5)
settext(sh(s,"TextBox 20"),
    "① 开源不只放权重：技术报告写清架构与训练策略，开源周放出 FlashMLA、DeepEP、DeepGEMM 等底层算子，推理与评测代码也公开（细节见 B7b）。② 价格与速度仍是最低一档：输出 8 元每百万 token，闲时半价，约 217 token/s。③ Flash 1017 已贴到 5.6 Sol 那一档：官方自报 DeepSWE 77.0，第三方复测还没出。", size=11.5)
settext(sh(s,"TextBox 22"), "短板 / 边界", size=11.5)
settext(sh(s,"TextBox 23"),
    "识图与图内文字弱（图纸 0/2）；提示词写得粗就明显掉档；工作日高峰限速；海外合规与企业治理弱于 OpenAI / Anthropic；1017 还没拿到 AA 分，对比它只能先看自报。", size=11.5)
settext(sh(s,"TextBox 26"),
    "鹈鹕 C（V4.1 Flash 1M · WorkBuddy，P19）；图纸 0/2（P29）；修编报告 34 分钟（P33）；价格（P22）；开源四层（B7b）。", size=11.5)
settext(sh(s,"TextBox 30"), "要开源可控、要便宜、要大批量跑时首选；识图与强合规场景别指望它", size=17)
settext(sh(s,"TextBox 7"),
    "来源：api-docs.deepseek.com 定价与 1017 发布说明；huggingface.co/deepseek-ai；deepseeklicense.github.io；2026-10-18 查询", size=10.5)

# ---------- s86 D3 个人推荐 ----------
s = S(86)
settext(sh(s,"TextBox 4"), "D3  B 站：我个人常看的两位 UP 主", size=26)
settext(sh(s,"TextBox 5"), "不是行业结论，只是我个人常看的两类内容：一个看真实项目实测，一个看每日资讯", size=17)
settext(sh(s,"TextBox 27"), "实测看“词元”，动态看“橘鸦”；这只是我的习惯，结论仍以原始出处和自己的测试为准", size=18)

save(prs)
print("saved")
