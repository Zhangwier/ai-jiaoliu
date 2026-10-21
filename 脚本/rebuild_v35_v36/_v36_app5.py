# -*- coding: utf-8 -*-
"""附录 F：原理与历史 + 附录 Z：来源"""
from ds import *
from app import sec, CH, hd
from app2 import lab, para, kv, BODY
from app3 import CH as _c
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ---------------- F1 发展史 ----------------
def f1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "从 1943 到今天：数据、算力、架构三条线汇到一起",
           "今天的大模型，就是 2017 年 Transformer 架构加上海量数据与算力。中间有过两次“寒冬”，也都过去了。")
    # 左：感知机照片
    PW = 2.54
    frame(s, ML, y + 0.04, PW, 2.14, LINE, 1.0, CARD)
    pic(s, "s111_1.png", ML + 0.05, y + 0.09, PW - 0.10, 2.04, mode="fill")
    text(s, ML, y + 2.24, PW, 0.42,
         "Mark I 感知机（康奈尔航空实验室），操作手册插图，1960 年。图：Hay & Murray，公有领域",
         8, GREY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    # 右：时间轴
    TX = ML + PW + 0.52
    TW = X2 - TX
    axis_y = y + 1.34
    hr(s, TX, axis_y, TW, 1.4, LINE2)
    N = [("1943", "数学神经元", "写成“加权求和、过阈值就输出”"),
         ("1958", "感知机", "Rosenblatt：机器第一次自己调权重"),
         ("2012", "深度学习", "AlexNet：ImageNet 错误率 26% → 15%"),
         ("2017", "Transformer", "整句并行处理，规模才做得大"),
         ("2022 起", "对话产品", "ChatGPT 2022 年 11 月上线"),
         ("现在", "多模态与工具", "看图、读文件、调工具，动手做事")]
    step = TW / 5
    for i, (yr, nm, ds) in enumerate(N):
        cx = TX + i * step
        circ(s, cx, axis_y, 0.115, TEAL)
        text(s, cx - 0.06, axis_y - 0.40, 1.30, 0.34, yr, 17, TEAL if i < 4 else ORANGE, MONO,
             True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cx - 0.06, axis_y + 0.14, step - 0.14, 0.26, nm, 11.5, NAVY, SANS, True,
             PP_ALIGN.LEFT, 1.1, check=False)
        nl = nlines(ds, 8.5, step - 0.16)
        text(s, cx - 0.06, axis_y + 0.40, step - 0.16, nl * 8.5 * LINES * 1.3 / 72 + 0.05, ds, 8.5,
             GREY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    # 低谷与转机
    BY = y + 2.62
    hr(s, ML, BY - 0.16, CW, 1.0, LINE)
    BX3 = [("中间的低谷", "1969 年《Perceptrons》指出单层网络学不会“异或”，研究冷了十几年", ORANGEL, ORANGE),
           ("转机：反向传播", "1986 年反向传播让多层网络能训练（见 F2）", MINT, TEAL),
           ("三条线汇合", "数据（互联网文本与图像）＋算力（GPU 集群）＋架构（感知机 → Transformer）", TEALXL, TEAL)]
    bw = (CW - 0.30 * 2) / 3
    for i, (t, b, fill, col) in enumerate(BX3):
        x = ML + i * (bw + 0.30)
        box(s, x, BY, bw, 0.90, fill)
        text(s, x + 0.20, BY + 0.12, bw - 0.40, 0.24, t, 9.5, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=0.6, check=False)
        nl = nlines(b, 9.5, bw - 0.40)
        text(s, x + 0.20, BY + 0.36, bw - 0.40, nl * 9.5 * LINES * 1.3 / 72 + 0.05, b, 9.5, BODY,
             SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：Krizhevsky 等 2012；Minsky & Papert 1969；Rumelhart 等 1986；Vaswani 等 2017；图 Wikimedia Commons（公有领域）")
    return s


# ---------------- F2 学习四步 ----------------
def f2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "神经网络怎么学：猜一次、量误差、往误差小的方向挪一点",
           "训练里没有“理解”这一步，只是让损失一点点变小。学到什么，取决于数据里有什么。")
    # 左：四步
    LW = 6.00
    ST = [("01", "猜", "用当前的 w 算出预测，比如预测用量 8.2"),
          ("02", "量", "和真实答案 9.0 比：损失 =（预测 − 实际）²，错得越多越大"),
          ("03", "求方向", "对 w 求导（梯度）：知道 w 往大挪还是往小挪，损失会变小"),
          ("04", "挪一小步", "w ← w − 学习率 × 梯度；换一批数据，回到 01，重复亿万次")]
    yy = y + 0.04
    for no, t, b in ST:
        hr(s, ML, yy, LW, 1.0, LINE)
        text(s, ML, yy + 0.14, 0.42, 0.26, no, 11, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.52, yy + 0.11, 1.20, 0.30, t, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        nl = nlines(b, 11, LW - 1.78)
        text(s, ML + 1.78, yy + 0.13, LW - 1.78, nl * 11 * LINES * 1.35 / 72 + 0.05, b, 11, BODY,
             SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
        yy += 0.42 + nl * 11 * LINES * 1.35 / 72 + 0.14
    hr(s, ML, yy, LW, 1.0, LINE)
    text(s, ML, yy + 0.16, LW, 0.60,
         "参数有千亿个时怎么办：反向传播用链式法则，从输出往回逐层算出每个参数该挪多少，一次算完。",
         11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.35, check=False)
    # 右：损失曲面图
    RX = ML + LW + 0.52
    RW = X2 - RX
    vr(s, RX - 0.26, y + 0.04, 3.60, 1.0, LINE)
    frame(s, RX, y + 0.06, RW, 2.90, LINE, 1.0, CARD)
    pic(s, "s112_1.png", RX + 0.06, y + 0.12, RW - 0.12, 2.78, mode="fill")
    text(s, RX, y + 3.02, RW, 0.46,
         "只有一个参数 w 时是这样；真模型是这件事在千亿维空间里做同一件事。图为示意。", 8.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：Rumelhart、Hinton & Williams（Nature，1986）；图为示意；学习率、损失与参数均为示例")
    return s


# ---------------- F3 token 与向量 ----------------
def f3(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "词怎么变成数字：token 与向量",
           "RAG 的“语义检索”、上下文的“token 上限”、按 token 计价，都建立在这一步上。模型不认字，只认编号。")
    # 上：三步
    LW = 6.30
    yy = y + 0.10
    text(s, ML, yy, 1.10, 0.22, "① 原文", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    text(s, ML + 0.90, yy - 0.04, 4.60, 0.30, "桥墩挡块需复核", 15, NAVY, SANS, True, PP_ALIGN.LEFT,
         1.1, check=False)
    yy += 0.44
    text(s, ML, yy, 1.10, 0.22, "② 切 token", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    TK = ["桥墩", "挡", "块", "需", "复核"]
    for i, t in enumerate(TK):
        bx = ML + 0.90 + i * 0.92
        box(s, bx, yy - 0.06, 0.80, 0.36, MINT)
        text(s, bx, yy - 0.005, 0.80, 0.28, t, 11, NAVY, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
    yy += 0.52
    text(s, ML, yy, 1.10, 0.22, "③ 查编号", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    IDS = ["50213", "8812", "3307", "1290", "41766"]
    for i, t in enumerate(IDS):
        bx = ML + 0.90 + i * 0.92
        text(s, bx, yy - 0.02, 0.84, 0.26, t, 10.5, TEAL, MONO, True, PP_ALIGN.CENTER, 1.0,
             check=False)
    yy += 0.40
    text(s, ML, yy, 1.10, 0.22, "④ 换向量", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    text(s, ML + 0.90, yy - 0.04, 5.20, 0.30, "[0.12, −0.83, 0.40, …]", 10.5, TEAL, MONO, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 0.90, yy + 0.28, 5.30, 0.24, "每个 token 一串上千个数；编号与数值仅为示意", 8.5, FAINT,
         SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    TY = yy + 0.66
    hr(s, ML, TY, LW, 1.0, LINE2)
    text(s, ML, TY + 0.14, LW, 0.28, "国王 − 男人 ＋ 女人 ≈ 女王", 15, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.1, check=False)
    text(s, ML, TY + 0.46, LW, 0.52,
         "向量做加减也有意义（word2vec）。这也是“意思相近”能被算出来的原因 —— 检索靠的正是这个距离。",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    # 右：嵌入散点
    RX = ML + LW + 0.50
    RW = X2 - RX
    vr(s, RX - 0.24, y + 0.10, 3.30, 1.0, LINE)
    frame(s, RX, y + 0.10, RW, 2.62, LINE, 1.0, CARD)
    pic(s, "s114_1.png", RX + 0.05, y + 0.15, RW - 0.10, 2.52, mode="fill")
    text(s, RX, y + 2.82, RW, 0.66,
         "同一批工程术语的向量分布：防护构件、合同条款、材料各聚成一簇。模型没见过你的项目，也能靠“近”找到同类。",
         8.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：Sennrich 等 2016（BPE 分词）；Mikolov 等 2013（word2vec）；切分、编号与向量均为示意；图为本地模型二维投影")
    return s


# ---------------- F4 Transformer 一层 ----------------
def f4(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "Transformer 一层：注意力和前馈，再叠几十上百层",
           "所谓“大模型”，主要是这一层更宽、叠得更多、喂的数据更多 —— 结构本身没变。")
    # 左：结构图
    LX = ML
    LW = 4.30
    bw = 2.70
    cx = LX + (LW - bw) / 2
    yy = y + 0.06
    text(s, cx, yy, bw, 0.26, "输入：token 向量 ＋ 位置信息", 10, FAINT, SANS, False, PP_ALIGN.CENTER,
         1.0, check=False)
    yy += 0.32
    # 主块
    def blk(y0, t, sub, fill, h=0.52):
        box(s, cx, y0, bw, h, fill, TEAL)
        text(s, cx, y0 + 0.05, bw, 0.26, t, 11.5, NAVY, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
        if sub:
            text(s, cx, y0 + 0.27, bw, 0.22, sub, 8.5, GREY, SANS, False, PP_ALIGN.CENTER, 1.0,
                 check=False)
    blk(yy, "多头注意力", "看别的词", MINT, 0.50)
    yy += 0.60
    blk(yy, "相加 & 归一化", "残差直通", TEALXL, 0.44)
    yy += 0.54
    blk(yy, "前馈网络", "在自己身上加工；参数大头", MINT, 0.50)
    yy += 0.60
    blk(yy, "相加 & 归一化", "残差直通", TEALXL, 0.44)
    yy += 0.54
    # 残差箭头
    line(s, cx - 0.12, y + 0.32, cx - 0.42, y + 0.32, LINE2, 1.0)
    vr(s, cx - 0.42, y + 0.32, 2.16, 1.0, LINE2)
    line(s, cx - 0.42, y + 2.48, cx - 0.14, y + 2.48, LINE2, 1.0, head=True)
    text(s, cx - 1.16, y + 1.16, 0.86, 0.44, "残差\n直通", 8, FAINT, SANS, False, PP_ALIGN.RIGHT, 1.3,
         check=False)
    # 右侧 ×N
    text(s, cx + bw + 0.20, y + 0.34, 0.70, 0.36, "× N 层", 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    line(s, cx + bw + 0.06, y + 0.50, cx + bw + 0.16, y + 0.50, TEAL, 1.2)
    line(s, cx + bw + 0.06, y + 2.30, cx + bw + 0.16, y + 2.30, TEAL, 1.2)
    vr(s, cx + bw + 0.11, y + 0.50, 1.80, 1.2, TEAL)
    text(s, cx, yy + 0.04, bw, 0.26, "→ 输出到下一层", 10, FAINT, SANS, False, PP_ALIGN.CENTER, 1.0,
         check=False)
    # 右：四段解释
    RX = ML + LW + 0.60
    RW = X2 - RX
    vr(s, RX - 0.30, y + 0.06, 3.60, 1.0, LINE)
    EX = [("注意力", "负责“看别的词”，把上下文里相关的位置取过来"),
          ("前馈网络", "每个位置单独过两层全连接，参数大头在这里；MoE 替换的就是这一块"),
          ("残差 ＋ 归一化", "每块输出加回输入，信息和梯度都有直通路，几十上百层才训得动"),
          ("位置信息", "注意力本身不分先后，要另加位置编码：原论文用正余弦，现在主流多用 RoPE")]
    yy = y + 0.06
    for t, b in EX:
        text(s, RX, yy, RW, 0.26, t, 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        nl = nlines(b, 10.5, RW)
        text(s, RX, yy + 0.26, RW, nl * 10.5 * LINES * 1.38 / 72 + 0.05, b, 10.5, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.38, check=False)
        yy += 0.26 + nl * 10.5 * LINES * 1.38 / 72 + 0.22
    # 规模
    box(s, RX, 5.10, RW, 0.94, TEALXL)
    text(s, RX + 0.20, 5.20, RW - 0.40, 0.24, "规模的两个坐标", 9.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.2, check=False)
    text(s, RX + 0.20, 5.44, RW - 0.40, 0.50,
         "原论文 6 层、512 维、8 头；GPT-3 为 96 层、12288 维、96 头，1750 亿参数。", 10.5, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：Vaswani 等 2017；Brown 等 2020（GPT-3）；Su 等 2021（RoPE）")
    return s


# ---------------- F5 注意力三步 ----------------
def f5(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "注意力怎么算：带着“问题”去和其他词“对暗号”",
           "权重经过归一化，但不是平均分配。长上下文里，模型可能难以稳定地定位和使用关键信息。")
    # 上：三步
    ST = [("01 打分", "“它”拿自己的问题，和每个词的标签比对，得到一组分数"),
          ("02 换算比例", "分数换成加起来 = 100% 的权重，重要的词占比高"),
          ("03 加权混合", "按比例把各词的内容加起来，得到“它”在这一层的新含义")]
    bw3 = (CW - 0.34 * 2) / 3
    for i, (t, b) in enumerate(ST):
        x = ML + i * (bw3 + 0.34)
        text(s, x, y + 0.02, bw3, 0.26, t, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        nl = nlines(b, 10, bw3)
        text(s, x, y + 0.30, bw3, nl * 10 * LINES * 1.35 / 72 + 0.05, b, 10, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.35, check=False)
    # 中：权重条（句子「图纸和清单要检查，它」）
    BY = y + 1.22
    hr(s, ML, BY - 0.12, CW, 1.0, LINE2)
    text(s, ML, BY - 0.02, 5.6, 0.24, "示例：“它”这一行，五个词的权重分布（示意，合计 100%）", 9.5, FAINT,
         SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    WD = [("图纸", 22), ("和", 6), ("清单", 19), ("检查", 24), ("它", 29)]
    bx = ML
    wid = (CW - 0.16 * 4) / sum(w for _, w in WD)
    yy = BY + 0.34
    for t, w in WD:
        ww = wid * w
        box(s, bx, yy, ww, 0.52, TEAL if w > 20 else TEALL)
        text(s, bx, yy + 0.13, ww, 0.28, t, 11.5, WHITE, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
        text(s, bx, yy + 0.58, ww, 0.24, f"{w}%", 9.5, TEAL, MONO, True, PP_ALIGN.CENTER, 1.0,
             check=False)
        bx += ww + 0.16
    # 下：多头 + 结论
    DY = yy + 0.96
    hr(s, ML, DY, CW, 1.0, LINE)
    text(s, ML, DY + 0.14, 2.30, 0.26, "多头：同时很多组在看", 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
         check=False)
    text(s, ML + 2.46, DY + 0.16, CW - 2.46, 0.50,
         "有的看指代（“它”指谁），有的看修饰，有的盯数字与单位。生成时有遮罩，不能读后面的 token。",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    text(s, ML, DY + 0.74, CW, 0.56,
         "边界：权重高不等于“理解对了”。这一层只是按相似度搬运信息，最后一步仍要看它拼出来的结论对不对。",
         11.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：Vaswani 等《Attention Is All You Need》（2017）；权重为示意，非真实模型输出")
    return s


# ---------------- F6 架构三态 ----------------
def f6(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "稠密、MoE、稀疏：差在“每次算多少”",
           "总参数看显存，激活参数看钱。两者别混着比 —— 一句“770B 单机可跑”常常漏掉了后半句。")
    G = [("稠密 Dense", "每个 token 都过全部参数",
          ["全部参数都参与计算", "行为稳定、可复现", "参数越多越贵越慢", "例：多数中小模型"], TEALXL),
         ("MoE 混合专家", "每个 token 只走少数专家",
          ["总参数 ≫ 激活参数", "如 770B 总 / 49B 激活", "知识容量大、单次算力小", "例：Mistral、Qwen、混元、DeepSeek"], MINT),
         ("稀疏 Sparse", "只在需要时算需要的那部分",
          ["MoE 属于参数稀疏", "还有稀疏注意力（少看一些位置）", "问“稀疏”要先问哪一种", "长文场景常用"], TEALXL)]
    bw = (CW - 0.32 * 2) / 3
    for i, (t, sub, items, fill) in enumerate(G):
        x = ML + i * (bw + 0.32)
        box(s, x, y + 0.04, bw, 0.06, TEAL)
        text(s, x, y + 0.18, bw, 0.30, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x, y + 0.50, bw, 0.24, sub, 10, TEAL, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        for j, it in enumerate(items):
            ry = y + 0.82 + j * 0.32
            text(s, x + 0.02, ry, 0.16, 0.22, "·", 11, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
                 check=False)
            text(s, x + 0.20, ry, bw - 0.20, 0.24, it, 9.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.2,
                 check=False)
    # 下：两个口径 + 三种注意力
    PY = y + 2.28
    hr(s, ML, PY, CW, 1.0, LINE2)
    text(s, ML, PY + 0.14, 2.30, 0.24, "两个参数口径", 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    P2 = [("总参数", "所有专家加起来的总量；决定显存要装多大"),
          ("激活参数", "一次前向真正算到的量；决定跑得多快多贵")]
    for i, (t, b) in enumerate(P2):
        x = ML + i * 3.30
        text(s, x, PY + 0.42, 3.10, 0.26, t, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
             check=False)
        text(s, x, PY + 0.68, 3.10, 0.44, b, 9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    AX = ML + 6.90
    text(s, AX, PY + 0.14, 5.0, 0.24, "注意力也在省算力", 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    A3 = [("MLA 多头潜在", "把 KV 缓存压小，长文更省显存"),
          ("线性 / 混合线性", "把平方增长压到接近线性，代价是可能丢细节"),
          ("滑动窗口", "只看附近一段，远处靠层数传导")]
    for i, (t, b) in enumerate(A3):
        ry = PY + 0.42 + i * 0.30
        text(s, AX, ry, 1.54, 0.24, t, 9.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, AX + 1.58, ry, 3.40, 0.24, b, 9, GREY, SANS, False, PP_ALIGN.LEFT, 1.25, check=False)
    text(s, ML, 6.32, CW, 0.30,
         "结论：架构决定成本曲线，不决定聪明程度；聪明程度看评测与任务。", 11.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：Shazeer 等 2017（MoE 原始论文）；DeepSeek-V3 与 Kimi K3 技术报告；本稿整理")
    return s


# ---------------- F7 规模定律 ----------------
def f7(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "规模定律：模型、数据、算力一起放大，损失按幂律下降",
           "它预测的是“损失”，不是某项具体能力。能力何时出现、够不够用，仍要实测。")
    LW = 6.10
    IT = [("Kaplan 等（2020）", "损失随参数量、数据量、算力呈幂律下降，部分趋势跨七个数量级以上"),
          ("Chinchilla（2022）", "同样算力下参数和数据同比例放大，约每个参数 20 个 token；700 亿参数＋1.4 万亿 token 胜过 2800 亿参数的 Gopher"),
          ("算力怎么估", "训练算力 ≈ 6 × 参数量 × token 数；Chinchilla 约 5.8×10²³ 次浮点运算"),
          ("现在的瓶颈", "高质量文本有限：转向合成数据、强化学习与推理时计算（见 F8）")]
    yy = y + 0.06
    for t, b in IT:
        hr(s, ML, yy, LW, 1.0, LINE)
        text(s, ML, yy + 0.14, LW, 0.26, t, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        nl = nlines(b, 11, LW)
        text(s, ML, yy + 0.42, LW, nl * 11 * LINES * 1.38 / 72 + 0.05, b, 11, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.38, check=False)
        yy += 0.42 + nl * 11 * LINES * 1.38 / 72 + 0.16
    hr(s, ML, yy, LW, 1.0, LINE)
    RX = ML + LW + 0.52
    RW = X2 - RX
    vr(s, RX - 0.26, y + 0.06, 4.10, 1.0, LINE)
    frame(s, RX, y + 0.10, RW, 3.10, LINE, 1.0, CARD)
    pic(s, "s118_1.png", RX + 0.06, y + 0.16, RW - 0.12, 2.98, mode="fill")
    text(s, RX, y + 3.28, RW, 0.66,
         "双对数坐标下是直线，代表幂律：数据翻倍，损失按固定斜率下降。斜率固定，所以效果能提前估算 —— 这是各家敢投大钱训练的依据。",
         8.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：Kaplan 等 2020（OpenAI）；Hoffmann 等 2022（DeepMind，Chinchilla）；图为示意")
    return s


# ---------------- F8 对齐与推理 ----------------
def f8(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "从会续写到会推理：对齐、强化学习与“想久一点”",
           "预训练只教会它续写。听指令、守边界、多步推理，靠后面的训练，以及回答时多花算力。")
    T3 = [("2022", "指令对齐：InstructGPT",
           "人写示范 ＋ 人给回答排序 → 奖励模型 → 强化学习（RLHF）",
           "13 亿参数的 InstructGPT，人工评比胜过 1750 亿参数的 GPT-3"),
          ("2024", "推理模型：o1", "先写一长段“思考”再回答", "成绩随训练时强化学习、推理时思考时长一起提升"),
          ("2025", "开源推理：DeepSeek-R1", "用能自动判分的题（数学、代码）做大规模强化学习",
           "长链推理随之出现；权重开源")]
    bw = (CW - 0.34 * 2) / 3
    for i, (yr, t, a, b) in enumerate(T3):
        x = ML + i * (bw + 0.34)
        text(s, x, y + 0.02, 1.20, 0.34, yr, 22, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, x, y + 0.46, bw, 1.4, TEAL)
        text(s, x, y + 0.58, bw, 0.28, t, 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        nl = nlines(a, 10.5, bw)
        text(s, x, y + 0.90, bw, nl * 10.5 * LINES * 1.38 / 72 + 0.05, a, 10.5, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.38, check=False)
        ay = y + 0.90 + nl * 10.5 * LINES * 1.38 / 72 + 0.12
        nl2 = nlines(b, 9.5, bw)
        text(s, x, ay, bw, nl2 * 9.5 * LINES * 1.35 / 72 + 0.05, b, 9.5, TEAL, SANS, False,
             PP_ALIGN.LEFT, 1.35, check=False)
    # 推理时计算
    BY = y + 2.30
    hr(s, ML, BY, CW, 1.0, LINE2)
    text(s, ML, BY + 0.14, 2.60, 0.26, "推理时计算", 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
         check=False)
    text(s, ML + 2.24, BY + 0.16, CW - 2.24, 0.50,
         "同一个模型，多想几步、多采样几次再投票，难题正确率会上升；代价是更慢、更贵 —— 思考过程的 token 也计费。",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    # 三条注意
    N3 = [("和正文的关系", "“按人的偏好打分”就是迎合的来源（P10、P14）"),
          ("它不能保证什么", "思考模式能减少粗心错，但算错、编来源一样会发生"),
          ("怎么用", "难题值得让它多想；查事实别指望它想 —— 回原文")]
    for i, (t, b) in enumerate(N3):
        x = ML + i * (bw + 0.34)
        box(s, x, BY + 0.86, bw, 0.86, MINT if i != 1 else ORANGEL)
        text(s, x + 0.18, BY + 0.96, bw - 0.36, 0.24, t, 9.5, TEAL if i != 1 else ORANGE, SANS,
             True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
        nl = nlines(b, 9.5, bw - 0.36)
        text(s, x + 0.18, BY + 1.20, bw - 0.36, nl * 9.5 * LINES * 1.3 / 72 + 0.05, b, 9.5, BODY,
             SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：Ouyang 等 2022（InstructGPT）；OpenAI《Learning to Reason with LLMs》2024-09；DeepSeek-AI 2025（arXiv 2501.12948）")
    return s


# ---------------- F9 多模态 ----------------
def f9(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "多模态：图怎么进模型，图和视频又是怎么画出来的",
           "“能看图”有档次：分辨率、切块方式、图有没有真的送进模型，决定了它能看见什么。")
    # 左：三块
    LW = 6.40
    IT = [("看图：ViT（2020）", "把图切成 16×16 像素的小块，每块当一个“词”，和文字一样进 Transformer"),
          ("图文对齐：CLIP（2021）", "约 4 亿对“图片—说明文字”训练，让图和字落在同一个向量空间"),
          ("画图：扩散模型（2020）", "从纯噪声出发，每一步去掉一点噪声；视频生成把这一步扩展到时间轴")]
    yy = y + 0.06
    for t, b in IT:
        hr(s, ML, yy, LW, 1.0, LINE)
        text(s, ML, yy + 0.13, LW, 0.26, t, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        nl = nlines(b, 10.5, LW)
        text(s, ML, yy + 0.41, LW, nl * 10.5 * LINES * 1.38 / 72 + 0.05, b, 10.5, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.38, check=False)
        yy += 0.41 + nl * 10.5 * LINES * 1.38 / 72 + 0.16
    hr(s, ML, yy, LW, 1.0, LINE)
    # 和图纸案例的关系
    box(s, ML, yy + 0.14, LW, 1.06, ORANGEL)
    text(s, ML + 0.20, yy + 0.24, LW - 0.40, 0.24, "和图纸案例的关系", 9.5, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, ML + 0.20, yy + 0.48, LW - 0.40, 0.62,
         "只靠 OCR 的配置零命中（P31）—— 图里的线、尺寸和符号根本没有进模型。文字识别得再准，也不等于看懂了图。",
         10, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    # 右：两张示意图
    RX = ML + LW + 0.52
    RW = X2 - RX
    vr(s, RX - 0.26, y + 0.06, 4.10, 1.0, LINE)
    frame(s, RX, y + 0.10, RW, 1.90, LINE, 1.0, CARD)
    pic(s, "s120_1.png", RX + 0.10, y + 0.16, RW - 0.20, 1.78, mode="fill")
    text(s, RX, y + 2.06, RW, 0.24, "挡块图纸局部切成小块，每块变成一个向量（示意）", 8.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.2, check=False)
    frame(s, RX, y + 2.40, RW, 1.30, LINE, 1.0, CARD)
    pic(s, "s120_2.png", RX + 0.10, y + 2.46, RW - 0.20, 1.18, mode="fit")
    text(s, RX, y + 3.78, RW, 0.24, "噪声 → 逐步去噪 → 图像（示意）", 8.5, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：Dosovitskiy 等 2020（ViT）；Radford 等 2021（CLIP）；Ho 等 2020（DDPM）；图为本地图纸局部，切块与去噪为示意")
    return s


# ---------------- Z 来源 ----------------
def z1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "参考资料来源：按主题编号，逐条可查",
           "正文与附录里的每个数字都挂在这张表上。编号与源稿一致（S01–S56），便于与原版对照。")
    G = [("S01–S08", "模型与参数", "各家模型页、技术报告、发布公告"),
         ("S09–S16", "能力与成本", "artificialanalysis.ai、Epoch、FrontierMath、ARC Prize、LMArena、LiveBench"),
         ("S17–S27", "订阅与工具", "各产品定价页与帮助中心、magpie、CLIProxyAPI、Cline、Hark"),
         ("S28–S32 · S36", "政策与安全", "OpenAI 使用政策、Anthropic AUP、OWASP LLM Top 10、备案公告"),
         ("S33–S35", "历史与原理", "Minsky & Papert 1969、Rumelhart 等 1986、Vaswani 等 2017、Kaplan 2020、Chinchilla 2022"),
         ("S37–S38", "能力与成本（续）", "Chatbot Arena 与各家评测说明、模型卡"),
         ("S39–S44", "破限与诉讼", "dsh-purge、社区破限帖、LASST 诉 OpenAI 案相关报道"),
         ("S45–S50", "AI 攻防事件", "Hugging Face 公告、OpenAI 声明、CrowdStrike 归因、BBC / CNBC / NYT / The Record"),
         ("S51–S56", "社区与客户端", "linux.do 相关帖、bbs.kanxue.com、各客户端 README")]
    yy = y + 0.06
    for no, nm, ds in G:
        hr(s, ML, yy, 8.20, 1.0, LINE)
        text(s, ML, yy + 0.13, 1.42, 0.24, no, 9.5, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 1.50, yy + 0.11, 1.72, 0.24, nm, 10.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
             check=False)
        nl = nlines(ds, 9.5, 4.80)
        text(s, ML + 3.30, yy + 0.12, 4.90, nl * 9.5 * LINES * 1.3 / 72 + 0.05, ds, 9.5, GREY, SANS,
             False, PP_ALIGN.LEFT, 1.3, check=False)
        yy += max(0.46, 0.14 + nl * 9.5 * LINES * 1.3 / 72 + 0.12)
    hr(s, ML, yy, 8.20, 1.0, LINE)
    # 右侧：本地案例索引
    RX = ML + 8.62
    RW = X2 - RX
    vr(s, RX - 0.26, y + 0.06, 4.06, 1.0, LINE)
    text(s, RX, y + 0.06, RW, 0.24, "本地案例索引", 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    LC = [("图纸审查", "P29–P32 · 原始记录与 10 份 Excel", TEAL),
          ("修编报告", "P33–P38 · 四份报告与勘误函", TEAL),
          ("鹈鹕盲测", "P19–P20 · 10 份 HTML 与录屏", TEAL),
          ("施工图做表", "P41 · 126 个文件索引表", TEAL),
          ("小虎沥实战", "P46–P49 · 工作流与多 Agent 记录", TEAL)]
    yy = y + 0.36
    for a, b, col in LC:
        text(s, RX, yy, RW, 0.24, a, 10.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        text(s, RX, yy + 0.22, RW, 0.24, b, 8.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.58
    box(s, RX, 5.36, RW, 0.86, ORANGEL)
    text(s, RX + 0.18, 5.46, RW - 0.36, 0.24, "口径与边界", 9, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.8, check=False)
    text(s, RX + 0.18, 5.68, RW - 0.36, 0.46,
         "图纸：mm/cm 不计为问题，21 条待核未计；报告：汇总表口径（不含暂列金额），112 元差额单独说明。",
         8.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：编号与源稿一致；逐条链接见下一页")
    return s


def z2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "链接清单：政策、判例、工具与社区",
           "只列官方页面与一手来源；社区帖标注为社区。查询日期统一为 2026-10-18。")
    G = [("政策与合规", ["openai.com/policies/usage-policies/", "anthropic.com/legal/aup",
                     "genai.owasp.org/llm-top-10/", "cac.gov.cn（生成式 AI 备案公告）"]),
         ("工具与开源", ["github.com/yetone/magpie", "github.com/router-for-me/CLIProxyAPI",
                    "github.com/YuJunZhiXue/dsh-purge", "github.com/himovo/movo",
                    "github.com/Stonewuu/agenteam"]),
         ("社区", ["linux.do（邀请制）", "bbs.kanxue.com/thread-255826.htm",
                "daily.juya.uk（中文日报）"]),
         ("事件与判例", ["huggingface.co/blog/security-incident-july-2026",
                    "openai.com/index/hugging-face-model-evaluation-security-incident/",
                    "bbc.com/news/articles/cd9w22n9e4go",
                    "cnbc.com/2026/09/30/openai-sued-cyberattack.html",
                    "therecord.media/south-korean-bank-hacks-ai-agents"]),
         ("榜单与检测", ["artificialanalysis.ai", "epoch.ai/frontiermath", "arcprize.org",
                    "lmarena.ai", "livebench.ai", "lm.ikale.io", "marginlab.ai"])]
    cw = (CW - 0.62) / 2
    TOPY = y + 0.14
    # 左列
    yy = TOPY
    for gi in (0, 1, 2):
        t, links = G[gi]
        text(s, ML, yy, cw, 0.22, t, 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
        ly = yy + 0.28
        for ln in links:
            text(s, ML + 0.02, ly, cw - 0.04, 0.22, ln, 9, BODY, MONO, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
            ly += 0.25
        yy = ly + 0.22
    # 右列
    RX = ML + cw + 0.62
    yy = TOPY
    for gi in (3, 4):
        t, links = G[gi]
        text(s, RX, yy, cw, 0.22, t, 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
        ly = yy + 0.28
        for ln in links:
            text(s, RX + 0.02, ly, cw - 0.04, 0.22, ln, 9, BODY, MONO, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
            ly += 0.25
        yy = ly + 0.22
    vr(s, RX - 0.31, TOPY - 0.04, 4.62, 1.0, LINE)
    hr(s, ML, 6.10, CW, 1.0, LINE)
    text(s, ML, 6.24, CW, 0.30,
         "引用原则：先看官方原帖与原始文档，再看二手整理；社区帖只当线索，不当结论。", 11, TEAL, SANS,
         True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：以上各站点（2026-10-18 查询）；社区内容为个人表述，未经独立核验")
    return s
