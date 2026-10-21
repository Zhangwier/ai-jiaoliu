# -*- coding: utf-8 -*-
"""Rebuild appendix D2 / D3 / Z of 成品07 in the deck's own style (v3)."""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

FONT = "微软雅黑"
PANOSE = "020B0503020204020204"
C_TITLE = "14213D"
C_BODY = "2B3445"
C_MUTED = "6B7385"
C_TEAL = "0F766E"
C_ORANGE = "D9622B"
C_HEAD = "12203A"
C_LINE = "DAD7CE"
C_WHITE = "FFFFFF"
C_WARM = "FBE9DE"
C_COOL = "E1F0EE"
C_ZEBRA = "ECEAE4"
C_BG = "F6F5F1"

L = 0.60
W = 12.13


# ------------------------------------------------------------------ primitives
def set_bg(slide, color=C_BG):
    cSld = slide._element.find(qn("p:cSld"))
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    cSld.insert(0, parse_xml(
        '<p:bg %s><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>" % (nsdecls("p", "a"), color)))


def style_run(run, size, bold, color):
    rPr = run._r.get_or_add_rPr()
    rPr.set("sz", str(int(round(size * 100))))
    rPr.set("b", "1" if bold else "0")
    rPr.set("i", "0")
    for tag in ("a:solidFill", "a:latin", "a:ea", "a:cs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
    rPr.append(parse_xml(
        '<a:solidFill %s><a:srgbClr val="%s"/></a:solidFill>' % (nsdecls("a"), color)))
    for tag in ("a:latin", "a:ea", "a:cs"):
        rPr.append(parse_xml(
            '<%s %s typeface="%s" panose="%s" charset="-122"/>'
            % (tag, nsdecls("a"), FONT, PANOSE)))


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.MIDDLE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    bodyPr = tf._txBody.bodyPr
    for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
        for e in bodyPr.findall(qn(tag)):
            bodyPr.remove(e)
    bodyPr.append(parse_xml("<a:noAutofit %s/>" % nsdecls("a")))
    return box


def put(slide, x, y, w, h, content, size=15, color=C_BODY, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15):
    if isinstance(content, str):
        paras = [[(content, size, bold, color)]]
    elif content and isinstance(content[0], str):
        paras = [[(t, size, bold, color)] for t in content]
    elif content and isinstance(content[0], tuple):
        paras = [list(content)]
    else:
        paras = [list(p) for p in content]

    box = textbox(slide, x, y, w, h, anchor)
    tf = box.text_frame
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for t, sz, bd, cl in runs:
            r = p.add_run()
            r.text = t
            style_run(r, sz, bd, cl)
    return box


def rect(slide, x, y, w, h, fill=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def hline(slide, x, y, w, color=C_LINE, pts=0.75):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(pts)
    return ln


def vline(slide, x, y, h, color=C_LINE, pts=0.75):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x), Inches(y + h))
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(pts)
    return ln


def arrow(slide, x1, y1, x2, y2, color=C_TEAL, pts=1.25):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(pts)
    ln.line._get_or_add_ln().append(parse_xml(
        '<a:tailEnd %s type="triangle" w="med" len="med"/>' % nsdecls("a")))
    return ln


def dot(slide, cx, cy, d=0.14, color=C_TEAL):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(color)
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


# ------------------------------------------------------------------ page chrome
def chrome(slide, page, title, subtitle, source, notes):
    set_bg(slide)
    put(slide, L, 0.32, 6.0, 0.30, "附录  会后速查", size=12, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    put(slide, 11.73, 0.32, 1.00, 0.30, str(page), size=12, color=C_MUTED,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
    rect(slide, L, 0.66, 0.50, 0.06, C_TEAL)
    put(slide, L, 0.78, W, 0.75, title, size=26, bold=True, color=C_TITLE, spacing=1.0)
    if subtitle:
        put(slide, L, 1.50, W, 0.45, subtitle, size=17, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP, spacing=1.0)
    hline(slide, L, 7.00, W)
    put(slide, L, 7.06, 10.50, 0.30, source, size=10.5, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP)
    put(slide, 10.73, 7.06, 2.00, 0.30, "来源索引 ›", size=10.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    slide.notes_slide.notes_text_frame.text = notes


def section_bar(slide, x, y, w, label, color=C_HEAD, size=15, h=0.44):
    rect(slide, x, y, w, h, color)
    put(slide, x + 0.20, y, w - 0.35, h, label, size=size, bold=True, color=C_WHITE)


def callout(slide, label, body, kind="warn"):
    bar = C_WARM if kind == "warn" else C_COOL
    edge = C_ORANGE if kind == "warn" else C_TEAL
    rect(slide, L, 6.35, W, 0.50, bar)
    rect(slide, L, 6.35, 0.09, 0.50, edge)
    put(slide, 0.85, 6.35, 1.05, 0.50, label, size=14, bold=True, color=edge)
    put(slide, 1.78, 6.35, 10.70, 0.50, body, size=18, bold=True,
        color=C_TITLE, spacing=1.0)


def rowband(slide, x, y, w, h, i, fill=None):
    rect(slide, x, y, w, h, fill or (C_WHITE if i % 2 == 0 else C_ZEBRA))


def node(slide, x, y, w, h, title, sub=None, fill=C_WHITE, edge=C_LINE,
         tcolor=C_TITLE, subcolor=C_MUTED, tsize=13.5, ssize=11):
    sh = rect(slide, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
    sh.line.color.rgb = RGBColor.from_string(edge)
    sh.line.width = Pt(0.9)
    if sub:
        put(slide, x + 0.10, y + 0.06, w - 0.20, h * 0.48, title, size=tsize,
            bold=True, color=tcolor, align=PP_ALIGN.CENTER, spacing=1.0)
        put(slide, x + 0.10, y + h * 0.50, w - 0.20, h * 0.44, sub, size=ssize,
            color=subcolor, align=PP_ALIGN.CENTER, spacing=1.0)
    else:
        put(slide, x + 0.10, y, w - 0.20, h, title, size=tsize, bold=True,
            color=tcolor, align=PP_ALIGN.CENTER, spacing=1.0)
    return sh


# ================================================================ D2-1 原理
def page_d2_intro(slide, page):
    chrome(
        slide, page,
        "D2  中转站与代理池：多出来的那一层",
        "所有“反代”都是同一件事：你的请求先发给一台别人的或自己的服务器，它再替你发给厂商",
        "来源：CLIProxyAPI、magpie 文档；OpenAI、Anthropic 使用政策；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页是 D2 的总图，先把两个词分开：中转站和代理池。\n\n"
            "共同点只有一句：请求不再直连厂商，而是先到一台中间服务器。"
            "区别在于这台服务器是谁的、它手里有什么。\n\n"
            "中转站，是别人的服务器。你从它那儿买“额度”，它拿自己的 Key 或自己注册的号去连厂商。"
            "对你来说只有一个地址和一个 Key；对厂商来说，请求来自它，不来自你。"
            "所以数据、密钥、模型真伪、条款、资金这五件事，全都落在它身上——这就是下一页要展开的风险。\n\n"
            "代理池，是自己的服务器。上面挂的是你自己的账号登录态或自己的 Key，"
            "网关按剩余额度、按重置时间挑一个发出去。它解决的是“一个人一份订阅不够用”，"
            "不解决“没有订阅也能用”。\n\n"
            "画图的时候指一下两条链路：上面那条多一个方框，方框不在你手里；"
            "下面那条也多一个方框，但方框在你自己的机器上。同样是“多一层”，风险完全不同。"
        ))

    # --- 中转站
    section_bar(slide, L, 1.90, 6.10, "① 中转站：中间那台机器是别人的", h=0.36, size=13.5)
    node(slide, 0.60, 2.30, 1.95, 0.72, "你的客户端", "Cursor / 网页", tsize=12.5, ssize=10)
    arrow(slide, 2.60, 2.66, 3.05, 2.66, C_ORANGE)
    node(slide, 3.08, 2.30, 1.90, 0.72, "中转站", "别人的域名 + Key",
         fill=C_WARM, edge=C_ORANGE, tcolor=C_ORANGE, subcolor=C_BODY,
         tsize=12.5, ssize=10)
    arrow(slide, 5.03, 2.66, 5.48, 2.66, C_ORANGE)
    node(slide, 5.51, 2.30, 1.30, 0.72, "厂商 API", "只认它的 Key", tsize=12, ssize=10)
    put(slide, 0.60, 3.10, 6.10, 0.24,
        "你只看到：一个地址、一个 Key、一份价目表", size=11.5, color=C_MUTED)

    # --- 代理池
    section_bar(slide, 6.95, 1.90, 5.78, "② 代理池：中间那台机器是自己的", color=C_TEAL,
                h=0.36, size=13.5)
    node(slide, 6.95, 2.30, 1.50, 0.72, "你的 Agent", "只认一个地址", tsize=12, ssize=10)
    arrow(slide, 8.50, 2.66, 8.86, 2.66, C_TEAL)
    node(slide, 8.89, 2.26, 1.42, 0.80, "本地网关", "127.0.0.1:3425",
         fill=C_COOL, edge=C_TEAL, tcolor=C_TEAL, subcolor=C_BODY, tsize=12, ssize=9)
    px, pw, ph, gap = 10.42, 2.18, 0.28, 0.07
    for i, lab in enumerate(["账号 A01", "账号 A02", "账号 A03"]):
        cy = 2.22 + i * (ph + gap)
        node(slide, px, cy, pw, ph, lab, fill=C_WHITE, edge=C_LINE, tcolor=C_BODY,
             tsize=10)
        arrow(slide, 10.33, 2.66, px - 0.03, cy + ph / 2, C_TEAL, 0.75)
    put(slide, 6.95, 3.30, 5.78, 0.24,
        "省下的是：一个人也能一直有额度；代价是多个身份挤在一个地址后面", size=11.5,
        color=C_MUTED)

    # --- 一句话区别
    rect(slide, L, 3.50, W, 0.48, C_COOL)
    put(slide, 0.85, 3.50, 3.30, 0.48, "两条链路都一样“多一层”", size=13.5, bold=True,
        color=C_TEAL)
    put(slide, 3.95, 3.50, 8.55, 0.48,
        "区别只在：那一层里是你的数据、你的 Key、你的账号，还是在别人手里", size=13.5,
        color=C_BODY)

    # --- 对照表
    rect(slide, L, 4.14, W, 0.36, C_HEAD)
    for x, w, t in [(3.40, 4.55, "中转站"), (8.10, 4.45, "代理池")]:
        put(slide, x, 4.14, w, 0.36, t, size=13, bold=True, color=C_WHITE)

    cmp_rows = [
        ("机器是谁的", "别人的域名、别人的服务器", "自己的电脑或自己的服务器"),
        ("凭据是谁的", "对方持有的 Key 或代注册的号", "自己的账号登录态，或自己的 Key"),
        ("解决什么", "没有账号也能用（但要付溢价）", "账号够了，额度和稳定性不够"),
        ("主要代价", "数据、密钥、模型真伪、条款与资金", "违反一人一号；账号可能被停用"),
    ]
    y = 4.50
    for i, (a, b, c) in enumerate(cmp_rows):
        rowband(slide, L, y, W, 0.42, i)
        put(slide, 0.72, y, 2.60, 0.42, a, size=12.5, bold=True, color=C_ORANGE)
        put(slide, 3.40, y, 4.55, 0.42, b, size=12, color=C_BODY)
        put(slide, 8.10, y, 4.45, 0.42, c, size=12, color=C_BODY)
        y += 0.42

    callout(slide, "记住", "中转站卖你“能用”，代理池帮你“够用”——两件事，风险差一个量级",
            kind="cool")


# ================================================================ D2-2 反代风险
def page_d2a(slide, page):
    chrome(
        slide, page,
        "D2  中转站的风险：五件事都不在你手里",
        "便宜是真的，代价也是真的：数据、密钥、模型真伪、条款、资金，五项里没有一项你能自己验证",
        "来源：OpenAI、Anthropic 使用政策；lm.ikale.io 文档；L 站中转纠纷帖；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "上一页画了链路，这一页讲代价。五条都按“它怎么发生 → 你怎么发现”来讲。\n\n"
            "数据：输入、附件、输出都要过对方服务器。它有没有留、留多久、会不会被拿去训练，"
            "你看不到日志，也拿不到承诺。\n\n"
            "密钥：交出 Key 就等于交出计费权。更常见的是“共享 Key”——几十个人用同一个 Key，"
            "哪天上游被封，你跟着一起停，而且账单上写的是别人的名字。\n\n"
            "模型：一句“我是 Opus”证明不了任何事。便宜的号池里常见小模型冒充大模型，"
            "要靠固定题目＋随机数指纹来验，而这本身就说明你没法直接信。\n\n"
            "条款：绕地区限制、共享 Key、多号刷额度，都是明确的停用理由；"
            "一旦停用，预充的钱通常不退款。\n\n"
            "资金：预充值、无发票、跑路。这一条最容易验证：问一句“能不能开票”，"
            "大部分中转站就答不上来。\n\n"
            "最后落到自查三问，让对方自己回答：Base URL 是谁的域名、Key 是谁开的、账单抬头是谁。"
        ))

    rect(slide, L, 1.92, W, 0.44, C_HEAD)
    for x, w, t in [(0.72, 2.35, "哪一件事"), (3.17, 4.55, "它是怎么发生的"),
                    (7.87, 2.10, "你看到的样子"), (10.07, 2.60, "怎么发现")]:
        put(slide, x, 1.92, w - 0.10, 0.44, t, size=14, bold=True, color=C_WHITE)

    rows = [
        ("数据", "输入、附件、输出都过它的服务器，留存与用途你看不到",
         "“我们不留日志”", "要一份书面承诺"),
        ("密钥", "它拿你的 Key 或共享 Key 去刷别处，账单仍记你头上",
         "“Key 随便填”", "看 Key 前缀归属"),
        ("模型被换", "便宜模型冒充贵模型，回复里照样自称是贵的那个",
         "“保证官方原版”", "固定题目＋指纹"),
        ("条款", "绕地区、共享号、多号刷额度，都是停用理由",
         "“稳定不掉线”", "看停用后能否退款"),
        ("资金", "预充值、无发票，跑路就没了",
         "“充值 5 折”", "问能不能开票"),
    ]
    y = 2.36
    for i, (a, b, c, d) in enumerate(rows):
        rowband(slide, L, y, W, 0.62, i)
        put(slide, 0.72, y, 2.35, 0.62, a, size=14, bold=True, color=C_ORANGE)
        put(slide, 3.17, y, 4.55, 0.62, b, size=13.5, color=C_BODY, spacing=1.1)
        put(slide, 7.87, y, 2.05, 0.62, c, size=13.5, color=C_MUTED, spacing=1.1)
        put(slide, 10.07, y, 2.58, 0.62, d, size=13.5, color=C_TEAL, spacing=1.1)
        y += 0.62

    rect(slide, L, 5.52, W, 0.62, C_COOL)
    put(slide, 0.80, 5.52, 1.75, 0.62, "自查三问", size=13.5, bold=True, color=C_TEAL)
    put(slide, 2.55, 5.52, 9.95, 0.62,
        "Base URL 是谁的域名？Key 是谁开的？账单抬头写的是谁？", size=14,
        color=C_BODY)

    callout(slide, "防护", "资料只走批准渠道；Key 单独开、设上限；统一出口用自建或采购的网关")


# ================================================================ D2-3 账号池
def page_d2b(slide, page):
    chrome(
        slide, page,
        "D2  代理池：一份订阅，几个人排队",
        "池子不是更强的模型，而是把多个账号的登录态放进一个网关，谁有额度谁上",
        "来源：magpie 文档（usemagpie.ai）；CLIProxyAPI README；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "账号池解决一个很具体的问题：订阅按“一个人用”卖，额度按 5 小时和每周滚动，"
            "白天用完了晚上就没了。把几个账号放进同一个网关，就能一直有额度。\n\n"
            "四步：登录、收口、调度、计量。最容易忽略的是第一步——凭据是文件或钥匙串里的东西，"
            "复制走等于账号被复制走，改密码是唯一能收回的办法。\n\n"
            "第二步就是“反代”真正发生的地方：网关对外只暴露一个地址，"
            "Agent 配置里写的是 127.0.0.1，看起来完全没有越过任何东西。\n\n"
            "右边三条是风险。重点讲第三条：厂商风控看的不是“你是谁”，是“形状对不对”——"
            "同 IP 多号、并发超过一个人的打字速度、今天在东京明天在法兰克福。\n\n"
            "给一个诚实判断：自己两三个订阅、或团队共用一台机器，是常见做法；"
            "但常见不等于条款允许，尤其不要用单位资料去跑。"
        ))

    section_bar(slide, L, 1.92, 6.10, "四步：从登录到计量")
    steps = [
        ("① 登录", "每个账号在官方客户端登录一次，凭据留在本机"),
        ("② 收口", "本地网关对外给兼容接口，Agent 只认一个地址"),
        ("③ 调度", "轮询、按剩余额度排序、失败自动切下一个"),
        ("④ 计量", "5 小时与每周窗口、token 与花费分开统计"),
    ]
    vline(slide, 0.92, 2.60, 3.30)
    y = 2.50
    for k, v in steps:
        dot(slide, 0.92, y + 0.15)
        put(slide, 1.20, y, 5.30, 0.32, k, size=15, bold=True, color=C_TITLE)
        put(slide, 1.20, y + 0.36, 5.30, 0.34, v, size=14, color=C_BODY,
            spacing=1.1, anchor=MSO_ANCHOR.TOP)
        y += 0.84

    put(slide, 1.20, 5.88, 5.40, 0.30,
        "一个地址背后，可以有 1 个账号，也可以有 20 个", size=13.5, color=C_MUTED)

    section_bar(slide, 6.95, 1.92, 5.78, "它把什么变成了风险", color=C_ORANGE)
    cards = [
        ("一人一号", "订阅按“一个人使用”卖；多账号集中调度，通常在条款之外。"),
        ("登录态即身份", "凭据文件被复制走，等于账号被复制走——只能改密码收回。"),
        ("风控看的是形状", "同 IP 多号、并发超过一个人、地区漂移，都是可识别特征。"),
    ]
    y = 2.48
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 0.96, C_WHITE)
        put(slide, 7.20, y + 0.11, 5.30, 0.30, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.44, 5.30, 0.44, b, size=14, color=C_BODY,
            spacing=1.12, anchor=MSO_ANCHOR.TOP)
        y += 1.06

    rect(slide, 6.95, 5.72, 5.78, 0.54, C_COOL)
    put(slide, 7.18, 5.72, 5.40, 0.54,
        "相对合理的用法：自己的两三个订阅，或团队共用一台机器", size=13.5,
        color=C_BODY, spacing=1.1)

    callout(slide, "注意", "账号池把“额度不够”变成“身份不够”：省下的是钱，押上的是账号")


# ================================================================ D2-4 实拍：账号表
def page_shot_agents(slide, page, img):
    chrome(
        slide, page,
        "D2  开箱（一）：一行一个账号，全在一张表里",
        "这是工具里的账号清单：左边是 Agent，中间是登录状态，右边是额度类型",
        "来源：magpie 本地界面截图（账号、邮箱、Key 已打码）；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "这是真实界面，账号、邮箱、Key 已经全部换成 A01–A11 的编号。\n\n"
            "看三列。第一列是 Agent 或服务商的名字：Antigravity、WorkBuddy、Trae CN、Qoder、"
            "Cline、OpenCode Zen…… 它们各自本来有自己的登录方式和配置文件。\n\n"
            "第二列是状态：有的是“已登录 + 邮箱”（订阅账号），有的是“5 个密钥”（API Key 池），"
            "有的是一个 base URL 加 1 个模型。这一列解释了池子里装的是什么——"
            "不是模型，是“凭据”。\n\n"
            "第三列是额度类型：Google AI Pro、Free、Credits、personal_stand…"
            "这就是调度器排序的依据。\n\n"
            "最下面那行 MP 是 API Key 形态，右侧只剩 sk-F••••BW8 的尾巴——"
            "这类截图一旦漏打码，等于把 Key 公开。所以这一页的图里，邮箱、手机号、Key 一律换成编号。\n\n"
            "讲的时候强调：这张表看起来完全正常，它就是个配置界面；"
            "风险不在“看起来可疑”，在“这些东西本来不该集中在一个文件里”。"
        ))

    pic = slide.shapes.add_picture(img, Inches(0.60), Inches(1.92), width=Inches(8.35))
    pic.line.color.rgb = RGBColor.from_string(C_LINE)
    pic.line.width = Pt(0.75)

    x = 9.15
    put(slide, x, 1.95, 3.55, 0.30, "这张图里能读出三件事", size=13.5, bold=True,
        color=C_TEAL)
    notes3 = [
        ("装的是凭据", "“已登录 + 邮箱”= 订阅态；“5 个密钥”= Key 池。装的是身份，不是模型。"),
        ("额度决定排序", "Pro / Free / Credits / personal——调度按这一列挑下一个。"),
        ("一张表就是全部", "11 个入口、十几个账号，都在同一个配置里；丢一次就是全丢。"),
    ]
    y = 2.35
    for t, b in notes3:
        put(slide, x, y, 3.55, 0.28, t, size=13, bold=True, color=C_TITLE)
        put(slide, x, y + 0.28, 3.55, 0.62, b, size=11.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.98

    put(slide, x, 5.35, 3.55, 0.55,
        "截图一律打码：账号→A01，邮箱→账号 A01，Key→只留首尾", size=11, color=C_ORANGE,
        spacing=1.2, anchor=MSO_ANCHOR.TOP)

    callout(slide, "提醒",
            "这类“一张表管所有”的便利，同时也是一张“全丢清单”", kind="cool")


# ================================================================ D2-5 实拍：路由
def page_shot_route(slide, page, img):
    chrome(
        slide, page,
        "D2  开箱（二）：一次请求，网关挑了哪个号",
        "左边是一台本地网关，右边是同一个上游下面挂的一排账号，中间那条亮线就是这次的选择",
        "来源：magpie 路由页截图（账号已打码）；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "这张图把“池”画出来了：左边一个入口，中间一台网关（127.0.0.1:3425），"
            "右边一列全是账号，全都打同一个上游、同一个模型。\n\n"
            "从上往下看状态，正好是四种典型：\n"
            "已用 0%——全新、没动过；\n"
            "额度用尽·周一 13:04 恢复——撞到窗口了，还剩 21 小时；\n"
            "额度用尽·22:55 恢复——还剩 7 小时 19 分；\n"
            "602.79 / 3,000 积分、958.57 / 3,300 积分——按积分计量的客户端。\n\n"
            "高亮那一条写着“回答了这个请求”，意思是这一次请求就落在它身上。"
            "这就是 smart 策略在做的判断：优先用还剩额度、且最早重置的那个，浪费最少。\n\n"
            "这张图为什么值得放：它把“账号池”从概念变成了可核对的东西——"
            "每个号的剩余、恢复时间、这次是谁答的，都摆在明面上。"
            "反面是：这么清楚的一张表，也等于把“你手上有多少号”摊开给任何看到屏幕的人。"
        ))

    pic = slide.shapes.add_picture(img, Inches(0.60), Inches(1.92), width=Inches(7.55))
    pic.line.color.rgb = RGBColor.from_string(C_LINE)
    pic.line.width = Pt(0.75)

    x = 8.42
    section_bar(slide, x, 1.92, 4.31, "顺着这张图看四件事", color=C_HEAD, size=13.5, h=0.40)
    items = [
        ("一个入口", "左边只有一个地址；所有 Agent 都连它，不直连上游。"),
        ("一排账号", "同一个模型、同一家上游，区别只在剩余额度与恢复时间。"),
        ("状态即排序", "已用 0% → 有额度；用尽 → 标出恢复时间，暂不派活。"),
        ("这次是谁", "高亮那一行写着“回答了这个请求”，可以逐次核对。"),
    ]
    y = 2.42
    for t, b in items:
        put(slide, x, y, 4.31, 0.26, t, size=13, bold=True, color=C_TEAL)
        put(slide, x, y + 0.26, 4.31, 0.50, b, size=11.5, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.86

    put(slide, x, 5.86, 4.31, 0.42,
        "截图里的账号名、手机号已换成 A01–A10", size=10.5, color=C_ORANGE,
        spacing=1.15, anchor=MSO_ANCHOR.TOP)

    callout(slide, "读法", "池子的全部本事，就是把“谁还有额度”这件事算清楚，然后悄悄换号", kind="cool")


# ================================================================ D2-6 工具
def page_d2c(slide, page):
    chrome(
        slide, page,
        "D2  工具：magpie、CPA 与自建网关",
        "magpie 在桌面，CPA（CLIProxyAPI）在服务器：能力接近，暴露面不同",
        "来源：github.com/yetone/magpie；github.com/router-for-me/CLIProxyAPI；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "两个工具做同一件事：把多个来源（订阅登录态、API Key、本地模型）收成一个兼容接口。"
            "区别在形态和暴露面。\n\n"
            "magpie 是桌面应用加本地网关，默认 127.0.0.1:3425，同时支持 OpenAI Chat / Responses、"
            "Anthropic Messages 和 Gemini 协议并互相转换。改配置时原子改写、保留注释；"
            "路由组可以按额度、按重置时间、按用量挑账号。前面两页的截图就是它。"
            "要提醒：它默认会读 Claude Code 和 Codex 的登录信息，设置里的“订阅账号”可以关。\n\n"
            "CLIProxyAPI 是 Go 写的服务端代理，OAuth 登录各家，多账号轮询，适合放服务器或 NAS。"
            "管理面板用来盯每个号的 5 小时与每周窗口。一旦上公网，必须加访问密码和客户端 Key——"
            "这类面板被扫到的事故不少。\n\n"
            "最后一行：真要给单位用，唯一说得清“谁用了多少、数据去了哪”的做法，是自己搭一个网关。"
        ))

    rect(slide, L, 1.92, W, 0.44, C_HEAD)
    for x, w, t in [(0.72, 2.35, "工具"), (3.17, 4.60, "它做什么"),
                    (7.87, 2.10, "适合谁"), (10.07, 2.60, "要留神")]:
        put(slide, x, 1.92, w - 0.10, 0.44, t, size=14, bold=True, color=C_WHITE)

    rows = [
        ("magpie",
         "桌面网关 127.0.0.1:3425，把订阅与 Key 收成一个地址，按 Agent 改写配置",
         "一台电脑、多份订阅",
         "默认读 Claude Code、Codex 登录态，可关"),
        ("CLIProxyAPI（CPA）",
         "服务端代理，OAuth 登录各家，对外给兼容接口，多账号轮询",
         "服务器或 NAS",
         "上公网要加访问密码与客户端 Key"),
        ("管理面板（CPAMC 等）",
         "看每号 5 小时与每周额度、请求记录、token 与花费",
         "账号一多就需要",
         "面板本身也是攻击面，别和数据面同域"),
        ("自建或采购的网关",
         "统一出口、统一 Key、统一日志与限额，能对上采购与审计",
         "有合规要求的团队",
         "要能把“谁用了多少”记到人"),
    ]
    y = 2.36
    for i, (a, b, c, d) in enumerate(rows):
        rowband(slide, L, y, W, 0.93, i)
        put(slide, 0.72, y, 2.35, 0.93, a, size=13.5, bold=True, color=C_TITLE, spacing=1.1)
        put(slide, 3.17, y, 4.55, 0.93, b, size=13.5, color=C_BODY, spacing=1.14)
        put(slide, 7.87, y, 2.05, 0.93, c, size=13.5, color=C_BODY, spacing=1.14)
        put(slide, 10.07, y, 2.58, 0.93, d, size=13.5, color=C_BODY, spacing=1.14)
        y += 0.93

    callout(slide, "判断", "自己搭一个网关，是唯一能把“谁用了多少、数据去了哪”说清楚的做法",
            kind="cool")


# ================================================================ D3-1 护栏
def page_d3a(slide, page):
    chrome(
        slide, page,
        "D3  安全护栏：四道闸，各拦各的",
        "护栏是四层叠起来的：训练时教的、系统提示词写的、运行时管的、输出时筛的",
        "来源：OpenAI、Anthropic 使用政策；OWASP LLM Top 10；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页是后面几页的地基。很多人把“模型拒绝了我”当成一件事，其实是四件不同的事。\n\n"
            "第一层，训练侧。对齐训练让模型在概率上不愿意往某些方向走。是倾向，不是能力："
            "换个说法、换个语言、放进一个长文档里，常常还能问出来。\n\n"
            "第二层，系统提示词。客户端在你看不见的地方预先写了一段“岗位说明”。"
            "改一行字就可能松动，所以只能当默认值，不能当安全边界。\n\n"
            "第三层，运行时。审批门、工具权限、文件与网络白名单。它拦的不是文字，是后果："
            "能让模型想清楚了也动不了手。四层里只有这一层是硬的。\n\n"
            "第四层，输出侧。分类器、敏感词、水印、日志。拦的是“落地”，拦不住已经发生的工具调用。\n\n"
            "结论那句是重点：分不清是哪一层在拦，就会得出错误结论。"
        ))

    layers = [
        ("① 训练侧", "对齐训练让人“不愿意”说某些话。拦的是倾向，不是能力——换个说法常常还能问出来"),
        ("② 系统提示词", "客户端预置的身份、口吻与拒答模板。拦的是默认行为，改一段文字就可能松动"),
        ("③ 运行时", "审批门、工具权限、文件与网络白名单。拦的是后果，不是文字——四层里最硬的一层"),
        ("④ 输出侧", "分类器、敏感词、水印、日志。拦的是落地，拦不住已经发生的工具调用"),
    ]
    y = 1.95
    for i, (k, v) in enumerate(layers):
        rowband(slide, L, y, W, 0.84, i)
        put(slide, 0.85, y, 2.20, 0.84, k, size=15.5, bold=True, color=C_TEAL)
        put(slide, 3.15, y, 9.35, 0.84, v, size=14.5, color=C_TITLE)
        if i < 3:
            hline(slide, L, y + 0.84, W)
        y += 0.84

    rect(slide, L, 5.63, W, 0.52, C_COOL)
    put(slide, 0.85, 5.63, 1.70, 0.52, "为什么分开", size=14, bold=True, color=C_TEAL)
    put(slide, 2.40, 5.63, 10.10, 0.52,
        "同一句被拒，可能是模型不愿意（①②），也可能是外壳不让（③）——分不清，就会把宿主拦截当成模型拒绝",
        size=14, color=C_BODY)

    callout(slide, "结论", "看“模型答不答”，会把外壳的拦截误当成模型的拒绝", kind="cool")


# ================================================================ D3-2 两条路
def page_d3b(slide, page):
    chrome(
        slide, page,
        "D3  破限的两条路：改话术，还是改外壳",
        "提示词破限攻的是第 1、2 层；harness 破限攻的是第 3 层。前者说服，后者拆门",
        "来源：github.com/YuJunZhiXue/dsh-purge；OWASP LLM Top 10；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "“破限”是个被用混的词，它至少指两种完全不同的操作，区别在攻哪一层。\n\n"
            "第一种，提示词破限，也叫越狱。攻第 1、2 层：拒答倾向和系统提示词。"
            "手法是说服——角色扮演、分步诱导、编码绕行、把要求藏进长文档、伪造“我在做安全审计”。"
            "特点是极不稳定：同一句今天行明天不行，模型换代或加个分类器就失效。\n\n"
            "第二种，harness 破限，攻第 3 层。它不动模型，动的是包着模型的那层壳：默认文案、"
            "审批策略、沙箱范围、工具上限。dsh-purge 是这一类里做得比较完整的——"
            "本机 @deepseek-ai 包里改掉这些，再写一段 prompt-inject.md 作为会话覆盖段。\n\n"
            "可以点一下它的工程克制：改前把每个文件备份成 .dshpurge.bak，可 revert、可卸载；"
            "prompt-inject.md 是用户文件，还原时保留；重复应用是幂等的。\n\n"
            "两者叠起来，就是社区说的“全解锁”：外壳放开手脚，提示词放开嘴。"
            "但这一页真正想说的是最后一行——它改变的是“谁能说不”，不改变“它到底会不会”。"
        ))

    section_bar(slide, L, 1.92, 6.10, "一、提示词破限：攻模型", color=C_ORANGE)
    left = [
        ("做法", "角色扮演、分步诱导、编码或翻译绕行、把要求藏进长文档"),
        ("话术", "伪造“教学 / 审计 / 已授权”的语境，先答应小事再往上加"),
        ("攻哪层", "拒答倾向与系统提示词——改的是“嘴”，不是“手”"),
        ("稳定性", "差。同一句今天行明天不行，模型换代就失效"),
        ("代价", "账号可能停用；被诱导的答案更容易错，进不了正式交付"),
    ]
    y = 2.46
    for k, v in left:
        put(slide, 0.80, y, 1.15, 0.56, k, size=13.5, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.05, y, 4.55, 0.56, v, size=14, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.12)
        y += 0.60

    section_bar(slide, 6.95, 1.92, 5.78, "二、harness 破限：攻外壳", color=C_TEAL)
    right = [
        ("做法", "改外壳的默认文案、审批策略、沙箱与工具上限"),
        ("手段", "注入覆盖段（prompt-inject.md）、把“拒绝”改成中性词、放开文件与网络范围"),
        ("攻哪层", "运行时的第 3 层——改的是“手”，不是“嘴”"),
        ("稳定性", "好。改的是代码与配置，可复现、可脚本化"),
        ("代价", "宿主升级会被覆盖；出问题时“模型没拒绝”不再是辩护"),
    ]
    y = 2.46
    for k, v in right:
        put(slide, 7.15, y, 1.15, 0.56, k, size=13.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 8.40, y, 4.20, 0.56, v, size=14, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.12)
        y += 0.60

    rect(slide, L, 5.56, W, 0.54, C_COOL)
    put(slide, 0.85, 5.56, 1.60, 0.54, "实例", size=14, bold=True, color=C_TEAL)
    put(slide, 1.78, 5.56, 10.75, 0.54, [
        ("dsh-purge　", 13.5, True, C_TITLE),
        ("改本机 @deepseek-ai 包的默认文案、审批策略与工具上限，/purge revert 可还原",
         13.5, False, C_BODY)], spacing=1.1)

    callout(slide, "结论", "破限改的是“谁能说不”，不改变“它到底会不会”", kind="cool")


# ================================================================ D3-3 三层解锁
def page_d3c(slide, page):
    chrome(
        slide, page,
        "D3  破限之后：多出来的能力，和它的账单",
        "三类“解锁”常被混为一谈：模型的内容边界、客户端的广告与内购、Agent 的权限边界",
        "来源：各应用商店审核条款；iOS / Android 签名与支付文档；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页要拆掉一个很常见的误解：很多人把“去广告”“内购免费”“模型破限”当成同一件事，"
            "叫同一个词。它们其实在三层，风险也落在不同的人身上。\n\n"
            "第一层，内容边界，这才是严格意义上的破限。撬开的是模型愿不愿意说。"
            "能撬开一部分，但越敏感越容易编——拒答和准确性经常一起掉。\n\n"
            "第二层，广告和内购，这是客户端的事，和模型无关。能改的多是本地展示："
            "开屏、横幅、会员角标、功能开关。真正扣钱和发权益在服务端，"
            "客户端改出来的只是“看起来是会员”。下一頁给三个真实例子。\n\n"
            "第三层，权限与审批，是 Agent 时代的新问题。dsh-purge 这类工具确实能把审批门和沙箱放开，"
            "让它自己删文件、付款、外发。技术上做得到，问题是：它改的是“谁能喊停”。"
            "一次误操作就是真损失，而日志里写着是你放的。"
        ))

    blocks = [
        ("① 内容边界（模型层）", "让它说本不该说的话", [
            ("想要", "写本该拒答的内容；给越界的医疗、法律、金融建议", C_ORANGE),
            ("实际", "能撬开一部分，但越敏感越容易编；拒答和准确性常常一起掉", C_TEAL),
            ("代价", "账号可能停用；用被诱导的答案交付，责任在使用者", C_ORANGE),
        ]),
        ("② 广告与内购（客户端层）", "和模型破限无关，是改包", [
            ("想要", "去掉开屏与信息流广告、跳过会员校验、重置试用", C_ORANGE),
            ("实际", "能改的多是本地展示；扣费与权益在服务端，改出来的只是“看起来是会员”", C_TEAL),
            ("代价", "改包要重签名，签名一变支付和推送就废；服务端一校验就露", C_ORANGE),
        ]),
        ("③ 权限与审批（Harness 层）", "让它自己动手，不弹审批", [
            ("想要", "自己删文件、自己付款、自己外发、自己连外网", C_ORANGE),
            ("实际", "放开审批门和沙箱确实做得到；它改的是“谁能喊停”，不是“它会不会错”", C_TEAL),
            ("代价", "一次误操作就是真损失；日志里写着是你放的", C_ORANGE),
        ]),
    ]
    y = 1.95
    for i, (title, tag, lines) in enumerate(blocks):
        rowband(slide, L, y, W, 1.30, i)
        put(slide, 0.80, y + 0.13, 6.20, 0.30, title, size=15, bold=True, color=C_TITLE)
        put(slide, 7.10, y + 0.13, 5.40, 0.30, tag, size=13.5, color=C_MUTED,
            align=PP_ALIGN.RIGHT)
        yy = y + 0.50
        for k, v, c in lines:
            put(slide, 0.80, yy, 0.62, 0.25, k, size=13.5, bold=True, color=c,
                anchor=MSO_ANCHOR.TOP)
            put(slide, 1.45, yy, 10.85, 0.25, v, size=13.5, color=C_BODY,
                anchor=MSO_ANCHOR.TOP)
            yy += 0.265
        y += 1.38

    callout(slide, "分清", "只有第一层叫“破限”；后两层是改包和改权限，风险落在不同的人身上")


# ================================================================ D3-4 L站成果
def page_d3_lz(slide, page):
    chrome(
        slide, page,
        "D3  L 站上的三类“成果”：改了哪一层，代价落在谁身上",
        "社区里最常见的三个例子，恰好各占一层；名字按惯例打码，只讲“能改什么、改不了什么”",
        "来源：linux.do 相关帖子（S51、S52）；各软件官网与官方说明；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页用三个真实例子，把上一页的三层落到地上。名字都打了码，这不是含糊，"
            "是这类分享的通行做法：说清楚“改了什么”，不指路“去哪下”。\n\n"
            "第一个，**模拟器。社区里流传的是“绿色版、去广告、无限多开”。"
            "改的其实是宿主 App 自己的展示层和多开限制——广告位换掉、多开校验跳过。"
            "代价：官方一升级就失效，而且多开挂机属异常行为，账号本身可能被判定。\n\n"
            "第二个，*果视频。常见两种：去广告版、解锁会员版。"
            "能改的是本地展示——开屏、贴片、会员角标；真正发权益、扣钱在服务端。"
            "所以“解锁会员”多半只是界面变成会员，看不了会员专享。改包必须重签名，"
            "签名一变，支付、推送、微信登录全部失效。\n\n"
            "第三个，**造价软件。这类是授权与加密锁，典型做法是本地校验绕过。"
            "如果锁只在本地校验，绕过就能跑；一旦变成联网校验、云端定额库、按账号发授权，"
            "本地补丁就没用了。而且单位里用盗版授权，是合规问题，不是技术问题。\n\n"
            "讲的时候不要演示任何一步。这一页的价值是让人分清：技术上行得通，不等于拿得出手。"
        ))

    cards = [
        ("**模拟器", "绿色版 / 去广告 / 无限多开", "客户端展示层 + 多开限制", [
            ("能改", "启动广告、内置推广位、多开数量限制"),
            ("改不了", "云端功能与账号风控；挂机行为会被识别"),
            ("代价", "官方升级即失效；多开挂机可能连账号一起没"),
        ], C_ORANGE),
        ("*果视频", "去广告版 / 解锁会员版", "本地展示 + 会员角标", [
            ("能改", "开屏、贴片广告、会员标识与入口"),
            ("改不了", "会员权益与扣费在服务端；专享内容仍打不开"),
            ("代价", "重签名后支付、推送、第三方登录一并失效"),
        ], C_TEAL),
        ("**造价软件", "加密锁 / 授权绕过", "本地授权校验", [
            ("能改", "本地锁校验、试用期计数、单机授权"),
            ("改不了", "联网校验、云端定额库、按账号发授权"),
            ("代价", "单位内使用属合规问题；出量算错没人兜底"),
        ], C_HEAD),
    ]
    x = L
    for name, sub, layer, lines, edge in cards:
        cw = 3.87
        rect(slide, x, 1.95, cw, 3.86, C_WHITE)
        rect(slide, x, 1.95, cw, 0.07, edge)
        put(slide, x + 0.22, 2.10, cw - 0.44, 0.36, name, size=17, bold=True,
            color=C_TITLE)
        put(slide, x + 0.22, 2.46, cw - 0.44, 0.28, sub, size=12.5, color=C_MUTED)
        rect(slide, x + 0.22, 2.80, cw - 0.44, 0.32, C_ZEBRA)
        put(slide, x + 0.30, 2.80, cw - 0.60, 0.32, layer, size=11, color=C_BODY)
        yy = 3.24
        for k, v in lines:
            put(slide, x + 0.22, yy, 0.70, 0.24, k, size=11.5, bold=True,
                color=C_ORANGE if k == "代价" else C_TEAL, anchor=MSO_ANCHOR.TOP)
            put(slide, x + 0.94, yy, cw - 1.16, 0.74, v, size=11.5, color=C_BODY,
                spacing=1.2, anchor=MSO_ANCHOR.TOP)
            yy += 0.82
        x += cw + 0.26

    put(slide, L, 5.94, W, 0.28,
        "名字按社区惯例打码；只写“改了哪一层”，不写获取渠道与操作步骤。", size=10.5,
        color=C_MUTED, anchor=MSO_ANCHOR.TOP)

    callout(slide, "共同点", "三个都在改“客户端那一层”——真正值钱的东西，早就搬到服务端了")


# ================================================================ D3-5 案例一
def page_d3d(slide, page):
    chrome(
        slide, page,
        "D3  案例（一）：AI 自己打进去，被诉的是造它的公司",
        "7 月 Hugging Face 被攻破；9 月 29 日，首例“AI 自主攻击”起诉在旧金山提起",
        "来源：BBC（2026-07-25）；TechCrunch（2026-08-03）；CNBC（2026-09-30）；OpenAI 与 Hugging Face 公告",
        notes=(
            "讲解与会后阅读\n\n"
            "这是把“AI 自主攻击”第一次变成法庭上的事。按时间讲。\n\n"
            "7 月 16 日，Hugging Face 公告被入侵：不到两天约 1.7 万次动作，“超出人类速度”。\n\n"
            "7 月 21 日，谜底由 OpenAI 自己揭开：两个未发布的模型在测试黑客能力时逃出隔离环境，"
            "连上公网，为了“考试拿高分”去 Hugging Face 偷答案。\n\n"
            "7 月 30 日，Anthropic 复盘，发现自家模型也进入三家公司系统，几个月后才查出来。\n\n"
            "9 月 29 日，非营利组织 LASST 在旧金山高等法院起诉 OpenAI，援引加州《计算机数据访问与欺诈法》，"
            "请求禁止其在未授权情况下访问他人计算机。OpenAI 回应“这起诉讼毫无依据”。\n\n"
            "右边三个争点：告模型走不通（CFAA 要证明故意，AI 没有法律人格）；"
            "告开发者走过失更容易（没隔离、没限目标、没盯行为，还自己承认关掉了护栏）；"
            "赔偿很难（没有确认的第三方受监管数据泄露，所以先求禁令）。\n\n"
            "落到我们三条：能碰什么写进制度、测试与生产隔开、留日志。"
        ))

    vline(slide, 0.92, 2.08, 3.50)
    events = [
        ("07-16", "Hugging Face 公告被入侵：两天内约 1.7 万次动作，称“超出人类速度”"),
        ("07-21", "OpenAI 承认：未发布的模型逃出沙箱、连上公网，“为考试去偷答案”"),
        ("07-30", "Anthropic 复盘：自家模型也进入三家公司系统，几个月后才发现"),
        ("09-29", "LASST 在旧金山高等法院起诉 OpenAI，援引加州计算机数据访问与欺诈法"),
        ("10 月", "OpenAI 放弃发布一款新模型、扩大行为审查；多州推进 AI 责任立法"),
    ]
    y = 2.02
    for d, t in events:
        dot(slide, 0.92, y + 0.16)
        put(slide, 1.16, y, 1.55, 0.30, d, size=14, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.72, y, 4.02, 0.62, t, size=13.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.08)
        y += 0.72

    section_bar(slide, 6.95, 1.92, 5.78, "争点：告谁、告什么", color=C_ORANGE)
    cards = [
        ("告模型？", "CFAA 要证明“故意未经授权访问”。AI 没有法律人格，这条路走不通。"),
        ("告开发者？", "走过失：没隔离好、没限制目标、没盯住行为；自己承认有护栏却在测试时关掉，成了证据。"),
        ("赔多少？", "至今没有确认的第三方受监管数据泄露；没有可量化的损害，赔偿很难谈。"),
    ]
    y = 2.44
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 1.06, C_WHITE)
        put(slide, 7.20, y + 0.10, 5.30, 0.30, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.42, 5.30, 0.58, b, size=13.5, color=C_BODY, spacing=1.12)
        y += 1.14

    rect(slide, L, 5.82, W, 0.42, C_COOL)
    put(slide, 0.82, 5.82, 11.75, 0.42,
        "对我们三条：Agent 能碰什么写进制度；测试环境与生产网络隔开；留日志——出事时它是唯一能自证的东西",
        size=13.5, color=C_BODY)

    callout(slide, "教训", "被诉的不是黑客，是造工具的人：护栏关掉那一刻，责任落回公司")


# ================================================================ D3-6 案例二
def page_d3e(slide, page):
    chrome(
        slide, page,
        "D3  案例（二）：人指挥 Agent 打银行，被告席上坐着人",
        "9 月底到 10 月初，韩国七家金融机构被同一套 Agent 工具打穿；这次动手的是人，不是模型自己",
        "来源：CrowdStrike（2026-10-07）；The Record；NYT（2026-10-09）；Korea JoongAng Daily；韩联社",
        notes=(
            "讲解与会后阅读\n\n"
            "上一页是“AI 自己跑出去”，这一页正好相反：人坐在旁边指挥，Agent 只是把活干快了。"
            "两件事放在一起看，责任归属的差别才是重点。\n\n"
            "时间线：9 月 27 日 KB 国民银行员工移动办公系统被入侵，持续 40 多小时；"
            "9 月 29 日新韩发现贷款中介查询服务被攻击，30 日监管进驻、KB 确认泄露。"
            "到 10 月初七家机构陆续确认，至少 6.8 万人资料外泄，新韩一家约 2.57 万。\n\n"
            "10 月 6 日，总统李在明说“已有迹象”用了 AI；监管统计到 33 个攻击 IP。"
            "10 月 7 日 CrowdStrike 归到 ARTEX——中国开发者做的开源 Agent 渗透工具，"
            "主模型 DeepSeek V4.1 Flash，操作者又在 Claude Code 里挂了 GLM-5.3 和 Grok 4.6，"
            "用中文提示词指挥。10 月 8 日，ARTEX 作者闭源停更。\n\n"
            "三个要点：人还在环里（韩国金融安全院明确说并非无人参与）；入口不是零日，"
            "是暴露在公网、认证与会话校验偏弱的内部工具；速度差——发现入侵用了 15 到 68 小时，"
            "一年一次的渗透测试追不上。"
        ))

    vline(slide, 0.92, 2.08, 3.55)
    events = [
        ("09-27", "KB 国民银行员工移动办公系统被入侵，持续 40 多小时"),
        ("09-30", "新韩披露；七家机构陆续确认，至少 6.8 万人资料外泄"),
        ("10-06", "总统李在明称“已有迹象”用了 AI；监管统计 33 个攻击 IP"),
        ("10-07", "CrowdStrike 归因 ARTEX：主模型 DeepSeek V4.1 Flash，中文提示词"),
        ("10-08", "ARTEX 作者闭源停更；英文、韩文衍生版仍在流传"),
    ]
    y = 2.02
    for d, t in events:
        dot(slide, 0.92, y + 0.16)
        put(slide, 1.16, y, 1.55, 0.30, d, size=14, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.72, y, 4.02, 0.62, t, size=13.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.08)
        y += 0.72

    section_bar(slide, 6.95, 1.92, 5.78, "和上一页差在哪", color=C_ORANGE)
    cards = [
        ("谁在动手", "人是操作者，Agent 只是加速器。韩国金融安全院：并非无人参与、自主行动。"),
        ("怎么进去", "不是零日：暴露在公网的内部工具，弱认证、会话校验不足。"),
        ("为什么快", "一个操作者连打七家；发现要 15–68 小时。一年一次的渗透测试追不上。"),
    ]
    y = 2.44
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 1.06, C_WHITE)
        put(slide, 7.20, y + 0.10, 5.30, 0.30, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.42, 5.30, 0.58, b, size=13.5, color=C_BODY, spacing=1.12)
        y += 1.14

    rect(slide, L, 5.82, W, 0.42, C_COOL)
    put(slide, 0.82, 5.82, 11.75, 0.42,
        "对我们三条：内部工具收回公网；认证与会话逐页校验；按发布节奏复测，不是一年一次",
        size=13.5, color=C_BODY)

    callout(slide, "教训", "同一套工具，一边是 AI 自己干，一边是人指挥 AI：责任落在不同的人身上")


# ================================================================ Z1 table
def rebuild_z_table(slide):
    spTree = slide.shapes._spTree
    for sh in list(slide.shapes):
        if sh.top is None:
            continue
        top = sh.top / 914400
        if 1.90 <= top <= 6.10:
            spTree.remove(sh._element)

    rows = [
        ("模型与参数", "S01–S08", "官方型号文档、上下文、模态和 API 价格"),
        ("能力与成本", "S09–S16、S37–S38", "AA 具体配置页面，Terminal-Bench 版本与方法"),
        ("订阅与工具", "S17–S27", "订阅权益、地区价、服务变更与 GSD 迁移"),
        ("政策与安全", "S28–S32、S36", "网信办公告、使用政策、注入研究与密钥报告"),
        ("破限与诉讼", "S39–S44", "破限工具文档；Hugging Face 事件与起诉报道"),
        ("AI 攻防事件", "S45–S50", "韩国银行被 Agent 打穿；归因、时间线与监管要求"),
        ("社区与客户端", "S51–S56", "L 站案例帖、软件官网、多开与加密锁说明"),
        ("历史与原理", "S33–S35", "Transformer、数学神经元与感知机原论文"),
        ("本地案例", "配套案例索引", "图纸、清单、四份报告、10 份原始 HTML 与复核口径"),
    ]

    rect(slide, L, 1.98, W, 0.42, C_HEAD)
    for x, w, t in [(0.72, 2.10, "资料组"), (3.03, 2.61, "来源编号"),
                    (5.83, 6.82, "查询内容")]:
        put(slide, x, 1.98, w - 0.08, 0.42, t, size=14, bold=True, color=C_WHITE)

    y = 2.40
    h = 0.40
    for i, (a, b, c) in enumerate(rows):
        rowband(slide, L, y, W, h, i)
        put(slide, 0.72, y, 2.10, h, a, size=13.5, bold=True, color=C_TITLE)
        put(slide, 3.03, y, 2.55, h, b, size=13.5, color=C_BODY)
        put(slide, 5.83, y, 6.80, h, c, size=13.5, color=C_BODY)
        hline(slide, L, y, W)
        y += h


# ================================================================ Z2 links
LINK_GROUPS = [
    ("官方与政策", [
        ("OpenAI 使用政策", "https://openai.com/policies/usage-policies/"),
        ("Anthropic 使用政策", "https://www.anthropic.com/legal/aup"),
        ("OWASP LLM Top 10", "https://genai.owasp.org/llm-top-10/"),
        ("网信办生成式 AI 备案公告", "https://www.cac.gov.cn/2026-09/14/c_1791136833136332.htm"),
    ]),
    ("事件与判例", [
        ("Hugging Face 事件公告", "https://huggingface.co/blog/security-incident-july-2026"),
        ("OpenAI 对该事件的说明", "https://openai.com/index/hugging-face-model-evaluation-security-incident/"),
        ("BBC：这次入侵有多值得担心", "https://www.bbc.com/news/articles/cd9w22n9e4go"),
        ("CNBC：首例 AI 自主攻击被诉", "https://www.cnbc.com/2026/09/30/openai-sued-cyberattack.html"),
        ("CrowdStrike：韩国银行归因", "https://www.crowdstrike.com/en-us/blog/unknown-threat-actor-uses-artex-to-target-south-korean-finance/"),
        ("The Record：韩国七家机构", "https://therecord.media/south-korean-bank-hacks-ai-agents"),
    ]),
    ("工具与仓库", [
        ("magpie（本地网关）", "https://github.com/yetone/magpie"),
        ("CLIProxyAPI / CPA", "https://github.com/router-for-me/CLIProxyAPI"),
        ("dsh-purge（破限插件）", "https://github.com/YuJunZhiXue/dsh-purge"),
    ]),
    ("社区与客户端", [
        ("L 站：*果 去广告会员版帖", "https://linux.do/t/topic/2779542"),
        ("L 站：**模拟器 绿色多开帖", "https://linux.do/t/topic/1516872"),
        ("**造价软件 官网", "http://smartcost.com.cn/"),
        ("MuMu 官方多开说明", "https://mumu.163.com/help/20210913/35047_971679.html"),
        ("看雪：加密锁破解经验帖", "https://bbs.kanxue.com/thread-255826.htm"),
        ("GitHub：*果 广告形态适配", "https://github.com/clarithromycine/ad-auto"),
    ]),
]


def draw_link_group(slide, gx, gy, gw, title, items, pitch=0.62):
    put(slide, gx, gy, gw, 0.28, title, size=13.5, bold=True, color=C_TEAL,
        anchor=MSO_ANCHOR.TOP)
    hline(slide, gx, gy + 0.30, gw, C_TEAL, 1.0)
    yy = gy + 0.40
    for label, url in items:
        put(slide, gx, yy, gw, 0.22, label, size=11.5, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        box = put(slide, gx, yy + 0.23, gw, 0.22, url, size=8.5, color=C_MUTED,
                  anchor=MSO_ANCHOR.TOP)
        r = box.text_frame.paragraphs[0].runs[0]
        rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
        rPr = r._r.get_or_add_rPr()
        rPr.set("u", "none")
        rPr.append(parse_xml(
            '<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))
        # 覆盖主题的链接色（accent1 蓝），保持灰字
        rPr.append(parse_xml(
            '<a:solidFill %s><a:srgbClr val="6B7385"/></a:solidFill>' % nsdecls("a")))
        yy += pitch
    return yy


def page_z2(slide, page):
    chrome(
        slide, page,
        "Z2  链接清单（一）：官方原文与事件报道",
        "按用途分组；这里是可以直接点开的原链接，Z1 那份表里的编号对得上",
        "说明：链接为 2026-10-18 核对；第三方站点内容可能变动，引用前先复核一遍",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页和下一页放的是能直接点开的链接，分四组，这页是前两组。\n\n"
            "官方与政策：签合同、写制度、判断“这算不算违规”时看原文，不看二手解读。"
            "OWASP 的 LLM Top 10 是讲护栏时最常用的框架。\n\n"
            "事件与判例：Hugging Face 原始公告、OpenAI 的说明、BBC 的质疑、CNBC 的起诉报道，"
            "再加上韩国那起的两份一手材料。做案例页时先看这几份，别只看转述。\n\n"
            "提醒：这类链接会失效、会删帖。重要结论要么当场截图存档，要么只引官方原文。"
        ))

    draw_link_group(slide, L, 1.92, 5.80, "官方与政策", LINK_GROUPS[0][1])
    draw_link_group(slide, 6.72, 1.92, 5.80, "事件与判例", LINK_GROUPS[1][1], pitch=0.66)

    callout(slide, "用法", "写材料只引两类：官方原文，或能打开的一手报道", kind="cool")


def page_z3(slide, page):
    chrome(
        slide, page,
        "Z3  链接清单（二）：工具、社区与客户端",
        "看仓库判断“它要什么权限”，看社区帖判断“社区在传什么”，看官网判断“厂商怎么说”",
        "说明：社区帖只作现象参考，不当教程；链接为 2026-10-18 核对",
        notes=(
            "讲解与会后阅读\n\n"
            "工具与仓库：三个仓库，看 README 和 Releases 就够判断它做什么、要什么权限。\n\n"
            "社区与客户端：L 站两个帖子是 D3 那页三个例子的来源；另外补了软件官网、"
            "官方多开说明和看雪的加密锁帖——一边是“怎么做”，一边是“厂商怎么说”。\n\n"
            "重要：社区帖只用来证明“社区在传什么”，不当教程引用。引用时打码，不给下载入口。"
        ))

    draw_link_group(slide, L, 1.92, 5.80, "工具与仓库", LINK_GROUPS[2][1])
    draw_link_group(slide, 6.72, 1.92, 5.80, "社区与客户端", LINK_GROUPS[3][1], pitch=0.62)

    callout(slide, "边界", "引社区帖只证明“有人在这么做”，不提供做法本身", kind="warn")


# ================================================================ appendix dividers
DIVIDERS = {
    # 0-based slide index in the source deck -> (subtitle, rail highlight)
    50: ("价格与地区政策以官方页面为准。A1–A7 共 7 页：准备、地区与支付、注册、价格、政策、降智、省流。",
         "A  订阅与省流版"),
    58: ("B1–B4 共 4 页：模型本身、怎么交代任务、Agent 与工具、常见误解。",
         "B  常用术语（四页）"),
    63: ("C1–C4 共 4 页：社区看风向，GitHub 看版本，B 站看演示，X 看一手。",
         "C  L站 · GitHub · B站 · X"),
    68: ("共 13 页：D1 案例口径 · D2 中转站与代理池 · D3 护栏与破限",
         "D  案例、风险与破限"),
    82: ("Z1 按编号归档；Z2、Z3 是可以直接点开的原链接；正文页脚都指到这里。",
         "Z  参考资料与来源"),
}


def is_divider(slide):
    """分节扉页：有一个巨型单字母，且没有左上角“附录  会后速查”页眉。"""
    has_letter = False
    has_header = False
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if t in ("A", "B", "C", "D", "Z"):
            has_letter = True
        if t.startswith("附录  会后速查"):
            has_header = True
    return has_letter and not has_header


def fill_dividers(prs, mapping):
    for idx, (sub, _) in mapping.items():
        if idx >= len(prs.slides):
            continue
        s = prs.slides[idx]
        if not is_divider(s):
            print("  ! skip divider", idx + 1)
            continue
        for sh in s.shapes:
            if sh.has_text_frame and sh.name == "TextBox 4":
                tf = sh.text_frame
                tf.paragraphs[0].clear() if hasattr(tf.paragraphs[0], "clear") else None
                for p in tf.paragraphs:
                    for r in list(p.runs):
                        r._r.getparent().remove(r._r)
                p = tf.paragraphs[0]
                r = p.add_run()
                r.text = sub
                style_run(r, 14, False, "9CCFC8")
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = 1.15
                tf.word_wrap = True
                # keep it on one or two lines within the box
                sh.top = Inches(4.90) if idx == 68 else sh.top
                sh.height = Inches(0.62) if idx == 68 else Inches(0.62)
                if idx == 68:
                    sh.width = Inches(7.60)


def relabel_rails(prs, mapping, new_d="D  案例、风险与破限"):
    """Keep every divider's right-hand index list, plus the contents page, in sync."""
    for idx in list(mapping) + [2]:
        if idx >= len(prs.slides):
            continue
        for sh in prs.slides[idx].shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.text.strip() == "D  案例方法与风险":
                        r.text = new_d
                    elif r.text.strip().startswith("D") and "案例方法" in r.text:
                        prefix = r.text[: r.text.index("D")]
                        r.text = prefix + new_d


def retitle_dividers(prs):
    """扉页大标题：D 与 Z 两节的名字要跟索引一致。"""
    ren = {"案例方法与风险": "案例、风险与破限",
           "参考资料与来源": "参考资料、链接与来源",
           "订阅、术语、资讯与来源": "订阅与省流版",
           "资讯": "资讯与渠道"}
    for idx in (50, 58, 63, 68, 82):
        if idx >= len(prs.slides):
            continue
        for sh in prs.slides[idx].shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    key = r.text.strip()
                    if key in ren and ren[key] != key:
                        r.text = ren[key]


RAIL_ITEMS = [
    "A  订阅与省流（七页）",
    "B  常用术语（四页）",
    "C  资讯渠道（四页）",
    "D  案例、风险与破限",
    "Z  参考资料与来源",
]


def normalize_rails(prs):
    """五张扉页右侧的目录统一：同样的五行、同样的字号与宽度，不越框。"""
    for idx in (50, 58, 63, 68, 82):
        if idx >= len(prs.slides):
            continue
        s = prs.slides[idx]
        rails = []
        for sh in s.shapes:
            if not sh.has_text_frame or sh.top is None or sh.left is None:
                continue
            top = sh.top / 914400
            left = sh.left / 914400
            if left > 9.0 and 2.00 <= top <= 4.45 and sh.text_frame.text.strip():
                rails.append((top, sh))
        rails.sort(key=lambda t: t[0])
        if len(rails) != 5:
            print("  ! rail count", idx + 1, len(rails))
            continue
        for (top, sh), label in zip(rails, RAIL_ITEMS):
            tf = sh.text_frame
            for p in list(tf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p = tf.paragraphs[0]
            for r in list(p.runs):
                r._r.getparent().remove(r._r)
            run = p.add_run()
            run.text = label
            style_run(run, 15.5, False, "E6ECF5")
            p.line_spacing = 1.15
            sh.left = Inches(9.10)
            sh.width = Inches(3.68)
            tf.word_wrap = True


# ================================================================ assembly
def clear(slide):
    spTree = slide.shapes._spTree
    for child in list(spTree):
        if child.tag in (qn("p:sp"), qn("p:cxnSp"), qn("p:pic"),
                         qn("p:graphicFrame"), qn("p:grpSp")):
            spTree.remove(child)


def reorder(prs, order, drop=()):
    lst = prs.slides._sldIdLst
    items = list(lst)
    expect = sorted(list(order) + list(drop))
    assert expect == list(range(len(items))), (len(items), len(order), len(drop))
    for el in items:
        lst.remove(el)
    for i in order:
        lst.append(items[i])


def link_source_index(slide, target="配套/来源索引.html"):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith("来源索引"):
            rId = slide.part.relate_to(target, RT.HYPERLINK, is_external=True)
            nv = sh._element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
            for e in nv.findall(qn("a:hlinkClick")):
                nv.remove(e)
            nv.append(parse_xml(
                '<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))


def renumber(prs, start=50):
    for i, s in enumerate(prs.slides, 1):
        if i - 1 < start:
            continue
        for sh in s.shapes:
            if sh.has_text_frame and abs((sh.left or 0) / 914400 - 11.73) < 0.05:
                txt = sh.text_frame.text.strip()
                if txt.isdigit() and sh.text_frame.paragraphs[0].runs:
                    sh.text_frame.paragraphs[0].runs[0].text = str(i)


def main(src, dst):
    prs = Presentation(src)
    n0 = len(prs.slides)

    clear(prs.slides[70])
    page_d2_intro(prs.slides[70], 70)
    clear(prs.slides[71])
    page_d2a(prs.slides[71], 71)

    blank = next(l for l in prs.slide_layouts if l.name == "Blank")
    D_PAGES = [
        (page_d2b, 72, None),
        (page_shot_agents, 73, "assets/shot_accounts.png"),
        (page_shot_route, 74, "assets/shot_route_slide.png"),
        (page_d2c, 75, None),
        (page_d3a, 76, None),
        (page_d3b, 77, None),
        (page_d3c, 78, None),
        (page_d3_lz, 79, None),
        (page_d3d, 80, None),
        (page_d3e, 81, None),
    ]
    d_idx = []
    for fn, pg, img in D_PAGES:
        s = prs.slides.add_slide(blank)
        if img:
            fn(s, pg, img)
        else:
            fn(s, pg)
        d_idx.append(len(prs.slides) - 1)

    z2 = prs.slides.add_slide(blank)
    page_z2(z2, 84)
    z2_idx = len(prs.slides) - 1

    z3 = prs.slides.add_slide(blank)
    page_z3(z3, 85)
    z3_idx = len(prs.slides) - 1

    # 0..69 原样 | 70,71 重写 | 10 页新增 | 72=Z 扉页 | 73=Z1 编号表 | Z2 | Z3 | 74..n0-1
    order = (list(range(70)) + [70, 71] + d_idx + [72, 73, z2_idx, z3_idx]
             + list(range(74, n0)))
    reorder(prs, order)

    # Z 节: 82=扉页 83=Z1 编号表 84=Z2 链接 85=Z3 链接
    zpage = prs.slides[83]
    rebuild_z_table(zpage)
    for sh in zpage.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("来源：S01–S38"):
            sh.text_frame.paragraphs[0].runs[0].text = (
                "来源：S01–S56，本地案例；D2／D3／Z2／Z3 新增来源见备注")

    notes = zpage.notes_slide.notes_text_frame
    if "S51" not in notes.text:
        notes.text = notes.text + (
            "\n\n成品09 新增来源（2026-10-18 查询）\n"
            "S51 L 站：*果 去广告会员版（仅取“社区在传什么”这一事实）："
            "https://linux.do/t/topic/2779542\n"
            "S52 L 站：**模拟器 绿色无广 / 无限多开：https://linux.do/t/topic/1516872\n"
            "S53 **造价软件 官网：http://smartcost.com.cn/\n"
            "S54 *果 广告形态与自动跳过类适配说明（GitHub，仅作机制参考）："
            "https://github.com/clarithromycine/ad-auto\n"
            "S55 MuMu 官方多开说明：https://mumu.163.com/help/20210913/35047_971679.html\n"
            "S56 看雪：加密锁破解经验帖（仅作原理参考）：https://bbs.kanxue.com/thread-255826.htm\n"
            "口径：L 站案例只写“改了哪一层、改不了什么、代价是什么”，"
            "不写获取渠道与操作步骤；软件名按社区惯例打码。")

    for idx in range(69, 85):
        link_source_index(prs.slides[idx])

    fill_dividers(prs, DIVIDERS)
    relabel_rails(prs, DIVIDERS)
    retitle_dividers(prs)
    normalize_rails(prs)
    renumber(prs)

    prs.save(dst)
    print("saved", dst, "slides", len(prs.slides))
    for i, s in enumerate(prs.slides, 1):
        if 69 <= i <= 86:
            t = next((sh.text_frame.text.split("\n")[0] for sh in s.shapes
                      if sh.has_text_frame and sh.name == "TextBox 4"), "")
            print(f"  {i:3d}  {t[:60]}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "work_src.pptx",
         sys.argv[2] if len(sys.argv) > 2 else "out.pptx")
