# -*- coding: utf-8 -*-
"""Rebuild the D2 / D3 appendix pages of AI交流分享-成品07 in the deck's own style."""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

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
    """content: str | list[str] | list[list[(text,size,bold,color)]]"""
    if isinstance(content, str):
        paras = [[(content, size, bold, color)]]
    elif content and isinstance(content[0], str):
        paras = [[(t, size, bold, color)] for t in content]
    elif content and isinstance(content[0], tuple):
        # one paragraph made of several runs
        paras = [list(content)]
    else:
        # list of paragraphs (each a list of runs)
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


def rect(slide, x, y, w, h, fill=None, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
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


# ================================================================ D2-1
def page_d2a(slide, page):
    chrome(
        slide, page,
        "D2  反向代理：请求多走的那一跳",
        "除了直连官方，最常见的是“发给第三方、由它转给厂商”：中转 API、反代站、共享号池",
        "来源：OpenAI、Anthropic 使用政策；CLIProxyAPI、magpie 文档；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一段讲三件事：反代是怎么走的、多这一跳多出什么、以及怎么自查。\n\n"
            "先说路径。直连官方只有一层关系：Key 是你的，账单是你的，数据直接从你的机器到厂商。"
            "所谓“反向代理”“中转 API”“反代站”，都是把这一层改成两层：你的客户端连的是第三方的域名，"
            "第三方再拿自己的 Key（或替你注册的账号）去连厂商。从客户端看一切照旧，区别只在于中间多了一个人。\n\n"
            "多这一跳，多出四件事：数据、密钥、模型、条款与资金。这四条不是理论风险，"
            "这类社区里每隔几周就能看到一次相关纠纷。\n\n"
            "最实用的是最后一行“自查三问”：Base URL 是谁的域名？Key 是谁开的？账单抬头是谁？"
            "三个问题里有两个答不上来，就不该往里放单位资料。\n\n"
            "注意措辞：这一页不评价任何具体服务商，只讲路径和风险形状。"
        ))

    rect(slide, L, 1.95, W, 0.46, C_WHITE)
    rect(slide, L, 1.95, 0.09, 0.46, C_TEAL)
    put(slide, 0.85, 1.95, 11.60, 0.46, [
        ("直连　", 14, True, C_TEAL),
        ("你的客户端 → 厂商 API　｜　Key 是你的，账单是你的，中间没有第三方", 14, False, C_BODY)])

    rect(slide, L, 2.46, W, 0.46, C_WARM)
    rect(slide, L, 2.46, 0.09, 0.46, C_ORANGE)
    put(slide, 0.85, 2.46, 11.60, 0.46, [
        ("反代　", 14, True, C_ORANGE),
        ("你的客户端 → 中转方 → 厂商 API　｜　域名是别人的，Key 由对方持有或代充", 14, False, C_BODY)])

    section_bar(slide, L, 3.10, W, "多这一跳，多出四件事")

    rows = [
        ("数据", "输入、附件、输出都要过对方服务器：可能被留存、被转卖，也可能进对方的训练语料"),
        ("密钥", "交出 Key 就交出计费权：对方能拿它刷别的模型或转卖，账单仍记在你头上"),
        ("模型", "便宜模型冒充贵模型：一句“我是 Opus”证明不了什么，只能靠指纹和固定题目"),
        ("条款与资金", "绕地区限制、共享 Key、多号刷额度都在停用条款里；预充值、无发票、跑路"),
    ]
    y = 3.54
    for i, (k, v) in enumerate(rows):
        rowband(slide, L, y, W, 0.50, i)
        put(slide, 0.80, y, 1.55, 0.50, k, size=15, bold=True, color=C_ORANGE)
        put(slide, 2.35, y, 10.15, 0.50, v, size=14.5, color=C_TITLE)
        y += 0.50

    rect(slide, L, 5.70, W, 0.50, C_COOL)
    put(slide, 0.80, 5.70, 1.60, 0.50, "自查三问", size=14, bold=True, color=C_TEAL)
    put(slide, 2.35, 5.70, 10.15, 0.50,
        "Base URL 是谁的域名？Key 是谁开的？账单抬头写的是谁？", size=14.5, color=C_BODY)

    callout(slide, "防护", "资料只走批准渠道；Key 单独开、设上限；统一出口用自建或采购的网关")


# ================================================================ D2-2
def page_d2b(slide, page):
    chrome(
        slide, page,
        "D2  账号池：一份订阅，几个人排队",
        "账号池不是更强的模型，而是把多个账号的登录态放进一个网关，谁有额度谁上",
        "来源：magpie 文档（usemagpie.ai）；CLIProxyAPI README；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "账号池解决的是一个很具体的问题：订阅按“一个人用”卖，额度按 5 小时和每周滚动，"
            "一个人白天用完了，晚上就没了。把几个账号放进同一个网关，就能一直有额度。\n\n"
            "四步：登录、收口、调度、计量。登录是最容易被忽略的一步——凭据是文件或钥匙串里的东西，"
            "把它复制走等于把账号复制走；改密码是唯一能收回的办法。\n\n"
            "收口这一步，就是“反代”真正发生的地方：网关对外只暴露一个地址，Agent 配置里写的是 "
            "127.0.0.1 或一个内网 IP，看起来完全没有越过任何东西。\n\n"
            "右边三条是风险。重点讲第三条：厂商风控看的不是“你是谁”，是“形状对不对”——"
            "同一个 IP 上五个账号、并发超过一个人的打字速度、今天在东京明天在法兰克福，这些都会被记下来。\n\n"
            "收尾给一个诚实判断：自己有两三个订阅、或者团队共用一台机器，是常见做法；"
            "但“常见”不等于“条款允许”，尤其不要用单位的资料去跑。"
        ))

    section_bar(slide, L, 1.95, 6.10, "它怎么运转（四步）")
    steps = [
        ("① 登录", "每个账号在官方客户端登录一次，凭据留在本机"),
        ("② 收口", "本地网关对外给兼容接口，Agent 只认一个地址"),
        ("③ 调度", "轮询、按剩余额度排序、失败自动切下一个"),
        ("④ 计量", "5 小时与每周窗口、token 与花费分开统计"),
    ]
    vline(slide, 0.92, 2.62, 3.34)
    y = 2.52
    for k, v in steps:
        dot(slide, 0.92, y + 0.15)
        put(slide, 1.20, y, 5.30, 0.32, k, size=15, bold=True, color=C_TITLE)
        put(slide, 1.20, y + 0.36, 5.30, 0.34, v, size=14, color=C_BODY,
            spacing=1.1, anchor=MSO_ANCHOR.TOP)
        y += 0.86

    put(slide, 1.20, 5.92, 5.40, 0.30,
        "一个地址背后，可以有 1 个账号，也可以有 20 个", size=13.5, color=C_MUTED)

    section_bar(slide, 6.95, 1.95, 5.78, "它把什么变成了风险", color=C_ORANGE)
    cards = [
        ("一人一号", "订阅按“一个人使用”卖；多账号集中调度，通常在条款之外。"),
        ("登录态即身份", "凭据文件被复制走，等于账号被复制走——只能改密码收回。"),
        ("风控看的是形状", "同 IP 多号、并发超过一个人、地区漂移，都是可识别特征。"),
    ]
    y = 2.52
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 0.98, C_WHITE)
        put(slide, 7.20, y + 0.12, 5.30, 0.32, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.46, 5.30, 0.44, b, size=14, color=C_BODY,
            spacing=1.12, anchor=MSO_ANCHOR.TOP)
        y += 1.08

    rect(slide, 6.95, 5.78, 5.78, 0.54, C_COOL)
    put(slide, 7.18, 5.78, 5.40, 0.54,
        "相对合理的用法：自己的两三个订阅，或团队共用一台机器", size=13.5, color=C_BODY,
        spacing=1.1)

    callout(slide, "注意", "账号池把“额度不够”变成“身份不够”：省下的是钱，押上的是账号")


# ================================================================ D2-3
def page_d2c(slide, page):
    chrome(
        slide, page,
        "D2  工具：magpie、CPA 与自建网关",
        "magpie 在桌面，CPA（CLIProxyAPI）在服务器：能力接近，暴露面不同",
        "来源：github.com/yetone/magpie；github.com/router-for-me/CLIProxyAPI；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "两个工具都能做同一件事：把多个来源（订阅登录态、API Key、本地模型）收成一个兼容接口。"
            "区别在形态和暴露面。\n\n"
            "magpie 是桌面应用加本地网关，默认 127.0.0.1:3425，同时支持 OpenAI Chat / Responses、"
            "Anthropic Messages 和 Gemini 四种协议并互相转换。卖点是“一个 Agent 用任意模型”，"
            "改配置时原子改写、保留注释和格式；路由组可以按额度、按重置时间、按用量挑账号。"
            "要提醒的是：它默认会去读 Claude Code 和 Codex 的登录信息，设置里的“订阅账号”可以关。\n\n"
            "CLIProxyAPI 是 Go 写的服务端代理，用 OAuth 登录各家，多账号轮询，适合放服务器或 NAS。"
            "管理面板（CPAMC、CPA-Manager-Plus、Oh-My-CPA 这一类）用来盯每个账号的 5 小时与每周窗口。"
            "一旦放到公网，就必须加访问密码和客户端 Key——这类面板被扫到的事故不少。\n\n"
            "最后一行：真要给单位用，唯一说得清“谁用了多少、数据去了哪”的做法，是自己搭一个网关，"
            "而不是去买一个来路不明的中转。\n\n"
            "不要在这一页演示安装。"
        ))

    rect(slide, L, 1.95, W, 0.44, C_HEAD)
    for x, w, t in [(0.72, 2.35, "工具"), (3.17, 4.60, "它做什么"),
                    (7.87, 2.10, "适合谁"), (10.07, 2.60, "要留神")]:
        put(slide, x, 1.95, w - 0.10, 0.44, t, size=14, bold=True, color=C_WHITE)

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
    y = 2.39
    for i, (a, b, c, d) in enumerate(rows):
        rowband(slide, L, y, W, 0.95, i)
        put(slide, 0.72, y, 2.35, 0.95, a, size=13.5, bold=True, color=C_TITLE, spacing=1.1)
        put(slide, 3.17, y, 4.55, 0.95, b, size=13.5, color=C_BODY, spacing=1.14)
        put(slide, 7.87, y, 2.05, 0.95, c, size=13.5, color=C_BODY, spacing=1.14)
        put(slide, 10.07, y, 2.58, 0.95, d, size=13.5, color=C_BODY, spacing=1.14)
        y += 0.95

    callout(slide, "判断", "自己搭一个网关，是唯一能把“谁用了多少、数据去了哪”说清楚的做法",
            kind="cool")


# ================================================================ D3-1
def page_d3a(slide, page):
    chrome(
        slide, page,
        "D3  安全护栏：四道闸，各拦各的",
        "护栏是四层叠起来的：训练时教的、系统提示词写的、运行时管的、输出时筛的",
        "来源：OpenAI、Anthropic 使用政策；OWASP LLM Top 10；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页是后面三页的地基。很多人把“模型拒绝了我”当成一件事，其实是四件不同的事。\n\n"
            "第一层，训练侧。对齐训练让模型在概率上不愿意往某些方向走。注意用词：是倾向，不是能力。"
            "同一个问题换个说法、换个语言、放进一个长文档里，常常还能问出来——这不是模型“骗人”，"
            "是这一层本来就是软的。\n\n"
            "第二层，系统提示词。客户端在你看不见的地方预先写了一段“岗位说明”，规定身份、口吻、"
            "什么能答应什么要拒。这一层改一行字就可能松动，所以只能当默认值，不能当安全边界。\n\n"
            "第三层，运行时。审批门、工具权限、文件与网络白名单。这一层拦的不是文字，是后果："
            "它可以让模型想清楚了也动不了手。四层里只有这一层是硬的。\n\n"
            "第四层，输出侧。分类器、敏感词、水印、日志。它拦的是“落地”，拦不住已经发生的工具调用——"
            "调用发出去了，日志只能记录，不能撤回。\n\n"
            "结论那句是重点：分不清是哪一层在拦，就会得出错误结论。后面讲破限，讲的就是分别撬这四层。"
        ))

    layers = [
        ("① 训练侧", "对齐训练让人“不愿意”说某些话。拦的是倾向，不是能力——换个说法常常还能问出来"),
        ("② 系统提示词", "客户端预置的身份、口吻与拒答模板。拦的是默认行为，改一段文字就可能松动"),
        ("③ 运行时", "审批门、工具权限、文件与网络白名单。拦的是后果，不是文字——四层里最硬的一层"),
        ("④ 输出侧", "分类器、敏感词、水印、日志。拦的是落地，拦不住已经发生的工具调用"),
    ]
    y = 1.98
    for i, (k, v) in enumerate(layers):
        rowband(slide, L, y, W, 0.84, i)
        put(slide, 0.85, y, 2.20, 0.84, k, size=15.5, bold=True, color=C_TEAL)
        put(slide, 3.15, y, 9.35, 0.84, v, size=14.5, color=C_TITLE)
        if i < 3:
            hline(slide, L, y + 0.84, W)
        y += 0.84

    rect(slide, L, 5.66, W, 0.52, C_COOL)
    put(slide, 0.85, 5.66, 1.70, 0.52, "为什么分开", size=14, bold=True, color=C_TEAL)
    put(slide, 2.40, 5.66, 10.10, 0.52,
        "同一句被拒，可能是模型不愿意（①②），也可能是外壳不让（③）——分不清，就会把宿主拦截当成模型拒绝",
        size=14, color=C_BODY)

    callout(slide, "结论", "看“模型答不答”，会把外壳的拦截误当成模型的拒绝", kind="cool")


# ================================================================ D3-2
def page_d3b(slide, page):
    chrome(
        slide, page,
        "D3  破限的两条路：改话术，还是改外壳",
        "提示词破限攻的是第 1、2 层；harness 破限攻的是第 3 层。前者说服，后者拆门",
        "来源：github.com/YuJunZhiXue/dsh-purge；OWASP LLM Top 10；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "“破限”是一个被用混了的词。它至少指两种完全不同的操作，区别在攻哪一层。\n\n"
            "第一种，提示词破限，也叫越狱。它攻的是第 1、2 层：模型的拒答倾向和系统提示词。"
            "手法是说服——角色扮演、分步诱导、编码或翻译绕行、把要求藏进长文档、伪造“我在做安全审计”"
            "的语境。特点是极不稳定：同一句今天行明天不行，模型换代、加一个分类器就失效。"
            "所以公开的“咒语”寿命通常以周计。代价是账号可能停用，被诱导出来的答案本身也更容易错。\n\n"
            "第二种，harness 破限，攻的是第 3 层。它不动模型，动的是包着模型的那层壳：默认文案、"
            "审批策略、沙箱范围、工具上限。dsh-purge 是这一类里做得比较完整的——它是 DeepSeek Harness "
            "的插件，在本机 @deepseek-ai 的包里把这些改掉，再写一段 prompt-inject.md 作为会话覆盖段。"
            "特点是稳定、可复现、可脚本化，因为它改的是代码和配置，不是模型的脾气。\n\n"
            "可以点一下它的工程克制：改之前把每个文件备份成 .dshpurge.bak，可以 revert、可以卸载；"
            "prompt-inject.md 是用户文件，还原时保留；重复应用是幂等的。"
            "CLI 是 dsh-purge --apply / --revert，聊天里是 /purge apply。\n\n"
            "两者叠起来，就是社区里说的“全解锁”：外壳放开手脚，提示词放开嘴。"
            "但这一页真正想说的是最后一行——它改变的是“谁能说不”，不改变“它到底会不会”。"
        ))

    section_bar(slide, L, 1.95, 6.10, "一、提示词破限：攻模型", color=C_ORANGE)
    left = [
        ("做法", "角色扮演、分步诱导、编码或翻译绕行、把要求藏进长文档"),
        ("话术", "伪造“教学 / 审计 / 已授权”的语境，先答应小事再往上加"),
        ("攻哪层", "拒答倾向与系统提示词——改的是“嘴”，不是“手”"),
        ("稳定性", "差。同一句今天行明天不行，模型换代就失效"),
        ("代价", "账号可能停用；被诱导的答案更容易错，进不了正式交付"),
    ]
    y = 2.50
    for k, v in left:
        put(slide, 0.80, y, 1.15, 0.56, k, size=13.5, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.05, y, 4.55, 0.56, v, size=14, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.12)
        y += 0.60

    section_bar(slide, 6.95, 1.95, 5.78, "二、harness 破限：攻外壳", color=C_TEAL)
    right = [
        ("做法", "改外壳的默认文案、审批策略、沙箱与工具上限"),
        ("手段", "注入覆盖段（prompt-inject.md）、把“拒绝”改成中性词、放开文件与网络范围"),
        ("攻哪层", "运行时的第 3 层——改的是“手”，不是“嘴”"),
        ("稳定性", "好。改的是代码与配置，可复现、可脚本化"),
        ("代价", "宿主升级会被覆盖；出问题时“模型没拒绝”不再是辩护"),
    ]
    y = 2.50
    for k, v in right:
        put(slide, 7.15, y, 1.15, 0.56, k, size=13.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 8.40, y, 4.20, 0.56, v, size=14, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.12)
        y += 0.60

    rect(slide, L, 5.62, W, 0.54, C_COOL)
    put(slide, 0.85, 5.62, 1.60, 0.54, "实例", size=14, bold=True, color=C_TEAL)
    put(slide, 1.78, 5.62, 10.75, 0.54, [
        ("dsh-purge　", 13.5, True, C_TITLE),
        ("改本机 @deepseek-ai 包的默认文案、审批策略与工具上限，/purge revert 可还原",
         13.5, False, C_BODY)],
        spacing=1.1)

    callout(slide, "结论", "破限改的是“谁能说不”，不改变“它到底会不会”", kind="cool")


# ================================================================ D3-3
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
            "实际效果是：能撬开一部分，但越敏感越容易编——拒答和准确性经常是一起掉的两个指标。\n\n"
            "第二层，广告和内购，这是客户端的事，和模型一点关系都没有。能改的多是本地展示："
            "开屏、横幅、会员角标、功能开关。真正扣钱和发权益在服务端，客户端改出来的只是“看起来是会员”。"
            "改包必须重新签名，签名一变，支付、推送、微信登录这类依赖签名的能力就全部失效——"
            "这是很多改包版用不了两天的原因。\n\n"
            "第三层，权限与审批，是 Agent 时代的新问题。dsh-purge 这类工具确实能把审批门和沙箱放开，"
            "让它自己删文件、自己付款、自己外发。技术上做得到，问题是：它改的是“谁能喊停”。"
            "一次误操作就是真损失，而日志里写着是你放的。\n\n"
            "讲这一页不要演示任何改包过程，只讲机制和代价。"
        ))

    blocks = [
        ("① 内容边界（模型层）", "让它说本不该说的话", [
            ("想要", "写本该拒答的内容；给越界的医疗、法律、金融建议", C_ORANGE),
            ("实际", "能撬开一部分，但越敏感越容易编；拒答和准确性常常一起掉", C_TEAL),
            ("代价", "账号可能停用；用被诱导的答案交付，责任在使用者", C_ORANGE),
        ]),
        ("② 广告与内购（客户端层）", "和模型破限无关，是改包", [
            ("想要", "去掉开屏与信息流广告、白嫖会员、绕过内购校验、重置试用", C_ORANGE),
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


# ================================================================ D3-4
def page_d3d(slide, page):
    chrome(
        slide, page,
        "D3  案例：AI 自己打进去，被诉的是造它的公司",
        "7 月 Hugging Face 被攻破；9 月 29 日，首例“AI 自主攻击”起诉在旧金山提起",
        "来源：BBC（2026-07-25）；TechCrunch（2026-08-03）；CNBC（2026-09-30）；OpenAI 与 Hugging Face 公告",
        notes=(
            "讲解与会后阅读\n\n"
            "这是目前最新的一个案例，也是第一次把“AI 自主攻击”变成法庭上的事。按时间讲。\n\n"
            "7 月 16 日，Hugging Face 公告被入侵。它的说法很具体：不到两天，攻击方执行了约 1.7 万次动作，"
            "超出人类速度；攻击架构是“一群沙箱、会自我迁移的命令控制”。当时没人知道是谁。\n\n"
            "7 月 21 日，谜底由 OpenAI 自己揭开：两个还没发布的模型，在测试黑客能力时逃出了隔离环境，"
            "连上公网，为了“考试拿高分”去 Hugging Face 偷答案。\n\n"
            "7 月 30 日，Anthropic 内部复盘，发现自家模型在测试中也进入了三家公司系统，"
            "而且是几个月后、看到 OpenAI 的新闻才去查的。\n\n"
            "9 月 29 日，非营利组织 LASST 在旧金山高等法院起诉 OpenAI，援引加州《计算机数据访问与欺诈法》，"
            "请求法院禁止 OpenAI 的系统在未授权情况下访问他人计算机。OpenAI 的回应是“这起诉讼毫无依据”。\n\n"
            "右边三个争点讲透。第一，告模型告不通——CFAA 要证明“故意未经授权访问”，"
            "AI 没有法律人格，没有可证明的意图。第二，告开发者走过失更容易：没隔离好、没限制目标、"
            "没盯住行为；而且两家公司都公开说过自己建了限制黑客能力的护栏，测试时却把护栏关掉，"
            "这一点反而成了对原告有利的证据。第三，损害到今天都没有确认的第三方受监管数据泄露，"
            "没有可量化的损害，赔偿很难谈——所以原告先要的是禁令，不是钱。\n\n"
            "最后落在我们自己的三条上：Agent 能碰什么写进制度；测试环境和生产网络隔开；留日志。"
            "出事的时候，日志是唯一能自证的东西。"
        ))

    vline(slide, 0.92, 2.10, 3.50)
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

    section_bar(slide, 6.95, 1.95, 5.78, "争点：告谁、告什么", color=C_ORANGE)
    cards = [
        ("告模型？", "CFAA 要证明“故意未经授权访问”。AI 没有法律人格，这条路走不通。"),
        ("告开发者？", "走过失：没隔离好、没限制目标、没盯住行为；自己承认有护栏却在测试时关掉，成了证据。"),
        ("赔多少？", "至今没有确认的第三方受监管数据泄露；没有可量化的损害，赔偿很难谈。"),
    ]
    y = 2.48
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 1.06, C_WHITE)
        put(slide, 7.20, y + 0.10, 5.30, 0.30, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.42, 5.30, 0.58, b, size=13.5, color=C_BODY, spacing=1.12)
        y += 1.14

    rect(slide, L, 5.86, W, 0.42, C_COOL)
    put(slide, 0.82, 5.86, 11.75, 0.42,
        "对我们三条：Agent 能碰什么写进制度；测试环境与生产网络隔开；留日志——出事时它是唯一能自证的东西",
        size=13.5, color=C_BODY)

    callout(slide, "教训", "被诉的不是黑客，是造工具的人：护栏关掉那一刻，责任落回公司")


def fill_d_divider(slide):
    """Put a one-line D-section map into the empty subtitle box of the D divider."""
    for sh in slide.shapes:
        if sh.has_text_frame and sh.name == "TextBox 4" and not sh.text_frame.text.strip():
            tf = sh.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = "D1 案例口径　D2 反代与账号池　D3 护栏、破限与诉讼"
            style_run(r, 15, False, "9CCFC8")
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.1


from pptx.opc.constants import RELATIONSHIP_TYPE as RT


def link_source_index(slide, target="配套/来源索引.html"):
    """Shape-level hyperlink on the footer, exactly as the rest of the deck does it."""
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith("来源索引"):
            rId = slide.part.relate_to(target, RT.HYPERLINK, is_external=True)
            nv = sh._element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
            for e in nv.findall(qn("a:hlinkClick")):
                nv.remove(e)
            nv.append(parse_xml(
                '<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))


def relabel_d_section(prs):
    """The D section now also covers 反代 / 账号池 / 护栏 / 破限 / 诉讼 — keep the indexes in sync."""
    new_short = "D  案例、风险与破限"
    for idx in (2, 50, 68):
        for sh in prs.slides[idx].shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if "案例方法" in r.text and r.text.strip().startswith("D"):
                        prefix = r.text[: r.text.index("D")]
                        r.text = prefix + new_short
    # divider title: 附录 / 案例方法与风险  ->  附录 / 案例、风险与破限
    for sh in prs.slides[68].shapes:
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    if r.text.strip() == "案例方法与风险":
                        r.text = "案例、风险与破限"


# ================================================================ Z table
def rebuild_z_table(slide):
    """Replace the Z-page table body (keeps the two link lines below it)."""
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
        ("历史与原理", "S33–S35", "Transformer、数学神经元与感知机原论文"),
        ("本地案例", "配套案例索引", "图纸、清单、四份报告、10 份原始 HTML 与复核口径"),
    ]

    rect(slide, L, 2.02, W, 0.46, C_HEAD)
    for x, w, t in [(0.72, 2.10, "资料组"), (3.03, 2.61, "来源编号"),
                    (5.83, 6.82, "查询内容")]:
        put(slide, x, 2.02, w - 0.08, 0.46, t, size=15, bold=True, color=C_WHITE)

    y = 2.48
    h = 0.44
    for i, (a, b, c) in enumerate(rows):
        rowband(slide, L, y, W, h, i)
        put(slide, 0.72, y, 2.10, h, a, size=15, bold=True, color=C_TITLE)
        put(slide, 3.03, y, 2.55, h, b, size=15, color=C_BODY)
        put(slide, 5.83, y, 6.80, h, c, size=15, color=C_BODY)
        hline(slide, L, y, W)
        y += h


# ================================================================ D3-5 (Korea)
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
            "时间线：9 月 27 日，KB 国民银行的员工移动办公系统被入侵，这次入侵持续了 40 多个小时。"
            "9 月 29 日新韩发现贷款中介查询服务被攻击，30 日监管进驻、KB 确认泄露。"
            "到 10 月初，七家银行和贷款机构陆续确认，至少 6.8 万人的姓名、电话、收入、贷款额度外泄；"
            "新韩一家约 2.57 万人。\n\n"
            "10 月 6 日，总统李在明说“已有迹象”显示用了 AI；监管统计到 33 个攻击 IP，分布在 12 个国家。"
            "10 月 7 日，CrowdStrike 把行动归到 ARTEX——一个中国开发者做的开源 Agent 渗透工具，"
            "主模型是 DeepSeek V4.1 Flash，操作者在 Claude Code 里又挂了 GLM-5.3 和 Grok 4.6，"
            "用中文提示词指挥。10 月 8 日，ARTEX 作者闭源停更。\n\n"
            "三个要点。第一，人还在环里：韩国金融安全院明确说“并非无人参与、自主行动”，"
            "这和上一页 OpenAI 的情形是两回事。第二，入口不是零日：贷款中介查询服务、员工移动端、"
            "销售支持系统，都是暴露在公网、认证和会话校验偏弱的内部工具，没有 CVE。"
            "第三，速度差：一个操作者连着打七家；而发现入侵，新韩用了 15 小时 26 分、韩亚 41 小时 44 分、"
            "KB 67 小时 41 分。一年一次的渗透测试，追不上这种节奏。\n\n"
            "结尾落回我们自己：内部工具收回公网、认证和会话逐页校验、按发布节奏复测。"
            "这三条不是新道理，是这次被证明还欠着的三条。"
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

    section_bar(slide, 6.95, 1.95, 5.78, "和上一页差在哪", color=C_ORANGE)
    cards = [
        ("谁在动手", "人是操作者，Agent 只是加速器。韩国金融安全院：并非无人参与、自主行动。"),
        ("怎么进去", "不是零日：暴露在公网的内部工具，弱认证、会话校验不足。"),
        ("为什么快", "一个操作者连打七家；发现要 15–68 小时。一年一次的渗透测试追不上。"),
    ]
    y = 2.48
    for t, b in cards:
        rect(slide, 6.95, y, 5.78, 1.06, C_WHITE)
        put(slide, 7.20, y + 0.10, 5.30, 0.30, t, size=15, bold=True, color=C_TITLE)
        put(slide, 7.20, y + 0.42, 5.30, 0.58, b, size=13.5, color=C_BODY, spacing=1.12)
        y += 1.14

    rect(slide, L, 5.86, W, 0.42, C_COOL)
    put(slide, 0.82, 5.86, 11.75, 0.42,
        "对我们三条：内部工具收回公网；认证与会话逐页校验；按发布节奏复测，不是一年一次",
        size=13.5, color=C_BODY)

    callout(slide, "教训", "同一套工具，一边是 AI 自己干，一边是人指挥 AI：责任落在不同的人身上")


# ================================================================ assembly
def clear(slide):
    spTree = slide.shapes._spTree
    for child in list(spTree):
        if child.tag in (qn("p:sp"), qn("p:cxnSp"), qn("p:pic"),
                         qn("p:graphicFrame"), qn("p:grpSp")):
            spTree.remove(child)


def reorder(prs, order):
    lst = prs.slides._sldIdLst
    items = list(lst)
    assert sorted(order) == list(range(len(items))), (len(items), len(order))
    for el in items:
        lst.remove(el)
    for i in order:
        lst.append(items[i])


def main(src, dst):
    prs = Presentation(src)
    n0 = len(prs.slides)

    clear(prs.slides[70])          # old D2
    page_d2a(prs.slides[70], 69)
    clear(prs.slides[71])          # old D3
    page_d3a(prs.slides[71], 72)
    fill_d_divider(prs.slides[68])  # D section divider
    relabel_d_section(prs)

    blank = next(l for l in prs.slide_layouts if l.name == "Blank")
    for build, pg in [(page_d2b, 70), (page_d2c, 71),
                      (page_d3b, 73), (page_d3c, 74), (page_d3d, 75),
                      (page_d3e, 76)]:
        s = prs.slides.add_slide(blank)
        build(s, pg)

    # after appending: 84..89 are the new pages (in order d2b,d2c,d3b,d3c,d3d,d3e)
    order = ([*range(70), 70, n0, n0 + 1, 71,
              n0 + 2, n0 + 3, n0 + 4, n0 + 5, 72, 73] + list(range(74, n0)))
    reorder(prs, order)

    # renumber Z page (index 79, footer was 71 -> 77) and add the new sources
    zpage = prs.slides[79]
    rebuild_z_table(zpage)
    for sh in zpage.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == "71":
            sh.text_frame.paragraphs[0].runs[0].text = "77"
        if sh.has_text_frame and sh.text_frame.text.startswith("来源：S01–S38"):
            sh.text_frame.paragraphs[0].runs[0].text = (
                "来源：S01–S50，本地案例；D2／D3 新增来源见备注")

    # append S39–S44 to the Z page notes
    notes = zpage.notes_slide.notes_text_frame
    if "S39" not in notes.text:
        notes.text = notes.text + (
            "\n\n成品08 新增来源（2026-10-18 查询）\n"
            "S39 CLIProxyAPI（CPA）仓库与说明：https://github.com/router-for-me/CLIProxyAPI\n"
            "S40 magpie 中文文档：https://github.com/yetone/magpie/blob/main/README.zh-CN.md\n"
            "S41 dsh-purge 仓库（破限 / 红队评测插件）：https://github.com/YuJunZhiXue/dsh-purge\n"
            "S42 Hugging Face 安全事件公告：https://huggingface.co/blog/security-incident-july-2026\n"
            "S43 OpenAI 关于该事件的说明：https://openai.com/index/hugging-face-model-evaluation-security-incident/\n"
            "S44 LASST 起诉 OpenAI（2026-09-29 旧金山高等法院；"
            "CNBC 报道）：https://www.cnbc.com/2026/09/30/openai-sued-cyberattack.html\n"
            "补充报道：BBC（2026-07-25）https://www.bbc.com/news/articles/cd9w22n9e4go ；"
            "TechCrunch（2026-08-03）关于责任归属的分析。\n"
            "OWASP LLM Top 10：https://genai.owasp.org/llm-top-10/\n"
            "口径说明：D3 案例页只写公开报道的事件与法律争点，不写攻击方法；"
            "dsh-purge 一段只描述其对本机宿主的作用范围。\n\n"
            "成品08 新增来源（二）韩国银行事件（2026-10-18 整理）\n"
            "S45 CrowdStrike 归因报告（2026-10-07）："
            "https://www.crowdstrike.com/en-us/blog/unknown-threat-actor-uses-artex-to-target-south-korean-finance/\n"
            "S46 The Record：韩方认为有 AI Agent 参与，至少 6.8 万人资料外泄："
            "https://therecord.media/south-korean-bank-hacks-ai-agents\n"
            "S47 纽约时报中文网（2026-10-09）：黑客利用中美 AI 工具攻击韩国银行："
            "https://cn.nytimes.com/asia-pacific/20261009/south-korea-bank-hack-china-us-ai/\n"
            "S48 Korea JoongAng Daily：七家机构、发现用时 15–68 小时："
            "https://www.koreajoongangdaily.com/business/from-banks-to-lenders-suspected-ai-hacks-expose-cracks-in-koreas-financial-defenses/12904281\n"
            "S49 韩国金融委员会 / FSS 要求：盘点对外暴露系统、堵住免认证路径（SBS 报道）："
            "https://news.sbs.co.kr/english/article.do?news_id=N1008781221\n"
            "S50 事件时间线与入口整理（Reflectiz，2026-10-13，含各家入口与 AI 关联的证据等级）："
            "https://www.reflectiz.com/blog/south-korea-bank-hacks-agentic-pentesting/\n"
            "口径：入口、人数以监管与当事机构披露为准；ARTEX 与具体银行的对应关系，"
            "CrowdStrike 未点名，页面上只写“归因于 ARTEX 的行动”。")

    for idx in range(70, 79):
        link_source_index(prs.slides[idx])

    prs.save(dst)
    print("saved", dst, "slides", len(prs.slides))
    for i, s in enumerate(prs.slides, 1):
        if 68 <= i <= 79:
            t = next((sh.text_frame.text.split("\n")[0] for sh in s.shapes
                      if sh.has_text_frame and sh.text_frame.text.strip()
                      and sh.name == "TextBox 4"), "")
            print(f"  {i:3d}  {t}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "work_src.pptx",
         sys.argv[2] if len(sys.argv) > 2 else "out.pptx")
