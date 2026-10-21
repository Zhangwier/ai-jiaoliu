# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

old = '''    M = [("同时看整句", "不是从左往右一个字一个字读，一次把整句放进来，谁离谁近都能算。"),
         ("远距离也能关联", "“清单”和“列明”隔了四个词，分数依然高 —— 这是它比老办法强的地方。"),
         ("叠很多层", "每一层重新算一遍注意力，前面的结果变成后面的输入。"),
         ("窗口长了会不会丢", "只能实测：把要求放在开头、结尾、中段，各跑一次看差异。")]
    yy = 2.06
    for i, (t, b) in enumerate(M):
        text(s, RX, yy, 0.42, 0.24, f"0{i+1}", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.48, yy - 0.03, 5.40, 0.30, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        nl = nlines(b, 12.5, 5.36)
        text(s, RX + 0.48, yy + 0.30, 5.36, nl * 12.5 * 1.32 * 1.45 / 72 + 0.06, b, 12.5, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
        yy += 0.36 + nl * 12.5 * 1.32 * 1.45 / 72 + 0.24
    box(s, RX, 5.68, 5.90, 0.78, fill=MINT)
    text(s, RX + 0.20, 5.76, 5.50, 0.68,
         "注意力不是“理解”：它只决定看哪里、看多重，不判断那里说的是对是错。", 13, TEAL, SANS,
         True, PP_ALIGN.LEFT, 1.4)'''

new = '''    M = [("同时看整句", "不是从左往右一个字一个字读，一次把整句放进来，谁离谁近都能算。"),
         ("远距离也能关联", "“清单”和“列明”隔了四个词，分数依然高 —— 这是它比老办法强的地方。"),
         ("叠很多层", "每一层重新算一遍注意力，上一层的输出变成下一层的输入。")]
    yy = 2.04
    for i, (t, b) in enumerate(M):
        text(s, RX, yy, 0.42, 0.24, f"0{i+1}", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.48, yy - 0.03, 5.40, 0.30, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        nl = nlines(b, 12.5, 5.36)
        text(s, RX + 0.48, yy + 0.30, 5.36, nl * 12.5 * 1.32 * 1.45 / 72 + 0.06, b, 12.5, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
        yy += 0.36 + nl * 12.5 * 1.32 * 1.45 / 72 + 0.30
    box(s, RX, 4.94, 5.90, 1.52, fill=MINT)
    text(s, RX + 0.24, 5.06, 5.42, 0.26, "注意力的边界", 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    text(s, RX + 0.24, 5.34, 5.44, 1.02,
         "它只决定看哪里、看多重，不判断那里说的是对是错。\\n窗口长了关键信息还能不能稳定用上，只能实测 —— 把要求放开头、结尾、中段各跑一次，比看参数表有用。",
         12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.42)'''

assert old in s
s = s.replace(old, new)
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
