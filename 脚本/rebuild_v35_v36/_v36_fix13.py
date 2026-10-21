# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

# p24：强制行高，结论条固定位置
rep('''    ty = y + 0.20
    end = table(s, ML, ty, CW, rows, [2.30, 3.20, 3.30, 3.20], size=13.5, head_size=13.5,
                minrow=0.44, head_h=0.44, first_bold=True, pad=0.14,
                colors=lambda r, c: (ORANGE if c == 2 else (GREY if c == 1 else C(0x2B, 0x34, 0x45))))
    takeaway(s, "所以榜单的正确用法是：先看你这类活被哪个榜考到了，再看那个榜的第一名是谁。", y=end + 0.26,
             tag="判断")''',
    '''    ty = y + 0.16
    table(s, ML, ty, CW, rows, [2.30, 3.20, 3.30, 3.20], size=13, head_size=13,
          minrow=0.50, head_h=0.42, rowh=0.50, first_bold=True, pad=0.14,
          colors=lambda r, c: (ORANGE if c == 2 else (GREY if c == 1 else C(0x2B, 0x34, 0x45))))
    takeaway(s, "所以榜单的正确用法是：先看你这类活被哪个榜考到了，再看那个榜的第一名是谁。", y=6.10,
             tag="判断")''')
rep('''["Epoch 综合", "跨多个基准加权", "Claude Opus 5.5", "GPT-6 Astra（几乎并列）"],''',
    '''["Epoch 综合", "跨多个基准加权", "Claude Opus 5.5", "GPT-6 Astra 几乎并列"],''')

# p26：表行高固定
rep('''    ty = y + 0.20
    end = table(s, ML, ty, CW, rows, [1.90, 6.10, 4.10], size=13, head_size=13, minrow=0.72,
                head_h=0.44, first_bold=True, pad=0.16,
                colors=lambda r, c: TEAL if c == 0 else (GREY if c == 2 else C(0x2B, 0x34, 0x45)))
    takeaway(s, "“脾气”不是性格，是训练取舍的结果 —— 也就意味着它稳定、可预期、可利用。", y=end + 0.22,
             tag="判断")''',
    '''    ty = y + 0.16
    table(s, ML, ty, CW, rows, [1.90, 6.10, 4.10], size=12.5, head_size=12.5, minrow=0.60,
          head_h=0.42, rowh=0.62, first_bold=True, pad=0.16,
          colors=lambda r, c: TEAL if c == 0 else (GREY if c == 2 else C(0x2B, 0x34, 0x45)))
    takeaway(s, "“脾气”不是性格，是训练取舍的结果 —— 也就意味着它稳定、可预期、可利用。", y=6.10,
             tag="判断")''')

# p27：表行高固定 + 说明下移
rep('''    ty = y + 0.18
    end = table(s, ML, ty, CW, rows, [2.10, 2.70, 3.40, 3.90], size=12.5, head_size=12.5,
                minrow=0.44, head_h=0.42, first_bold=True, pad=0.14,
                colors=lambda r, c: (TEAL if (c == 2 and r > 0) else
                                     (GREY if c in (1, 3) else C(0x2B, 0x34, 0x45))))
    text(s, ML, end + 0.22, CW, 0.52,''',
    '''    ty = y + 0.14
    table(s, ML, ty, CW, rows, [2.10, 2.70, 3.40, 3.90], size=12, head_size=12,
          minrow=0.40, head_h=0.38, rowh=0.40, first_bold=True, pad=0.14,
          colors=lambda r, c: (TEAL if (c == 2 and r > 0) else
                               (GREY if c in (1, 3) else C(0x2B, 0x34, 0x45))))
    text(s, ML, 6.06, CW, 0.52,''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
