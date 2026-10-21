# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

rep('''        inside = BX + 0.14
        room = bw - 0.24
        if est_w(models, 10.5, True) / 72.0 <= room:
            text(s, inside, yy + 0.075, room, 0.24, models, 10.5, WHITE, SANS, True,
                 PP_ALIGN.LEFT, 1.0, check=False)
        else:
            text(s, BX + bw + 0.22, yy + 0.075, BW - bw - 0.30, 0.24, models, 10.5, GREY, SANS,
                 False, PP_ALIGN.LEFT, 1.0, check=False)''',
    '''        room = bw - 0.28
        if i == 3:
            text(s, BX, yy + 0.34, BW, 0.24, models, 10, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
        else:
            text(s, BX + 0.14, yy + 0.075, room, 0.24, models, 10.5, WHITE, SANS, True,
                 PP_ALIGN.LEFT, 1.0, check=False)''')
rep('''    text(s, ML + 6.84, 6.16, 5.28, 0.60,
         "分层比的是“最好那一档”。日常任务里，第二梯队与第一梯队的差距往往小于\\n“外壳是否合适”带来的差距。",
         11, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)''',
    '''    text(s, ML + 6.84, 6.16, 5.28, 0.26,
         "分层比的是“最好那一档”，不是平均水位。", 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
         check=False)
    text(s, ML + 6.84, 6.44, 5.28, 0.26,
         "日常任务里，外壳选错造成的差距常常大于档位差距。", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
