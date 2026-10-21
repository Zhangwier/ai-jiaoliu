# -*- coding: utf-8 -*-
import io
f="app.py"; s=io.open(f,encoding="utf-8").read()
old = '''        rich(s, x, yy + 0.46, cw, 0.20,
             [{"runs": [(good + "\u3000", C(0x2B, 0x34, 0x45), 9.0, False),
                        ("\u8fb9\u754c\uff1a" + bad, ORANGE, 9.0, False)], "ls": 1.05}])'''
new = '''        text(s, x, yy + 0.46, cw * 0.46, 0.20, good, 9.0, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.05, check=False)
        text(s, x + cw * 0.47, yy + 0.46, cw * 0.53, 0.20, "\u8fb9\u754c\uff1a" + bad, 9.0, ORANGE, SANS,
             False, PP_ALIGN.LEFT, 1.05, check=False)'''
if old not in s: raise SystemExit("ANCHOR FAIL")
io.open(f,"w",encoding="utf-8").write(s.replace(old,new,1))
print("ok")
