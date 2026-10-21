# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()
old = '''    text(s, ML, 5.68, 5.62, 1.12,
         "最高的是“招标”（31%），抽中的是“列明”（38% 之外的另一个位置）—— 这张表本身就是示意的，\\n但机制是真的：每次抽一点随机性，长文本就会分叉。",'''
new = '''    text(s, ML, 5.68, 5.62, 1.12,
         "概率最高的是“列明”（38%），实际抽中的是“招标”（31%）—— 表里的数字是示意，\\n但机制是真的：每一步都留一点随机性，长文本就会越走越岔。",'''
assert old in s
s = s.replace(old, new)
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
