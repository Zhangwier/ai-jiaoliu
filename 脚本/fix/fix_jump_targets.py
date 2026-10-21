# -*- coding: utf-8 -*-
"""修正两处跳转目标：原指向旧附录 A 分节页（slide58 号），重排后应指向新附录首节 A（现 slide58）。"""
import zipfile, shutil, os, re

F = r"D:\MG\program\ai交流\AI交流分享-成品26.pptx"
TMP = F + ".tmp"

# 现在 slide81 = 附录 C 分节页；新附录首节 A = slide58
TARGET_OLD = "slide81.xml"
TARGET_NEW = "slide58.xml"

zin = zipfile.ZipFile(F, "r")
zout = zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED)
n = 0
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename in ("ppt/slides/_rels/slide3.xml.rels", "ppt/slides/_rels/slide57.xml.rels"):
        t = d.decode("utf-8")
        # 只替换 slide 类型的目标，不动 image 等
        t2 = t.replace('Target="%s"' % TARGET_OLD, 'Target="%s"' % TARGET_NEW)
        if t2 != t:
            n += 1
        d = t2.encode("utf-8")
    zout.writestr(it, d)
zout.close()
zin.close()
os.replace(TMP, F)
print("rels updated:", n)
