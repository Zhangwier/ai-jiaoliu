import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s=prs.slides[57]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip()=="官方承认过 bug（见 C6）":
        settext(sh, "官方承认过 bug（见 B1c）", size=10.5)
        print("s58 Ok")
s=prs.slides[66]
for sh in s.shapes:
    if sh.has_text_frame and "（C4、C7）" in sh.text_frame.text:
        settext(sh, sh.text_frame.text.replace("（C4、C7）","（C4、C5）"), size=11.5)
        print("s67 Ok", sh.text_frame.text[:80])
save(prs); print("saved")
