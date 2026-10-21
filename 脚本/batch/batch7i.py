import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[85]
for sh in s.shapes:
    if sh.name == "TextBox 27":
        settext(sh, "开源中转 / 公益站：AnyRouter、opencode go 一类（L 站常见）\n闲鱼共享号、代充「jio」一类：几元/月\n提示词与资料要过别人的服务器\n维护靠个人，随时可能停", size=12)
        print("ok")
save(prs)
