import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[85]
m = {
 "TextBox 13": "Cline：首月 $1（之后 $9.99）\nCodex：Go 档 $8/月\nopencode：Go 档含免费额度\nGLM Coding：Lite ¥118/月\n首月半价：先看续费价",
 "TextBox 27": "公益站、中转：AnyRouter 一类\n闲鱼共享号 / 代充「jio」：几元/月\n资料要过别人的服务器\n维护靠人，随时会停",
}
for sh in s.shapes:
    if sh.name in m:
        settext(sh, m[sh.name], size=12)
save(prs); print("ok")
