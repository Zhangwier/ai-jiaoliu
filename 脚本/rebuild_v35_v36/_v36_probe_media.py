# -*- coding: utf-8 -*-
import zipfile, re
z = zipfile.ZipFile(r"D:\MG\program\ai交流\成品\AI交流分享-成品34.pptx")
for sn in [1, 2, 4, 9, 46, 113]:
    t = "ppt/slides/_rels/slide%d.xml.rels" % sn
    if t not in z.namelist():
        print(sn, "missing")
        continue
    rels = z.read(t).decode("utf-8")
    hits = [h for h in re.findall(r'Target="\.\./media/([^"]+)"', rels) if h.endswith((".mp4", ".gif"))]
    print(sn, hits)
