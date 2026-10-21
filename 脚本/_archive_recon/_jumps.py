import zipfile, re

z = zipfile.ZipFile("AI交流分享-成品26.pptx")
for n in z.namelist():
    if n.startswith("ppt/slides/_rels/"):
        d = z.read(n).decode("utf-8")
        tg = re.findall(r'Target="(slide\d+\.xml)"', d)
        if tg:
            print(n, "->", tg)
