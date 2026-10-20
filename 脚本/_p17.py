import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (70,72,73,74,58,59,60,61,62,63,64,65,66,67,68,69):
    pass
for i in (70,72,74):
    d[i].get_pixmap(dpi=115).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print("ok")
