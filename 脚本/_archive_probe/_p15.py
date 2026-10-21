import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (18,24,58,59,60,66,68,73,74):
    d[i].get_pixmap(dpi=120).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print("ok")
