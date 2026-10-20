import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (58,60,62,64,72,74):
    d[i].get_pixmap(dpi=100).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\b{i+1}.png")
print("ok")
