import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (18,20,22,23,50,51,58,66,88):
    d[i].get_pixmap(dpi=130).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print("ok")
