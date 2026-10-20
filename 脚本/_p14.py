import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (18,82):
    d[i].get_pixmap(dpi=130).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print(len(d))
