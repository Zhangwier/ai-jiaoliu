import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
d[88].get_pixmap(dpi=130).save(r"D:\MG\program\ai交流\成品\_work29\pdf\s89.png")
print("ok")
