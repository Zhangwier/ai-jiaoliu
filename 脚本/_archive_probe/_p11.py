import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
d[23].get_pixmap(dpi=130).save(r"D:\MG\program\ai交流\成品\_work29\pdf\s24.png")
print("ok")
