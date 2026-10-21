import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
d[82].get_pixmap(dpi=130).save(r"D:\MG\program\ai交流\成品\_work29\pdf\s83.png")
