import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
for i in (88,89,90,91,18,57):
    d[i].get_pixmap(dpi=120).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print("ok")
