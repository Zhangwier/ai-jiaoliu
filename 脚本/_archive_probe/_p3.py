import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
print("pages", d.page_count)
for i in (20,21,22,23,57,58,59,61,67,69,88,89):
    d[i].get_pixmap(dpi=130).save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
print("ok")
