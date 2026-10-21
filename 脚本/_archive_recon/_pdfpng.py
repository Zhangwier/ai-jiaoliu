import pymupdf as fitz
d = fitz.open(r"D:\MG\program\ai交流\成品\_work29\pdf\base.pdf")
print("pages", d.page_count, "rect", d[0].rect)
for i in (4,19,24,25,50,51,88,89,90,91):
    pix = d[i].get_pixmap(dpi=130)
    pix.save(rf"D:\MG\program\ai交流\成品\_work29\pdf\s{i+1}.png")
    print(i+1, pix.width, pix.height)
