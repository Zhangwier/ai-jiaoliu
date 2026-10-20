import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
print("rotation", pg.rotation)
print("rect", pg.rect)
# make rotated page copy so display coords == page coords
mat = pymupdf.Matrix(pg.rotation)
t = pymupdf.TextWriter if False else None
