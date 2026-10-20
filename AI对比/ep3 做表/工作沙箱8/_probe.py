import fitz, json
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = fitz.open(p)
for i in [1,2,3]:
    pg = d[i]
    t = pg.get_text("text")
    print("="*30, "PAGE", i+1, "="*30)
    print(t[:2000])
    print("CHARS:", len(t.strip()), "IMAGES:", len(pg.get_images(full=True)), "DRAWINGS:", len(pg.get_drawings()))
