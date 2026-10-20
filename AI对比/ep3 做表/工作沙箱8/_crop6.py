import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
S=400/72
ys=[560,629,697,767,835,904,973,1042,1111,1180,1249,1317,1387,1455,1524,1593,1731,1800,1869,1937,2007,2075,2144,2207,2280,2351,2420,2489,2557,2627,2695,2764,2833,2902,2971,3040,3109,3177,3247,3314,3385,3453,3522,3591,3660,3729,3798,3867,3935,4005,4073,4142,4210]
ys=[y/S for y in ys]
bands=[(560/S, ys[8]),(ys[8],ys[16]),(ys[16],ys[24]),(ys[24],ys[32]),(ys[32],ys[40]),(ys[40],ys[48]),(ys[48],4215/S)]
for i,(a,b) in enumerate(bands,1):
    r=pymupdf.Rect(1160/S,a,2870/S,b)
    pix=pg.get_pixmap(dpi=300, clip=r); out=rf"work\LB{i}.png"; pix.save(out); print(out,pix.width,pix.height)
