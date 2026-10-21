import os
from PIL import Image, ImageDraw
m=r"_v34\media"
files=["s001_1.png","s012_1.png","s016_1.png","s029_1.png","s029_2.png","s029_3.png",
"s033_1.png","s034_1.png","s034_2.png","s035_1.png","s035_2.png","s037_1.png","s038_1.png","s038_2.png",
"s060_2.jpg","s062_1.png","s095_1.png","s096_1.jpg","s100_1.png","s101_1.png","s111_1.png","s112_1.png",
"s114_1.png","s118_1.png","s120_1.png","s120_2.png","s064_1.gif","s065_1.jpg","s019_1.gif","s024_8.jpg","s026_6.jpg","s027_8.jpg"]
CW,CH=360,260
cols=6; rows=(len(files)+cols-1)//cols
sheet=Image.new("RGB",(cols*CW,rows*(CH+18)),(255,255,255))
d=ImageDraw.Draw(sheet)
for i,f in enumerate(files):
    p=os.path.join(m,f)
    if not os.path.exists(p): continue
    im=Image.open(p); 
    if im.mode=="P": im=im.convert("RGBA")
    im=im.convert("RGB")
    w,h=im.size
    s=min((CW-8)/w,(CH-8)/h)
    im=im.resize((max(1,int(w*s)),max(1,int(h*s))),Image.LANCZOS)
    x=(i%cols)*CW; y=(i//cols)*(CH+18)
    sheet.paste(im,(x+4,y+4))
    d.text((x+6,y+CH+2),f"{f} {w}x{h}",fill=(0,0,0))
sheet.save(r"_v35\sheet.png")
print(sheet.size)
