import os
from PIL import Image, ImageDraw
m=r"_v34\media"
files=["s111_1.png","s112_1.png","s114_1.png","s120_1.png","s120_2.png","s095_1.png","s096_1.jpg","s100_1.png","s101_1.png","s062_1.png","s012_1.png","s016_1.png"]
CW,CH=520,380
cols=4; rows=(len(files)+cols-1)//cols
sheet=Image.new("RGB",(cols*CW,rows*(CH+22)),(250,250,250))
d=ImageDraw.Draw(sheet)
for i,f in enumerate(files):
    p=os.path.join(m,f)
    im=Image.open(p).convert("RGB")
    w,h=im.size
    s=min((CW-10)/w,(CH-10)/h)
    im=im.resize((max(1,int(w*s)),max(1,int(h*s))),Image.LANCZOS)
    x=(i%cols)*CW; y=(i//cols)*(CH+22)
    sheet.paste(im,(x+5,y+5))
    d.text((x+6,y+CH+4),f"{f} {w}x{h}",fill=(0,0,0))
sheet.save(r"_v35\sheet2.png"); print(sheet.size)
