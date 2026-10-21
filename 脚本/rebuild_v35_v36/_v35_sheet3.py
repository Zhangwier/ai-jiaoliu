from PIL import Image
import os
files=[('s114_1.png',820),('s118_1.png',820),('s112_1.png',820),('s111_1.png',700),('s012_1.png',820),('s062_1.png',820)]
tiles=[]
for f,w in files:
    im=Image.open('_v34/media/'+f).convert('RGBA')
    bg=Image.new('RGBA',im.size,(255,255,255,255))
    bg.alpha_composite(im)
    im=bg.convert('RGB')
    s=w/im.size[0]
    im=im.resize((w,int(im.size[1]*s)),Image.LANCZOS)
    tiles.append((f,im))
# 3列网格
cols=3; pad=16
cw=max(t[1].size[0] for t in tiles)
rows=[]
for i in range(0,len(tiles),cols):
    rows.append(tiles[i:i+cols])
rh=[max(t[1].size[1] for t in r) for r in rows]
W=cols*cw+(cols+1)*pad; H=sum(rh)+ (len(rows)+1)*pad
sheet=Image.new('RGB',(W,H),(245,245,245))
y=pad
for ri,r in enumerate(rows):
    x=pad
    for f,im in r:
        sheet.paste(im,(x,y)); x+=cw+pad
    y+=rh[ri]+pad
sheet.save('_v35/sheet3.png')
print(sheet.size)
