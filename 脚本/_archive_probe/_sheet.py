import os, glob
from PIL import Image
d = r"D:\MG\program\ai交流\成品\_work29\logos2"
files = sorted(glob.glob(os.path.join(d,"*.png")))
cell=160; cols=7
rows=(len(files)+cols-1)//cols
sheet=Image.new("RGB",(cols*cell, rows*cell),(255,255,255))
for i,f in enumerate(files):
    im=Image.open(f).convert("RGBA")
    print(os.path.basename(f), im.size, "alpha0%=", round(100*sum(1 for p in im.getdata() if p[3]<10)/ (im.size[0]*im.size[1]),1))
    im.thumbnail((cell-16,cell-16))
    x=(i%cols)*cell; y=(i//cols)*cell
    sheet.paste(im,(x+8,y+8),im)
sheet.save(r"D:\MG\program\ai交流\成品\_work29\sheet2.png")
print("saved")
