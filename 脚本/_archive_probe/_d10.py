from PIL import Image
import glob, os
d=r"D:\MG\program\ai交流\成品\_work29\logos4"
for f in sorted(glob.glob(d+r"\*.png")):
    im = Image.open(f).convert("RGBA")
    px = [p for p in im.get_flattened_data() if p[3]>128] if hasattr(im,'get_flattened_data') else [p for p in im.getdata() if p[3]>128]
    if not px: print(f"   {os.path.basename(f):<16} EMPTY"); continue
    r=sum(p[0] for p in px)/len(px); g=sum(p[1] for p in px)/len(px); b=sum(p[2] for p in px)/len(px)
    print(f"   {os.path.basename(f):<16} lum={0.2126*r+0.7152*g+0.0722*b:6.0f}  size={im.size}")
