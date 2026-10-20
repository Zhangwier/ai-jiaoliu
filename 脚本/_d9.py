from PIL import Image
import glob, os
for d in (r"D:\MG\program\ai交流\成品\_work29\logos3", r"D:\MG\program\ai交流\成品\_work29\logos2"):
    print("##", d)
    for f in sorted(glob.glob(d+r"\*.png")):
        im = Image.open(f).convert("RGBA")
        px = [p for p in im.getdata() if p[3]>128]
        if not px: print("  ", os.path.basename(f), "EMPTY"); continue
        r=sum(p[0] for p in px)/len(px); g=sum(p[1] for p in px)/len(px); b=sum(p[2] for p in px)/len(px)
        print(f"   {os.path.basename(f):<16} rgb=({r:.0f},{g:.0f},{b:.0f}) lum={0.2126*r+0.7152*g+0.0722*b:.0f} n={len(px)}")
