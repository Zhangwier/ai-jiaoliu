from PIL import Image
import glob
for f in sorted(glob.glob(r"D:\MG\program\ai交流\成品\_work29\png25\*.png")):
    im=Image.open(f); print(f.split("\\")[-1], im.size)
