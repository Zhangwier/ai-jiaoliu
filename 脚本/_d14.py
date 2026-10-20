from PIL import Image
import numpy as np
for name in (r"D:\MG\program\ai交流\成品\_work29\png_orig\page-0001.png", r"D:\MG\program\ai交流\成品\_work29\png25\page-0002.png"):
    im = Image.open(name).convert("RGB")
    a = np.array(im)
    print("##", name, a.shape)
    # teal 0F766E
    m = (np.abs(a[:,:,0].astype(int)-15)<14)&(np.abs(a[:,:,1].astype(int)-118)<14)&(np.abs(a[:,:,2].astype(int)-110)<14)
    ys,xs = np.nonzero(m)
    if len(xs): print("  teal bbox x", xs.min(), xs.max(), "y", ys.min(), ys.max())
    # dark navy 14213D columns
    m2 = (a[:,:,0]<40)&(a[:,:,1]<50)&(a[:,:,2]<80)
    ys2,xs2=np.nonzero(m2)
    print("  navy bbox x", xs2.min(), xs2.max(), "y", ys2.min(), ys2.max())
