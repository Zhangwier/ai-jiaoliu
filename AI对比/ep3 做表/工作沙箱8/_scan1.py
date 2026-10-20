import numpy as np
from PIL import Image
im = Image.open(r"work\page3_150.png").convert("L")
a = np.array(im)
print(a.shape)
dark = (a<128)
# left table x range guess: find columns with long vertical dark runs in upper half
colsum = dark[200:1400,:].sum(axis=0)
cands = [i for i,v in enumerate(colsum) if v>800]
print("vert cols >800:", cands[:40], "...", cands[-20:])
