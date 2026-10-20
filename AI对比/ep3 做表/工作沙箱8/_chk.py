import numpy as np
from PIL import Image
im = Image.open(r"work\p3_400.png").convert("L")
a = np.array(im)
for y0,y1 in [(565,620),(700,760)]:
    sub = a[y0:y1,:]
    darkcol = (sub<128).sum(axis=0)
    idx=[i for i,v in enumerate(darkcol) if v> (y1-y0)*0.85]
    g=[]
    for i in idx:
        if g and i-g[-1][-1]<=3: g[-1].append(i)
        else: g.append([i])
    print(y0,y1,[(round(sum(x)/len(x),1), round(sum(x)/len(x)/5.5556,1)) for x in g])
