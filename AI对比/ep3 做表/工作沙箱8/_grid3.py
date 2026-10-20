import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
print(img.shape)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
# vertical structure: for left table area and right table
def vlines(x0,x1,y0,y1,minfrac=0.5):
    sub = bw[y0:y1, x0:x1]
    s = (sub>0).sum(axis=0)/ (y1-y0)
    idx = [x0+i for i,v in enumerate(s) if v>minfrac]
    # group
    groups=[]
    for i in idx:
        if groups and i-groups[-1][-1]<=3: groups[-1].append(i)
        else: groups.append([i])
    return [(g[0],g[-1]) for g in groups]
def hlines(x0,x1,y0,y1,minfrac=0.5):
    sub = bw[y0:y1, x0:x1]
    s = (sub>0).sum(axis=1)/ (x1-x0)
    idx=[y0+i for i,v in enumerate(s) if v>minfrac]
    groups=[]
    for i in idx:
        if groups and i-groups[-1][-1]<=3: groups[-1].append(i)
        else: groups.append([i])
    return [(g[0],g[-1]) for g in groups]
print("LEFT V:", vlines(950,4300,560,4200))
print("RIGHT V:", vlines(4300,5800,560,2700))
