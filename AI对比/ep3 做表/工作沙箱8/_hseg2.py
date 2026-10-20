import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
hk = cv2.getStructuringElement(cv2.MORPH_RECT,(40,1))
hm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, hk)
n, lab, stats, cent = cv2.connectedComponentsWithStats(hm, 8)
S=400/72
segs=[]
for i in range(1,n):
    x,y,w,h,area = stats[i]
    if w>50:
        segs.append(((y+h/2)/S, x/S, (x+w)/S))
segs.sort()
for s in segs[80:]:
    print(f"y={s[0]:7.1f}  x {s[1]:6.1f}..{s[2]:6.1f}")
