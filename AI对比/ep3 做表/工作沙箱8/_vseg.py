import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
vk = cv2.getStructuringElement(cv2.MORPH_RECT,(1,40))
vm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, vk)
n, lab, stats, cent = cv2.connectedComponentsWithStats(vm, 8)
S=400/72
lines=[]
for i in range(1,n):
    x,y,w,h,area = stats[i]
    if h>100:
        lines.append((round((x+w/2)/S,1), round(y/S,1), round((y+h)/S,1), round(h/S,1)))
lines.sort()
for l in lines: print(f"x={l[0]:7.1f}  y {l[1]:6.1f}..{l[2]:6.1f}  h={l[3]:6.1f}")
