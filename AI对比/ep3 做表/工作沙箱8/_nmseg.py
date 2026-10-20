import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
hm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT,(20,1)))
n, lab, stats, cent = cv2.connectedComponentsWithStats(hm, 8)
S=400/72
for i in range(1,n):
    x,y,w,h,area = stats[i]
    a,b = x/S,(x+w)/S
    if 158<=a<=215 and b<=215 and w>20:
        print(f"y={(y+h/2)/S:7.1f}  x {a:6.1f}..{b:6.1f}")
print("---- name col 2 ----")
for i in range(1,n):
    x,y,w,h,area = stats[i]
    a,b = x/S,(x+w)/S
    if 486<=a<=540 and b<=542 and w>20:
        print(f"y={(y+h/2)/S:7.1f}  x {a:6.1f}..{b:6.1f}")
