import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
hm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT,(20,1)))
n, lab, stats, cent = cv2.connectedComponentsWithStats(hm, 8)
S=400/72
print("== name col right table (489.4..536.8) ==")
for i in range(1,n):
    x,y,w,h,area = stats[i]
    a,b = x/S,(x+w)/S
    if a<=491 and b>=535:
        print(f"y={(y+h/2)/S:7.1f}  x {a:6.1f}..{b:6.1f}  len={b-a:6.1f}")
print("== full-width rows ==")
for i in range(1,n):
    x,y,w,h,area = stats[i]
    a,b = x/S,(x+w)/S
    if a<=210 and b>=815:
        print(f"y={(y+h/2)/S:7.1f}")
