import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
hm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT,(20,1)))
n, lab, stats, cent = cv2.connectedComponentsWithStats(hm, 8)
S=400/72
def rows(x0,x1,tag):
    out=[]
    for i in range(1,n):
        x,y,w,h,area = stats[i]
        a,b = x/S,(x+w)/S
        if a<=x0+2 and b>=x1-2:
            out.append(round((y+h/2)/S,1))
    out=sorted(set(out))
    print(tag, out)
rows(161.8,209.4,"LEFT name col")
rows(489.4,536.8,"RIGHT name col")
