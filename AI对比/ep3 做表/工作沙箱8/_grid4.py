import cv2, numpy as np
img = cv2.imread(r"work\p3_400.png", cv2.IMREAD_GRAYSCALE)
bw = cv2.threshold(img,180,255,cv2.THRESH_BINARY_INV)[1]
# horizontal mask
hk = cv2.getStructuringElement(cv2.MORPH_RECT,(150,1))
hk2 = cv2.getStructuringElement(cv2.MORPH_RECT,(60,1))
hm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, hk2)
vk = cv2.getStructuringElement(cv2.MORPH_RECT,(1,60))
vm = cv2.morphologyEx(bw, cv2.MORPH_OPEN, vk)
np.save("work/hm.npy", hm); np.save("work/vm.npy", vm)
rows = np.where(hm.sum(axis=1)>500)[0]
cols = np.where(vm.sum(axis=0)>500)[0]
def group(a):
    g=[]
    for i in a:
        if g and i-g[-1][-1]<=3: g[-1].append(i)
        else: g.append([i])
    return [(int(x[0]),int(x[-1])) for x in g]
print("H rows:", group(rows))
print("V cols:", group(cols))
