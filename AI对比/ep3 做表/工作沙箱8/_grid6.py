import numpy as np
vm = np.load("work/vm.npy")
seg = vm[560:4210, 400:3000]
s = (seg>0).sum(axis=0)
g=[]
for i,v in enumerate(s):
    if v>150:
        if g and i-g[-1][-1][0]<=4: g[-1].append((i,int(v)))
        else: g.append([(i,int(v))])
for grp in g:
    xs=[x for x,_ in grp]; vs=[v for _,v in grp]
    print(f"x={400+sum(xs)/len(xs):7.1f}px  {((400+sum(xs)/len(xs))/5.5556):7.1f}pt maxrun={max(vs)} w={len(xs)}")
