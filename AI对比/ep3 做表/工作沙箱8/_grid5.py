import numpy as np
vm = np.load("work/vm.npy")
seg = vm[560:4210,:]
s = (seg>0).sum(axis=0)
cand = [(i,int(v)) for i,v in enumerate(s) if v>200]
# group
g=[]
for i,v in cand:
    if g and i-g[-1][-1][0]<=4: g[-1].append((i,v))
    else: g.append([(i,v)])
for grp in g:
    xs=[x for x,_ in grp]; vs=[v for _,v in grp]
    print(f"x={sum(xs)/len(xs):7.1f}pt  maxrun={max(vs)}  width={len(xs)}")
