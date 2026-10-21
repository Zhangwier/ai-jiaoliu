import io
s=io.open('ds.py',encoding='utf-8').read()
for nm in ['def slide','def box','def frame','def _nostyle','def bgrect']:
    i=s.find(nm)
    print('-----',nm); print(s[i:i+430])
