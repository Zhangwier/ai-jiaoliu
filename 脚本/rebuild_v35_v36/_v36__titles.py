import io,re
for f in ['open.py','ch1.py','ch2.py','ch3.py','ch4.py']:
    s=io.open(f,encoding='utf-8').read()
    print('=====',f)
    for m in re.finditer(r'^def (p\d+)\(prs, pno\):', s, re.M):
        st=m.end()
        nx=s.find('def ',st)
        blk=s[st: st+900]
        t=re.findall(r'"([^"]{10,70})"', blk)
        print(m.group(1),'|', t[0] if t else '')
