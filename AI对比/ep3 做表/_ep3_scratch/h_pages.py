import json, pandas as pd
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
log = json.load(open(OUT + r'\e2_log.json', encoding='utf-8'))
print('=== 图纸数量 by bucket (来源页) ===')
for b in log['合计数量']['draw_qty_by_bucket']:
    print(json.dumps(b, ensure_ascii=False))

df = pd.read_csv(OUT + r'\normalized_tables.csv', encoding='utf-8-sig')
zl = df[df['构件类别'] == '主梁'].copy()
def num(s): return pd.to_numeric(s, errors='coerce')
zl['_t'] = num(zl['厚']); zl['_L'] = num(zl['长']); zl['_W'] = num(zl['宽'])
det = zl[zl['_L'].notna() & zl['_W'].notna() & zl['_t'].notna()]
print('\n=== 复核明细来源页 by thickness ===')
for t in sorted(det['_t'].unique()):
    sub = det[det['_t'] == t]
    pages = sorted(set(str(int(p)) for p in sub['来源页'].dropna().unique() if str(p).replace(',', '').replace(' ', '').isdigit()))
    print(f'  厚={t:g}  行={len(sub)}  页={",".join(pages[:20])}{"..." if len(pages)>20 else ""}')
print('明细总行', len(det))
