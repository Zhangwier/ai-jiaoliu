import pandas as pd, json, os, collections
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
df = pd.read_csv(OUT + r'\normalized_tables.csv', encoding='utf-8-sig')
print('cols', list(df.columns))
print('total rows', len(df))
print('构件类别 分布:'); print(df['构件类别'].value_counts(dropna=False).to_string())

zl = df[df['构件类别'] == '主梁'].copy()
print('\n主梁 rows', len(zl))

def num(s):
    return pd.to_numeric(s, errors='coerce')

zl['_thk'] = num(zl['厚'])
zl['_qty'] = num(zl['数量'])
zl['_L'] = num(zl['长']); zl['_W'] = num(zl['宽'])
zl['_H'] = num(zl['单件重'])

print('\n=== 主梁 厚度档分布（数量非空） ===')
th = zl['_thk'].value_counts(dropna=False).sort_index()
print(th.to_string())

print('\n=== 主梁 厚度档 × 来源页（前若干） ===')
for t in sorted(zl['_thk'].dropna().unique()):
    sub = zl[zl['_thk'] == t]
    pages = sorted(set(str(p) for p in sub['来源页'].dropna().unique()))
    print(f'  厚={t:g}  行={len(sub)}  页={pages[:12]}{"..." if len(pages)>12 else ""}')

# missing dims
nd = zl['_L'].isna() | zl['_W'].isna() | zl['_thk'].isna()
print('\n=== 主梁 缺尺寸(长/宽/厚任一缺) ===', int(nd.sum()))
print(' 其中 有数量:', int((nd & zl['_qty'].notna()).sum()))
print(' 其中 校验状态分布:'); print(zl[nd]['校验状态'].value_counts(dropna=False).to_string())
print(' 缺尺寸行 表名示例:'); print(zl[nd]['表名'].value_counts().head(10).to_string())

print('\n=== 主梁 校验状态分布 ===')
print(zl['校验状态'].value_counts(dropna=False).to_string())
fl = zl[zl['校验状态'] == 'fail']
print('fail rows', len(fl))
print(fl[['梁板编号','板件名称','厚','长','宽','数量','单件重','合计重量','备注','来源页','表名']].head(20).to_string())

print('\n=== 全表 表名清单（含工程数量表/材料数量表） ===')
tabs = df['表名'].astype(str)
cand = sorted(set(t for t in tabs if ('数量表' in t or '工程数量' in t)))
for t in cand[:80]:
    print('  ', t)
print('  候选表数', len(cand))
json.dump({'n_rows': len(df)}, open(r'd:\MG\program\ai交流\AI对比\ep3 做表\_ep3_scratch\recon.json','w'), ensure_ascii=False)
