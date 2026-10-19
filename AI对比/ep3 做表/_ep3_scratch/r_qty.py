import pandas as pd
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
raw = pd.read_csv(OUT + r'\raw_tables.csv', encoding='utf-8-sig')
print('raw rows', len(raw), 'tables', raw.groupby(['page','表序号']).ngroups)

# tables whose title or cells mention 工程数量 / 汇总 / 板厚
titles = raw.groupby(['page','表序号'])['table_title'].first().reset_index()
sel = titles[titles['table_title'].astype(str).str.contains('工程数量|汇总|数量表', na=False)]
print('\n=== 含 工程数量/汇总/数量表 的表 ===')
print(sel.to_string())

def show(page, tidx, maxrows=30):
    sub = raw[(raw['page'] == page) & (raw['表序号'] == tidx)]
    print(f'\n--- page {page} 表{tidx} title={sub["table_title"].iloc[0]!r} rows={sub["行号"].max()+1} cols={sub["列号"].max()+1} ---')
    grid = {}
    for _, r in sub.iterrows():
        grid.setdefault(int(r['行号']), {})[int(r['列号'])] = str(r['文本'])
    for rr in sorted(grid)[:maxrows]:
        print(f'  {rr:>2}', [grid[rr].get(c, '') for c in range(0, min(sub['列号'].max() + 1, 14))])

# 工程数量表 pages
for pg, ti in [(4,0),(5,0),(5,1),(4,1)]:
    try: show(pg, ti)
    except Exception as e: print('skip', pg, ti, e)

# any table containing '板厚'
mask = raw['文本'].astype(str).str.contains('板厚|厚度|Q355|Q420', na=False)
cand = raw[mask].groupby(['page','表序号']).size().reset_index(name='n')
print('\n=== 出现 板厚/厚度/Q355/Q420 的表 ===')
print(cand.head(40).to_string())
