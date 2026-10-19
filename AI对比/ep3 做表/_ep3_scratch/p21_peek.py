import pandas as pd
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'

man = pd.read_csv(OUT + r'\page_manifest.csv', encoding='utf-8-sig')
print('=== manifest 页 18..24 ===')
print(man[man['page_index'].between(18, 24)][['page_index','page_no','char_count','is_table_page','table_likelihood','table_title']].to_string())

raw = pd.read_csv(OUT + r'\raw_tables.csv', encoding='utf-8-sig')
print('\n=== raw_tables 中 page=21 的表 ===')
sub = raw[raw['page'].isin([20, 21])]
for (pg, ti), g in sub.groupby(['page','表序号']):
    print(f'--- page={pg} 表序号={ti} title={g["table_title"].iloc[0]!r} rows={int(g["行号"].max())+1} cols={int(g["列号"].max())+1} cells={len(g)} ---')
    grid = {}
    for _, r in g.iterrows():
        grid.setdefault(int(r['行号']), {})[int(r['列号'])] = str(r['文本'])
    nc = int(g['列号'].max()) + 1
    for rr in sorted(grid):
        print(f'  {rr:>2}|', ' | '.join(grid[rr].get(c, '') for c in range(nc)))
