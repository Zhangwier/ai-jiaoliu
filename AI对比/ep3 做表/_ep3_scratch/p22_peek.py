import pandas as pd, openpyxl
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
raw = pd.read_csv(OUT + r'\raw_tables.csv', encoding='utf-8-sig')
print('=== raw_tables page=22 ===')
sub = raw[raw['page'] == 22]
for (pg, ti), g in sub.groupby(['page','表序号']):
    print(f'--- page={pg} 表序号={ti} title={g["table_title"].iloc[0]!r} rows={int(g["行号"].max())+1} cols={int(g["列号"].max())+1} cells={len(g)} ---')
    grid = {}
    for _, r in g.iterrows():
        grid.setdefault(int(r['行号']), {})[int(r['列号'])] = str(r['文本'])
    nc = int(g['列号'].max()) + 1
    for rr in sorted(grid):
        print(f'  {rr:>2}|', ' | '.join(grid[rr].get(c, '') for c in range(nc)))

wb = openpyxl.load_workbook(OUT + r'\final_v2.xlsx', data_only=True)
hq = wb['合计数量']
print('\n=== 现有汇总 合计数量 (final_v2, 值) ===')
for r in range(1, hq.max_row + 1):
    print(r, [hq.cell(r, c).value for c in range(1, 10)])
zl = wb['主梁数量']
print('\n=== 主梁数量 汇总块 3778..3781 ===')
for r in range(3778, 3782):
    print(r, [zl.cell(r, c).value for c in range(1, 24)])
