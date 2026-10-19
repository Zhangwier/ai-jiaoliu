import pandas as pd, re, openpyxl, json
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
raw = pd.read_csv(OUT + r'\raw_tables.csv', encoding='utf-8-sig')
sub = raw[(raw['page'] == 21) & (raw['表序号'] == 0)]

grid = {}
for _, r in sub.iterrows():
    grid.setdefault(int(r['行号']), {})[int(r['列号'])] = str(r['文本'])

def thk(s):
    s = s.strip()
    if '=' not in s: return None
    tail = s.split('=')[-1]
    tail = tail.split('(')[0].strip()
    m = re.match(r'^\d+(?:\.\d+)?$', tail)
    return float(tail) if m else None

rows = []
for r in sorted(grid):
    c0 = grid[r].get(0, '')
    t = thk(c0)
    tot = grid[r].get(8, '')
    def f(x):
        try: return float(str(x).replace(',', ''))
        except Exception: return None
    rows.append((r, c0, t, f(tot)))

print('=== 第21页 逐行（厚度解析 + 合计列） ===')
for r, c0, t, tot in rows:
    print(f'  行{r:>2} 项目={c0!r:22} 厚度={t} 合计={tot}')

agg = {}
for r, c0, t, tot in rows:
    if t is None or tot is None: continue
    agg[t] = agg.get(t, 0) + tot
print('\n=== 第21页 按厚度聚合（两 block 相加, 含 Z35 细分） ===')
for t in sorted(agg): print(f'  δ={t:g}  {agg[t]:.1f}')

wb = openpyxl.load_workbook(OUT + r'\final_v2.xlsx', data_only=True)
hq = wb['合计数量']
cur = {}
for r in range(4, 25):
    v = hq.cell(r, 2).value
    if v is None: continue
    q = hq.cell(r, 6).value
    cur[float(v)] = (q if isinstance(q, (int, float)) else None)
print('\n=== 对照：第21页 vs 现有汇总 图纸数量(合计列) ===')
allt = sorted(set(agg) | set(cur))
for t in allt:
    a = agg.get(t); c = cur.get(t)
    if a is None: mark = '【漏项:第21页无,汇总有】'
    elif c is None: mark = '【漏项:汇总无来源】'
    else:
        d = c - a
        mark = f'差 {d:+.1f}' + ('  ≠' if abs(d) > 1 else '  =')
    print(f'  δ={t:g}  第21页={a if a is None else round(a,1)}  汇总={c}   {mark}')

print('\n另有第21页有、汇总无对应行的科目：焊缝重 / 螺栓 M30x200 / M24；部位维度(检修轨道及检查车/主桁/主桁拼接/上层桥面系/伸缩缝/集中排水)')
json.dump({'p21_by_thickness': agg, 'summary_draw_qty': {str(k): v for k, v in cur.items()}},
          open(r'd:\MG\program\ai交流\AI对比\ep3 做表\_ep3_scratch\p21_cmp.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
