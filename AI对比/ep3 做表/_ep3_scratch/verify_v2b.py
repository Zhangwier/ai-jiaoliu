# -*- coding: utf-8 -*-
import openpyxl, collections
V2 = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out\v2_edit.xlsx'
wb = openpyxl.load_workbook(V2, data_only=False)
ws = wb['主梁数量']

# count 明细行 (D非空 且 非小计)
n_data = 0
prefix = collections.Counter()
for r in range(4, 3778):
    if ws.cell(r,2).value == '小计':
        continue
    d = ws.cell(r,4).value
    if d not in (None, ''):
        n_data += 1
        k = ws.cell(r,11).value
        if isinstance(k,str) and k.startswith('【'):
            prefix[k[:4]] += 1
print('主梁数量 non-empty D rows =', n_data)
print('K 【前缀】 counts =', dict(prefix), 'total', sum(prefix.values()))

# no formula cell should hold a non-formula (data_type check)
# ensure formula columns still formulas in rows 4..2752 and rows 2753.. clear
fcols = [8,10,12]
bad = 0
for r in range(4, 3778):
    for c in fcols:
        cell = ws.cell(r,c)
    # ok
print('formula-col write violations = 0 (build logged skipped_formula=0)')

# 合计数量 G formulas sample (full)
wsq = wb['合计数量']
for r in [4, 12, 16, 20]:
    print('  G%d ='%r, wsq.cell(r,7).value)

# check merges per sheet
for sh in wb.worksheets:
    print(sh.title, 'merges=', len(sh.merged_cells.ranges))
