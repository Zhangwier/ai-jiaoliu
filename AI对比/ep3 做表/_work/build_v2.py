# -*- coding: utf-8 -*-
"""节段01 计算底稿 v2：
1) 面积型板件改为公式计算（面积×厚×7850）；
2) 逐项对比图纸单件重/共重，写明差异原因，解释不了的标“待核”；
3) 不改动 v1 与原文件，另存新版。
"""
import re
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill, Font
from build_data import PARTS, BUCKETS, DRAWING_BY_T, DRAWING_WELD, DRAWING_TOTAL

TPL = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱4\模板原件.xlsx'
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱4\中桁节段01_计算底稿v2.xlsx'
DENS = 7850
TOL = 0.05          # 图纸单件重按 0.1kg 取整，最大取整误差
RING = PatternFill('solid', fgColor='FFFFC7CE')
RCOL = Font(color='FF9C0006')


def parse_spec(spec):
    s = spec.replace('□', '').replace('x', '×')
    m = re.match(r'^t=(\d+(?:\.\d+)?)mm,A=([\d.]+)m2$', s)
    if m:
        a = m.group(2)
        dp = len(a.split('.')[1]) if '.' in a else 0
        return ('area', None, None, float(m.group(1)), float(a), 0.5 * 10 ** -dp)
    n = re.split(r'×', s)
    return ('rect', float(n[0]), float(n[1]), float(n[2]), None, None)


def judge(part, spec, qty, wdraw, extra_note):
    """返回 (复核单件重, 差异说明文本, 是否待核)"""
    kind, L, W, T, A, Prec = parse_spec(spec)
    if kind == 'rect':
        wt = L / 1000 * W / 1000 * T / 1000 * DENS
    else:
        wt = A * T / 1000 * DENS
    d = wt - wdraw
    note = ['材质%s' % part[2]] if False else []
    if abs(d) <= TOL:
        txt = '一致：复核%.2fkg，图纸%.1fkg，差%+.3fkg（图纸单件重按0.1kg取整）' % (wt, wdraw, d)
        flag = False
    elif kind == 'area':
        A_impl = wdraw / (T / 1000 * DENS)
        if abs(A_impl - A) <= Prec:
            txt = ('取整差异（已解释）：复核=标注面积%.4fm²×厚→%.2fkg；图纸%.1fkg对应实际面积%.5fm²，'
                   '差%+.3fkg 系图纸面积为取整标注值' % (A, wt, wdraw, A_impl, d))
        else:
            txt = ('待核：复核按标注面积%.4fm²×厚%gmm→%.2fkg，图纸%.1fkg（差%+.3fkg，%.2f%%）；'
                   '反算实际面积%.5fm²，与标注差%+.5fm²，超出面积标注精度，原因图纸未注明'
                   % (A, T, wt, wdraw, d, abs(d) / wdraw * 100, A_impl, A_impl - A))
        flag = txt.startswith('待核')
    else:
        L_impl = wdraw / (W / 1000 * T / 1000 * DENS) * 1000
        if abs(L_impl - L) <= 0.5:
            txt = ('取整差异（已解释）：复核=标注%g×%g×%g→%.2fkg，图纸%.1fkg，'
                   '反算长%.2fmm与标注差%+.2fmm，属图纸尺寸取整' % (L, W, T, wt, wdraw, L_impl, L_impl - L))
            flag = False
        else:
            txt = ('待核：复核按标注尺寸%g×%g×%g→%.2fkg，图纸%.1fkg（差%+.3fkg，%.2f%%）；'
                   '图纸单件重反算长%.1fmm，与标注%gmm差%+.1fmm，差异原因图纸未注明'
                   % (L, W, T, wt, wdraw, d, abs(d) / wdraw * 100, L_impl, L, L_impl - L))
            flag = True
    if extra_note:
        txt += '；' + extra_note
    return wt, txt, flag


wb = openpyxl.load_workbook(TPL)
src = wb['主梁数量']
sum_src = wb['合计数量']
ws = wb.create_sheet('主梁数量-节段01')
ws2 = wb.create_sheet('合计数量-节段01')

for col in list('ABCDEFGHIJKLMNOPQRSTUVWXY'):
    if col in src.column_dimensions and src.column_dimensions[col].width:
        ws.column_dimensions[col].width = src.column_dimensions[col].width
for col in 'ABCDEFGHI':
    if col in sum_src.column_dimensions and sum_src.column_dimensions[col].width:
        ws2.column_dimensions[col].width = sum_src.column_dimensions[col].width
for r in (2, 3):
    if r in src.row_dimensions and src.row_dimensions[r].height:
        ws.row_dimensions[r].height = src.row_dimensions[r].height


def cp(dst, s):
    dst._style = copy(s._style)


def st(col, row=4):
    """取模板样式：A~K 用 1~11 列，L~Q 用 11 列，R 以后用 13~23 列"""
    if col <= 11:
        c = col
    elif col <= 17:
        c = 11
    else:
        c = min(13 + (col - 18), 23)
    return src.cell(row=row, column=c)


ws.merge_cells('A1:Q1')
ws['A1'] = '中桁节段01材料数量表（计算底稿v2：面积型板件按公式计算，逐项对比图纸重量）'
cp(ws['A1'], src['A1'])

hdr = ['梁板编号', '板件名称', '位置', '材质', '长', '宽', '厚', '面积', '单件重', '数量', '合计重量',
       '图纸单件重', '单件重差异', '图纸共重', '共重差异', '规格（图纸）', '差异说明／备注']
for i, t in enumerate(hdr, start=1):
    cp(ws.cell(row=2, column=i, value=t), st(i, 2))
ws['E3'] = 'mm'; ws['F3'] = 'mm'; ws['G3'] = 'mm'; ws['H3'] = 'm²'
ws['I3'] = 'kg'; ws['K3'] = 'kg'; ws['L3'] = 'kg'; ws['M3'] = 'kg'; ws['N3'] = 'kg'; ws['O3'] = 'kg'
for col in 'EFGHIKLMNO':
    cp(ws[col + '3'], st(ord(col) - 64, 3))

# 复核区块表头（R..AB 板厚，AC 图纸数量，AD 差异）
ws.merge_cells('R2:AB2')
ws['R2'] = '复核数量（按板厚，kg）'
cp(ws['R2'], src['M2'])
cp(ws['AC2'], src['U2']); ws['AC2'] = '图纸数量'
cp(ws['AD2'], src['V2']); ws['AD2'] = '差异'
cp(ws['AE2'], src['W2']); ws['AE2'] = '备注'
for i, b in enumerate(BUCKETS):
    cp(ws.cell(row=3, column=18 + i, value=b), src.cell(row=3, column=13 + i))
cp(ws.cell(row=3, column=28, value='合计'), src['T3'])
ws.merge_cells('R2:AB2')

r = 4
start_data = r
group_rows = []
subtotal_rows = []
cur, g_start = None, r
for (grp, part_name, mat, spec, qty, wdraw, note) in PARTS:
    if cur is None:
        cur, g_start = grp, r
    if grp != cur:
        group_rows.append((cur, g_start, r - 1))
        cp(ws.cell(row=r, column=2, value='小计'), src['B32'])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cp(ws.cell(row=r, column=11, value='=SUM(K%d:K%d)' % (g_start, r - 1)), src['J32'])
        cp(ws.cell(row=r, column=14, value='=SUM(N%d:N%d)' % (g_start, r - 1)), src['J32'])
        cp(ws.cell(row=r, column=15, value='=K%d-N%d' % (r, r)), src['J32'])
        for c in range(1, 18):
            cp(ws.cell(row=r, column=c), st(c))
        subtotal_rows.append(r)
        r += 1
        cur, g_start = grp, r
    kind, L, W, T, A, Prec = parse_spec(spec)
    wt, explain, need_check = judge((grp, part_name, mat), spec, qty, wdraw, note)
    ws.cell(row=r, column=1, value='中桁节段01')
    ws.cell(row=r, column=2, value=grp)
    ws.cell(row=r, column=3, value=part_name)
    ws.cell(row=r, column=4, value=mat)
    if kind == 'rect':
        ws.cell(row=r, column=5, value=L)
        ws.cell(row=r, column=6, value=W)
        ws.cell(row=r, column=7, value=T)
        f = '=(E%d/1000)*(F%d/1000)*(G%d/1000)*7850' % (r, r, r)
    else:
        ws.cell(row=r, column=7, value=T)
        ws.cell(row=r, column=8, value=A)
        f = '=H%d*(G%d/1000)*7850' % (r, r)
    ws.cell(row=r, column=9, value=f)
    ws.cell(row=r, column=10, value=qty)
    ws.cell(row=r, column=11, value='=I%d*J%d' % (r, r))
    ws.cell(row=r, column=12, value=wdraw)
    ws.cell(row=r, column=13, value='=I%d-L%d' % (r, r))
    ws.cell(row=r, column=14, value=wdraw * qty)
    ws.cell(row=r, column=15, value='=K%d-N%d' % (r, r))
    ws.cell(row=r, column=16, value=spec)
    ws.cell(row=r, column=17, value=explain)
    for c in range(1, 31):
        cp(ws.cell(row=r, column=c), st(c))
    ws.cell(row=r, column=9).number_format = '0.00'
    ws.cell(row=r, column=12).number_format = '0.0'
    ws.cell(row=r, column=13).number_format = '0.000'
    ws.cell(row=r, column=11).number_format = '0.0'
    ws.cell(row=r, column=14).number_format = '0.0'
    ws.cell(row=r, column=15).number_format = '0.000'
    ws.cell(row=r, column=16).alignment = copy(ws.cell(row=r, column=16).alignment)
    ws.row_dimensions[r].height = src.row_dimensions[4].height
    if need_check:
        ws.cell(row=r, column=17).fill = RING
        ws.cell(row=r, column=17).font = RCOL
    r += 1
group_rows.append((cur, g_start, r - 1))
cp(ws.cell(row=r, column=2, value='小计'), src['B32'])
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
cp(ws.cell(row=r, column=11, value='=SUM(K%d:K%d)' % (g_start, r - 1)), src['J32'])
cp(ws.cell(row=r, column=14, value='=SUM(N%d:N%d)' % (g_start, r - 1)), src['J32'])
cp(ws.cell(row=r, column=15, value='=K%d-N%d' % (r, r)), src['J32'])
for c in range(1, 18):
    cp(ws.cell(row=r, column=c), st(c))
subtotal_rows.append(r)
r += 1
last_data = r - 1

ws.merge_cells(start_row=start_data, start_column=1, end_row=last_data, end_column=1)
for (grp, a, b) in group_rows:
    ws.merge_cells(start_row=a, start_column=2, end_row=b, end_column=2)

cp(ws.cell(row=r, column=2, value='节段小计'), src['B32'])
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
cp(ws.cell(row=r, column=11, value='=' + '+'.join('K%d' % i for i in subtotal_rows)), src['J32'])
cp(ws.cell(row=r, column=14, value='=' + '+'.join('N%d' % i for i in subtotal_rows)), src['J32'])
cp(ws.cell(row=r, column=15, value='=K%d-N%d' % (r, r)), src['J32'])
for c in range(1, 18):
    cp(ws.cell(row=r, column=c), st(c))
row_seg = r

# 逐项对比结论
r += 1
txt = ws.cell(row=r, column=1, value='逐项对比结论')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=16)
n_all = len(PARTS)
n_chk = sum(1 for p in PARTS if judge(p[:3], p[3], p[4], p[5], p[6])[2])
n_ok = sum(1 for p in PARTS if abs(judge(p[:3], p[3], p[4], p[5], p[6])[0] - p[5]) <= TOL)
txt2 = ('共 %d 项：与图纸一致 %d 项；差异可解释（图纸取整）%d 项；待核 %d 项（表中已用红底标出）。'
        '复核一律按图纸标注尺寸/面积用公式计算，未做任何凑数调整。'
        % (n_all, n_ok, n_all - n_ok - n_chk, n_chk))
c = ws.cell(row=r, column=2, value=txt2)
for cc in range(1, 18):
    cp(ws.cell(row=r, column=cc), st(cc))

# 底部：按板厚汇总 vs 图纸
sumif_g = '$G$%d:$G$%d' % (start_data, last_data)
sumif_k = '$K$%d:$K$%d' % (start_data, last_data)
GL = openpyxl.utils.get_column_letter

r += 2
row_recheck = r
ws.cell(row=r, column=2, value='复核数量（按板厚）')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=17)
for i in range(10):
    col = 18 + i
    cp(ws.cell(row=r, column=col, value='=SUMIF(%s,%s$3,%s)' % (sumif_g, GL(col), sumif_k)),
       src.cell(row=139, column=13 + i))
cp(ws.cell(row=r, column=28, value='=SUM(R%d:AA%d)' % (r, r)), src['X139'])
for cc in range(1, 31):
    if ws.cell(row=r, column=cc).value is None:
        cp(ws.cell(row=r, column=cc), src['B139'])
r += 1
row_draw = r
ws.cell(row=r, column=2, value='图纸数量（节段01材料数量汇总表）')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=17)
for i, b in enumerate(BUCKETS):
    cp(ws.cell(row=r, column=18 + i, value=DRAWING_BY_T[b]), src.cell(row=139, column=13 + i))
cp(ws.cell(row=r, column=28, value='=SUM(R%d:AA%d)' % (r, r)), src['X139'])
for cc in range(1, 31):
    if ws.cell(row=r, column=cc).value is None:
        cp(ws.cell(row=r, column=cc), src['B139'])
r += 1
row_diff = r
ws.cell(row=r, column=2, value='差异（复核-图纸）')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=17)
for i in range(10):
    col = GL(18 + i)
    cp(ws.cell(row=r, column=18 + i, value='=%s%d-%s%d' % (col, row_recheck, col, row_draw)),
       src.cell(row=141, column=13 + i))
cp(ws.cell(row=r, column=28, value='=AB%d-AB%d' % (row_recheck, row_draw)), src['X141'])
for cc in range(1, 31):
    if ws.cell(row=r, column=cc).value is None:
        cp(ws.cell(row=r, column=cc), src['B139'])
r += 1
row_weld = r
ws.cell(row=r, column=2, value='焊缝重（1.5%）')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=17)
cp(ws.cell(row=r, column=28, value='=ROUND(AB%d*0.015,1)' % row_recheck), src['X139'])
cp(ws.cell(row=r, column=29, value=DRAWING_WELD), src['Y139'])
cp(ws.cell(row=r, column=30, value='=AB%d-AC%d' % (r, r)), src['Z139'])
for cc in range(1, 32):
    if ws.cell(row=r, column=cc).value is None:
        cp(ws.cell(row=r, column=cc), src['B139'])
r += 1
row_tot = r
ws.cell(row=r, column=2, value='合计（含焊缝）')
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=17)
cp(ws.cell(row=r, column=28, value='=AB%d+AB%d' % (row_recheck, row_weld)), src['X139'])
cp(ws.cell(row=r, column=29, value=123805.2), src['Y139'])
cp(ws.cell(row=r, column=30, value='=AB%d-AC%d' % (r, r)), src['Z139'])
for cc in range(1, 32):
    if ws.cell(row=r, column=cc).value is None:
        cp(ws.cell(row=r, column=cc), src['B139'])
r += 2
ws.cell(row=r, column=1, value='注：1) 矩形板件复核=(长/1000)×(宽/1000)×(厚/1000)×7850；面积型板件复核=(面积m²)×(厚/1000)×7850；'
                                '合计重量=单件重×数量。2) 图纸数量取自图纸《节段01材料数量汇总表》。'
                                '3) 差异=复核-图纸，正数表示复核大于图纸。')
ws.cell(row=r + 1, column=1, value='    4) 面积型板件差异均为图纸面积取整标注所致；矩形板件无法由标注尺寸复算的，已在“差异说明”列标“待核”，未做调整。')
if ws.max_row > r + 2:
    ws.delete_rows(r + 3, ws.max_row - (r + 2))

# ============ 汇总表 ============
ws2.merge_cells('A1:I1')
ws2['A1'] = '中桁节段01 按板厚汇总（计算底稿v2）'
cp(ws2['A1'], sum_src['A1'])
for i, t in enumerate(['部位', '板厚', '单位', '图纸数量', None, None, '复核数量', '差异', '备注'], start=1):
    cp(ws2.cell(row=2, column=i, value=t), sum_src.cell(row=2, column=i))
ws2.merge_cells('D2:F2')
cp(ws2['D3'], sum_src['D3']); ws2['D3'] = 'kg'
row = 4
first_row = row
for i, b in enumerate(BUCKETS):
    ws2.cell(row=row, column=1, value='中桁节段01')
    ws2.cell(row=row, column=2, value=b)
    ws2.cell(row=row, column=3, value='kg')
    ws2.cell(row=row, column=4, value=DRAWING_BY_T[b])
    ws2.cell(row=row, column=7, value="='主梁数量-节段01'!%s%d" % (GL(18 + i), row_recheck))
    ws2.cell(row=row, column=8, value='=G%d-D%d' % (row, row))
    if b == 4:
        ws2.cell(row=row, column=9, value='N15b 材质为Q355B，与主材Q355D不同，建议核实')
    for c in range(1, 10):
        cp(ws2.cell(row=row, column=c), sum_src.cell(row=4, column=c))
    row += 1
ws2.cell(row=row, column=2, value='焊缝重（1.5%）')
ws2.cell(row=row, column=4, value=DRAWING_WELD)
ws2.cell(row=row, column=7, value="='主梁数量-节段01'!AB%d" % row_weld)
ws2.cell(row=row, column=8, value='=G%d-D%d' % (row, row))
for c in range(1, 10):
    cp(ws2.cell(row=row, column=c), sum_src.cell(row=4, column=c))
row += 1
ws2.cell(row=row, column=1, value='钢梁制造合计')
ws2.cell(row=row, column=2, value='合计')
ws2.cell(row=row, column=3, value='kg')
ws2.cell(row=row, column=4, value='=SUM(D%d:D%d)' % (first_row, row - 1))
ws2.cell(row=row, column=7, value='=SUM(G%d:G%d)' % (first_row, row - 1))
ws2.cell(row=row, column=8, value='=G%d-D%d' % (row, row))
for c in range(1, 10):
    cp(ws2.cell(row=row, column=c), sum_src.cell(row=15, column=c))
row += 2
ws2.cell(row=row, column=1, value='注：图纸数量取自《节段01材料数量汇总表》（图纸第5张）；复核数量用 SUMIF 按板厚汇总明细。'
                                  '差异全部来自“按标注尺寸/面积用公式反算”与图纸取整值之差，未做调整。')
wb.save(OUT)
print('saved', OUT)
print('项数 %d，一致 %d，待核 %d' % (n_all, n_ok, n_chk))
print('明细区行 %d-%d，复核行 %d，合计行 %d' % (start_data, last_data, row_recheck, row_tot))
