# -*- coding: utf-8 -*-
"""逐项差异分析：复核重量（公称尺寸/面积公式） vs 图纸单件重"""
import re
from build_data import PARTS

DENS = 7850


def parse(spec):
    s = spec.replace('□', '').replace('x', '×')
    m = re.match(r'^t=(\d+(?:\.\d+)?)mm,A=([\d.]+)m2$', s)
    if m:
        return ('area', None, None, float(m.group(1)), float(m.group(2)))
    n = re.split(r'×', s)
    return ('rect', float(n[0]), float(n[1]), float(n[2]), None)


rows = []
for (grp, part, mat, spec, qty, wdraw, note) in PARTS:
    kind, L, W, T, A = parse(spec)
    if kind == 'rect':
        wt = L / 1000 * W / 1000 * T / 1000 * DENS
    else:
        wt = A * T / 1000 * DENS
    d = wt - wdraw
    rel = d / wdraw * 100
    extra = ''
    if kind == 'area':
        A_impl = wdraw / (T / 1000 * DENS)
        extra = 'Aimpl=%.5f (dA=%+.5f)' % (A_impl, A_impl - A)
    else:
        L_impl = wdraw / (W / 1000 * T / 1000 * DENS) * 1000
        extra = 'Limpl=%.2f (dL=%+.2f)' % (L_impl, L_impl - L)
    rows.append((grp, part, spec, qty, T, wt, wdraw, d, rel, extra))

print('%-8s %-14s %-12s %8s %8s %8s %7s  %s' % ('件号', '部件', '规格', '复核', '图纸', '差', '相对%', '反算'))
for (grp, part, spec, qty, T, wt, wd, d, rel, extra) in rows:
    print('%-8s %-14s %-12s %8.3f %8.1f %8.3f %6.3f%%  %s' % (part, grp, spec, wt, wd, d, rel, extra))

print()
n0 = sum(1 for r in rows if abs(r[7]) < 0.05)
n1 = sum(1 for r in rows if 0.05 <= abs(r[7]) <= 0.1)
n2 = len(rows) - n0 - n1
print('差<0.05kg: %d   0.05-0.1: %d   >0.1: %d  总: %d' % (n0, n1, n2, len(rows)))
print('总复核:%.1f  总图纸:%.1f  差:%.1f (%.3f%%)' % (
    sum(r[5] * r[3] for r in rows), sum(r[6] * r[3] for r in rows),
    sum((r[5] - r[6]) * r[3] for r in rows),
    sum((r[5] - r[6]) * r[3] for r in rows) / sum(r[6] * r[3] for r in rows) * 100))
