# -*- coding: utf-8 -*-
"""新沙高架桥 主桥钢结构 计算底稿生成（沿用潭州高架桥模板的 sheet 结构/公式/格式）"""
import json, re
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

D = json.load(open("sheets_final.json", encoding="utf-8"))
DECK = json.load(open("deck_trans.json", encoding="utf-8"))

TH = [4, 6, 8, 10, 12, 14, 16, 20, 24, 28, 32, 36, 40, 44, 48, 50, 60, 80, 100]
MATS = ["Q355B", "Q355D", "Q420D"]

# ---------- 图纸数量（主桥上构工程数量表 一/二） ----------
DWG = {
    ("主桁", "Q420D"): {100: 8737.1, 80: 6989.7, 60: 2458.8, 48: 191337.7 + 56689.0,
                        44: 149564.2 + 96292.7, 40: 307680.3, 36: 3283.8, 32: 275007.2,
                        16: 27946.9, 14: 2518.5},
    ("主桁", "Q355D"): {100: 3815.1, 80: 3052.1, 48: 33973.3, 44: 140063.0,
                        40: 314452.2 + 182732.3, 36: 521193.4, 32: 398720.8, 28: 60106.8,
                        24: 301501.0, 20: 423735.0, 16: 115785.1, 14: 56947.4},
    ("主桁", "Q355B"): {20: 311.1},
    ("主桁拼接", "Q420D"): {48: 1633.4, 44: 10164.6, 40: 33604.5, 36: 8986.5,
                            32: 10990.2, 28: 12062.2, 24: 11198.7},
    ("主桁拼接", "Q355D"): {44: 10730.2, 36: 10581.6, 32: 9954.8, 28: 45262.6,
                            24: 43378.8, 20: 22910.7, 16: 9958.7, 14: 554.4,
                            12: 2006.1, 8: 1352.2},
    ("上层桥面系", "Q355D"): {24: 545983.2, 20: 431122.0, 16: 1867.2,
                              14: 789637.2, 12: 191453.0, 10: 40128.0},
    ("下层桥面系", "Q355D"): {60: 9646.4, 50: 4396.8, 44: 9560.8, 28: 16076.8,
                              24: 495891.2, 20: 69469.6, 16: 5013.6, 14: 2052399.0,
                              12: 102008.0, 10: 35164.8, 8: 823554.0, 6: 3898.8, 4: 3898.8},
}

thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_BODY = Font(name="宋体", size=9)
F_HEAD = Font(name="宋体", size=9, bold=True)
AL = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HEAD
        cell.alignment = AL
        cell.border = BORDER


def style_range(ws, r1, r2, ncol):
    for r in range(r1, r2 + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = F_BODY
            cell.border = BORDER
            if cell.alignment.horizontal is None:
                cell.alignment = AL


def write_detail(wb, title, rows, colA):
    """明细表：节段/板件/编号/材质/长/宽/厚/单件重/数量/合计/备注 + 按厚度汇总"""
    ws = wb.create_sheet(title)
    ncol = 11
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.cell(row=1, column=1, value=f"{title}  板件计算表（新沙高架桥 K6+418.3 主桥上构）").font = Font(name="宋体", size=12, bold=True)
    ws.cell(row=1, column=1).alignment = AL
    heads = [colA, "板件名称", "编号", "材质", "长(mm)", "宽(mm)", "厚(mm)", "单件重(kg)", "数量", "合计重量(kg)", "备注"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=2, column=i, value=h)
    for i, u in enumerate(["", "", "", "", "mm", "mm", "mm", "kg", "件", "kg", ""], 1):
        ws.cell(row=3, column=i, value=u if u else None)
    style_header(ws, 2, ncol)
    style_header(ws, 3, ncol)

    r = 4
    for d in rows:
        ws.cell(row=r, column=1, value=d.get("seg"))
        ws.cell(row=r, column=2, value=d.get("part"))
        ws.cell(row=r, column=3, value=d.get("code"))
        ws.cell(row=r, column=4, value=d.get("mat"))
        ws.cell(row=r, column=5, value=d.get("L"))
        ws.cell(row=r, column=6, value=d.get("W"))
        ws.cell(row=r, column=7, value=d.get("t"))
        ws.cell(row=r, column=8, value=f"=(E{r}/1000)*(F{r}/1000)*(G{r}/1000)*7850")
        ws.cell(row=r, column=9, value=d.get("qty"))
        ws.cell(row=r, column=10, value=f"=H{r}*I{r}")
        ws.cell(row=r, column=11, value=d.get("spec") or d.get("note") or None)
        r += 1
    last = r - 1
    style_range(ws, 4, last, ncol)
    for rr in range(4, last + 1):
        ws.cell(row=rr, column=8).number_format = "0.00_ "
        ws.cell(row=rr, column=10).number_format = "0.0_ "

    # 按厚度汇总
    sr = last + 2
    ws.cell(row=sr, column=1, value="按板厚汇总（kg）").font = F_HEAD
    ws.cell(row=sr + 1, column=1, value="材质")
    for i, t in enumerate(TH):
        ws.cell(row=sr + 1, column=2 + i, value=t)
    ws.cell(row=sr + 1, column=2 + len(TH), value="合计")
    style_header(ws, sr + 1, 2 + len(TH))
    mat_rows = {}
    for j, m in enumerate(MATS):
        rr = sr + 2 + j
        mat_rows[m] = rr
        ws.cell(row=rr, column=1, value=m)
        for i, t in enumerate(TH):
            col = 2 + i
            ws.cell(row=rr, column=col,
                    value=f'=SUMIFS($J$4:$J${last},$D$4:$D${last},$A{rr},$G$4:$G${last},{get_column_letter(col)}${sr + 1})')
        ws.cell(row=rr, column=2 + len(TH), value=f"=SUM({get_column_letter(2)}{rr}:{get_column_letter(1 + len(TH))}{rr})")
        for c in range(1, 2 + len(TH)):
            ws.cell(row=rr, column=c).number_format = "0.0_ " if c > 1 else "General"
            ws.cell(row=rr, column=c).font = F_BODY
            ws.cell(row=rr, column=c).border = BORDER
            ws.cell(row=rr, column=c).alignment = AL
    # 复核合计行
    tr = sr + 2 + len(MATS)
    ws.cell(row=tr, column=1, value="合计")
    for i in range(len(TH) + 1):
        col = 2 + i
        L = get_column_letter(col)
        ws.cell(row=tr, column=col, value=f"=SUM({L}{sr + 2}:{L}{sr + 1 + len(MATS)})")
        ws.cell(row=tr, column=col).number_format = "0.0_ "
        ws.cell(row=tr, column=col).font = F_HEAD
        ws.cell(row=tr, column=col).border = BORDER
        ws.cell(row=tr, column=col).alignment = AL
    ws.cell(row=tr, column=1).font = F_HEAD
    ws.cell(row=tr, column=1).border = BORDER
    ws.cell(row=tr, column=1).alignment = AL

    widths = [13, 15, 10, 7, 9, 8, 6, 9, 6, 10, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(len(TH)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 8
    ws.freeze_panes = "E4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.print_title_rows = "2:3"
    return ws, mat_rows, last, sr + 1, sr + 2 + len(MATS)


wb = Workbook()
wb.remove(wb.active)
ws_main, main_rows, main_last, main_th, main_sum = write_detail(wb, "主梁数量", D["main"], "节段")
ws_hb, hb_rows, hb_last, hb_th, hb_sum = write_detail(wb, "横隔板", D["hb"], "节段")
ws_sp, sp_rows, sp_last, sp_th, sp_sum = write_detail(wb, "主桁拼接", D["spl"], "节段")
deck_rows = [dict(seg=d["seg"], part=d["part"], code=d["code"], mat="Q355D",
                  L=None, W=None, t=d["t"], qty=d["qty"], spec=d["spec"]) for d in DECK]
ws_dk, dk_rows, dk_last, dk_th, dk_sum = write_detail(wb, "桥面系", deck_rows, "部位")

# 桥面系单件重按图纸取用
for i, d in enumerate(DECK):
    r = 4 + i
    ws_dk.cell(row=r, column=8, value=round(d["unit"], 2))
    ws_dk.cell(row=r, column=10, value=f"=H{r}*I{r}")
    ws_dk.cell(row=r, column=5, value=None)
    ws_dk.cell(row=r, column=6, value=None)
    ws_dk.cell(row=r, column=11, value=d["spec"])

# ---------- 合计数量 ----------
ws = wb.create_sheet("合计数量", 0)
ws.merge_cells("A1:I1")
ws.cell(row=1, column=1, value="新沙高架桥主桥上构钢结构数量复核汇总表（K6+418.3）").font = Font(name="宋体", size=14, bold=True)
ws.cell(row=1, column=1).alignment = AL
ws.merge_cells("A2:A3"); ws.merge_cells("B2:B3"); ws.merge_cells("C2:C3")
ws.merge_cells("D2:F2"); ws.merge_cells("G2:G3"); ws.merge_cells("H2:H3"); ws.merge_cells("I2:I3")
ws["A2"] = "部位"; ws["B2"] = "板厚"; ws["C2"] = "单位"
ws["D2"] = "图纸数量(kg)"; ws["D3"] = "Q355D"; ws["E3"] = "Q420D"; ws["F3"] = "合计"
ws["G2"] = "复核数量(kg)"; ws["H2"] = "差异"; ws["I2"] = "备注"
for row in (2, 3):
    for c in range(1, 10):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HEAD
        cell.alignment = AL
        cell.border = BORDER

parts = [("主桁", "主梁数量", "横隔板"), ("主桁拼接", "主桁拼接", None),
         ("上层桥面系", "桥面系", None), ("下层桥面系", "桥面系", None)]
r = 4
first_data = r
part_rows = {}
DET_T = {
    "主桁": {int(d["t"]) for d in D["main"] + D["hb"]},
    "主桁拼接": {int(d["t"]) for d in D["spl"]},
    "上层桥面系": {int(d["t"]) for d in DECK if d["seg"] == "上层桥面系"},
    "下层桥面系": {int(d["t"]) for d in DECK if d["seg"] == "下层桥面系"},
}
for part, s1, s2 in parts:
    start = r
    ths = set(DET_T[part])
    for m in MATS:
        ths |= {int(t) for t in DWG.get((part, m), {})}
    for t in sorted(ths):
        ws.cell(row=r, column=1, value=part)
        ws.cell(row=r, column=2, value=t)
        ws.cell(row=r, column=3, value="kg")
        d355 = DWG.get((part, "Q355D"), {}).get(t, 0) or None
        d420 = DWG.get((part, "Q420D"), {}).get(t, 0) or None
        ws.cell(row=r, column=4, value=d355)
        ws.cell(row=r, column=5, value=d420)
        ws.cell(row=r, column=6, value=f'=N(D{r})+N(E{r})')
        r += 1
    part_rows[part] = (start, r - 1)
    ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
    ws.cell(row=start, column=1).alignment = AL

# 复核数量公式
for part, (s, e) in part_rows.items():
    for row in range(s, e + 1):
        t = ws.cell(row=row, column=2).value
        formula = None
        if part == "主桁":
            f1 = f"SUMIFS('主梁数量'!$J$4:$J${main_last},'主梁数量'!$D$4:$D${main_last},D$3,'主梁数量'!$G$4:$G${main_last},$B{row})"
            f2 = f"SUMIFS('主梁数量'!$J$4:$J${main_last},'主梁数量'!$D$4:$D${main_last},E$3,'主梁数量'!$G$4:$G${main_last},$B{row})"
            f3 = f"SUMIFS('横隔板'!$J$4:$J${hb_last},'横隔板'!$D$4:$D${hb_last},D$3,'横隔板'!$G$4:$G${hb_last},$B{row})"
            f4 = f"SUMIFS('横隔板'!$J$4:$J${hb_last},'横隔板'!$D$4:$D${hb_last},E$3,'横隔板'!$G$4:$G${hb_last},$B{row})"
            formula = f"=SUM({f1},{f2},{f3},{f4})"
        elif part == "主桁拼接":
            f1 = f"SUMIFS('主桁拼接'!$J$4:$J${sp_last},'主桁拼接'!$D$4:$D${sp_last},D$3,'主桁拼接'!$G$4:$G${sp_last},$B{row})"
            f2 = f"SUMIFS('主桁拼接'!$J$4:$J${sp_last},'主桁拼接'!$D$4:$D${sp_last},E$3,'主桁拼接'!$G$4:$G${sp_last},$B{row})"
            formula = f"=SUM({f1},{f2})"
        else:
            formula = (f"=SUMIFS('桥面系'!$J$4:$J${dk_last},'桥面系'!$A$4:$A${dk_last},\"{part}\","
                       f"'桥面系'!$G$4:$G${dk_last},$B{row})")
        ws.cell(row=row, column=7, value=formula)
        ws.cell(row=row, column=8, value=f"=G{row}-F{row}")

# 合计行
tot = r
ws.cell(row=tot, column=1, value="钢结构合计")
ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=3)
ws.cell(row=tot, column=4, value=f"=SUM(D{first_data}:D{tot - 1})")
ws.cell(row=tot, column=5, value=f"=SUM(E{first_data}:E{tot - 1})")
ws.cell(row=tot, column=6, value=f"=SUM(F{first_data}:F{tot - 1})")
ws.cell(row=tot, column=7, value=f"=SUM(G{first_data}:G{tot - 1})")
ws.cell(row=tot, column=8, value=f"=G{tot}-F{tot}")
ws.cell(row=tot, column=9, value="含主桁、主桁拼接、上/下层桥面系；不含焊缝、高强螺栓、涂装、检查车")

# 焊缝行（按图纸 1.5% 计）
wf = tot + 1
ws.cell(row=wf, column=1, value="焊缝重(1.5%)")
ws.merge_cells(start_row=wf, start_column=1, end_row=wf, end_column=3)
ws.cell(row=wf, column=6, value=16927.6 + 38341.2 + 30002.9 + 54464.7)
ws.cell(row=wf, column=7, value=f"=ROUND(G{tot}*0.015,1)")
ws.cell(row=wf, column=8, value=f"=G{wf}-F{wf}")
ws.cell(row=wf, column=9, value="图纸焊缝：主桁 16927.6+38341.2，拼接 30002.9，下层桥面系 54464.7")

tr = wf + 1
ws.cell(row=tr, column=1, value="总合计")
ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=3)
ws.cell(row=tr, column=6, value=f"=F{tot}+F{wf}")
ws.cell(row=tr, column=7, value=f"=G{tot}+G{wf}")
ws.cell(row=tr, column=8, value=f"=G{tr}-F{tr}")

style_range(ws, 4, tr, 9)
for row in range(4, tr + 1):
    for c in (4, 5, 6, 7, 8):
        ws.cell(row=row, column=c).number_format = "#,##0.0_ "
for c in range(1, 10):
    ws.cell(row=tot, column=c).font = F_HEAD
    ws.cell(row=tr, column=c).font = F_HEAD
    ws.cell(row=wf, column=c).font = F_BODY
ws.cell(row=tr, column=1).font = F_HEAD
ws.cell(row=tr, column=1).alignment = AL

widths = {"A": 14, "B": 8, "C": 6, "D": 14, "E": 14, "F": 14, "G": 14, "H": 11, "I": 46}
for k, v in widths.items():
    ws.column_dimensions[k].width = v
ws.freeze_panes = "A4"
ws.page_setup.orientation = "portrait"
ws.page_setup.paperSize = 9

# 说明
note_r = tr + 2
notes = [
    "注：",
    "1、本底稿依据《K6+418.3 新沙高架桥 主桥上构工程数量表》（图号 SZYS/04/02/02/01/01）及各节段设计图（图号 09～52）逐板复算。",
    "2、主桁含中桁 22 个节段、边桁 22 个节段；横隔板单列（N15、N25、N6x、N7x、N8x 系列），同属主桁部位，汇总时与主梁合并对图纸“主桁”列。",
    "3、单件重按 Q7850kg/m³ 计算：箱形板件 长×宽×厚×7850；面积法板件（t=××mm,A=××m²）按 厚×面积×7850。",
    "4、边桁数量按图纸单侧统计，未乘 2；差异栏为正表示复核大于图纸，负表示复核小于图纸。",
    "5、焊缝重、10.9S 高强螺栓（M30×200、M24）、剪力钉、涂装、检查车、支座等未计入板件重量，另列。",
]
for i, t in enumerate(notes):
    ws.cell(row=note_r + i, column=1, value=t).font = Font(name="宋体", size=9)
    ws.cell(row=note_r + i, column=1).alignment = Alignment(horizontal="left", vertical="center")

# ---------- 梁底挡块（本桥为钢桁梁，无梁底挡块，保留模板表式） ----------
ws = wb.create_sheet("梁底挡块")
ws.merge_cells("A1:K1")
ws.cell(row=1, column=1, value="梁底挡块数量表").font = Font(name="宋体", size=12, bold=True)
ws.cell(row=1, column=1).alignment = AL
heads = ["梁板编号", "板件名称", "位置", "材质", "长(mm)", "宽(mm)", "厚(mm)", "数量", "单件重(kg)", "全桥数量", "合计重量(kg)"]
for i, h in enumerate(heads, 1):
    ws.cell(row=2, column=i, value=h)
style_header(ws, 2, 11)
ws.cell(row=3, column=1, value="本桥主桥上构为钢桁梁，主桁节点采用整体节点板连接，不设混凝土梁底挡块，本表无内容。")
ws.merge_cells("A3:K3")
ws.cell(row=3, column=1).alignment = AL_L
ws.cell(row=3, column=1).font = F_BODY
ws.cell(row=3, column=1).border = BORDER
for i, w in enumerate([13, 15, 12, 8, 9, 9, 8, 8, 11, 9, 12], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---------- 附属预埋件 ----------
ws = wb.create_sheet("附属预埋件")
ws.merge_cells("A1:K1")
ws.cell(row=1, column=1, value="附属工程预埋件数量表").font = Font(name="宋体", size=12, bold=True)
ws.cell(row=1, column=1).alignment = AL
for i, h in enumerate(heads, 1):
    ws.cell(row=2, column=i, value=h)
style_header(ws, 2, 11)
aux = [
    ("排水", "泄水管及管卡", "下层桥面系", "Q355C", 780.6, "套", "图二 附属工程-集中排水 φ150 泄水管；钢构件预埋重", 780.6),
    ("护栏", "防撞护栏预埋钢板", "上层桥面系", "Q355C", 632.0, "kg", "图二 附属工程-上层防撞护栏 M16×50 螺栓配套预埋板", 632.0),
    ("护栏", "护栏底座预埋", "上层桥面系", "Q355C", 782.3, "kg", "图二 附属工程-上层防撞护栏 M16×170", 782.3),
    ("护栏", "人行道栏杆预埋", "上层桥面系", "Q355C", 607.9, "kg", "图二 附属工程-上层防撞护栏 M16×55", 607.9),
    ("检查车", "下弦检查车轨道预埋", "主桁下弦", "Q355B", 65513.5, "kg", "图二 附属工程-检修轨道及检查车 Q355B", 65513.5),
    ("伸缩缝", "伸缩缝预埋钢筋", "上/下层桥面系", "HRB400", 95.1, "kg", "图二 附属工程-伸缩缝 D240/D320 锚固钢筋", 95.1),
]
r = 3
for a in aux:
    ws.cell(row=r, column=1, value=a[0])
    ws.cell(row=r, column=2, value=a[1])
    ws.cell(row=r, column=3, value=a[2])
    ws.cell(row=r, column=4, value=a[3])
    ws.cell(row=r, column=5, value=None)
    ws.cell(row=r, column=6, value=None)
    ws.cell(row=r, column=7, value=None)
    ws.cell(row=r, column=8, value=1)
    ws.cell(row=r, column=9, value=a[4])
    ws.cell(row=r, column=10, value=1)
    ws.cell(row=r, column=11, value=f"=I{r}*J{r}")
    r += 1
style_range(ws, 3, r - 1, 11)
for rr in range(3, r):
    ws.cell(row=rr, column=11).number_format = "0.0_ "
    ws.cell(row=rr, column=9).number_format = "0.0_ "
ws.cell(row=r, column=1, value="合计")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
ws.cell(row=r, column=11, value=f"=SUM(K3:K{r - 1})")
for c in range(1, 12):
    ws.cell(row=r, column=c).font = F_HEAD
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).alignment = AL
ws.cell(row=r, column=11).number_format = "0.0_ "
for i, w in enumerate([12, 20, 14, 9, 9, 9, 8, 8, 12, 9, 12], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
note_r = r + 2
ws.cell(row=note_r, column=1, value="注：附属工程钢构件按《主桥上构工程数量表（二）》附属工程列录入，仅为预埋/附属钢件，不含混凝土、支座本体及涂装。").font = F_BODY

wb.save("新沙高架桥主桥上构钢结构计算底稿.xlsx")
print("saved")
for s in wb.sheetnames:
    print(s, wb[s].max_row, wb[s].max_column)
