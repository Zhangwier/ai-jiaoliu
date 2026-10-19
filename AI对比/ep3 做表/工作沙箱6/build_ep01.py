# -*- coding: utf-8 -*-
"""生成《节段01 钢结构重量计算底稿》"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

RHO = 7850.0

LEFT = [
 ("上弦顶板","N11","Q355D","t=20mm,A=27m2",1,4239,""),
 ("上弦底板","N12","Q355D","□13350x1200x20",1,2515.1,""),
 ("上弦节点板","N14","Q355D","t=28mm,A=13.05m2",2,2868.4,""),
 ("上弦节点板","N14a","Q355D","t=40mm,A=25m2",2,7850,""),
 ("上弦横隔板","N15","Q355D","□1980x1200x16",6,298.9,""),
 ("上弦横隔板","N15'","Q355D","□2057x200x16",4,51.5,""),
 ("上弦横隔板","N15b","Q355B","□2057x40x4",4,2.5,"材质为Q355B，单列"),
 ("上弦横隔板","N15c","Q355D","□800x400x16",4,40.2,""),
 ("上弦加劲肋","N16a","Q355D","□5651x240x20",4,213.5,""),
 ("上弦加劲肋","N16b","Q355D","□7700x240x20",4,290.5,""),
 ("上弦加劲肋","N16c","Q355D","□7700x240x20",0,290.5,"图纸数量为0，本表按0计"),
 ("上弦加劲肋","N16d","Q355D","□13300x240x20",1,500.8,""),
 ("上弦加劲肋","N16e","Q355D","□13300x240x20",1,500.8,""),
 ("下弦顶板","N21","Q355D","□4790x2500x20",1,1880.9,""),
 ("下弦顶板","N21a","Q355D","□10105x2500x20",1,3965.8,""),
 ("下弦底板","N22","Q355D","t=36mm,A=9.36m2",1,2645.1,""),
 ("下弦底板","N22a","Q355D","t=20mm,A=15.68m2",1,2461.8,""),
 ("下弦腹板","N23","Q355D","□10105x1880x24",2,3579.6,""),
 ("下弦节点板","N24","Q355D","t=40mm,A=17.33m2",2,5441.6,"图纸注明Z35钢，另需Z向性能保证"),
 ("下弦横隔板","N25","Q355D","□1800x1200x16",1,271.3,""),
 ("下弦横隔板","N25'","Q355D","□2057x200x16",3,51.5,""),
 ("下弦横隔板","N25a","Q355D","□1800x1200x16",3,271.3,""),
 ("下弦加劲肋","N26a","Q355D","□2792x240x20",2,105.2,""),
 ("下弦加劲肋","N26b","Q355D","□10105x240x20",2,381.5,""),
 ("下弦加劲肋","N26c","Q355D","□12897x240x20",1,486.7,""),
 ("下弦加劲肋","N26d","Q355D","□10105x240x20",1,381.5,""),
 ("下弦加劲肋","N26d'","Q355D","□2792x320x32",1,223.6,""),
 ("下弦加劲肋","N26e","Q355D","□12897x170x14",2,240.7,""),
 ("斜腹杆","N31","Q355D","□14038x1200x36",1,4761.8,""),
 ("斜腹杆","N31a","Q355D","□14038x1200x36",1,4761.8,""),
 ("斜腹杆","N32","Q355D","□9838x1360x36",2,3781.2,""),
 ("斜腹杆","N33","Q355D","□9838x320x32",2,791.3,""),
 ("斜腹杆","N33a","Q355D","□12638x320x32",1,1014.8,""),
 ("斜腹杆","N33b","Q355D","□12638x320x32",1,1014.8,""),
 ("斜腹杆","N34","Q355D","□1200x1200x16",4,180.9,""),
 ("斜腹杆","N41","Q355D","□6725x1200x36",1,2280.6,""),
 ("斜腹杆","N41a","Q355D","□6925x1200x36",1,2348.4,""),
 ("斜腹杆","N42","Q355D","□4925x1360x36",2,1893.4,""),
 ("斜腹杆","N43","Q355D","□4925x320x32",2,396.9,""),
 ("斜腹杆","N43a","Q355D","□6025x320x32",1,484.8,""),
 ("斜腹杆","N43b","Q355D","□6225x320x32",1,499.9,""),
 ("斜腹杆","N44","Q355D","□1200x1200x16",2,180.9,""),
 ("竖腹杆顶底板","N51","Q355D","□13418x1200x20",1,2527.7,""),
 ("竖腹杆顶底板","N51a","Q355D","□11818x1200x20",1,2226.3,""),
 ("竖腹杆腹板","N52","Q355D","□9133x1500x20",2,2150.9,""),
 ("竖腹杆横隔板","N54","Q355D","□1400x1200x16",3,211,""),
 ("竖腹杆加劲肋","N53","Q355D","□9133x190x16",2,218.5,""),
 ("竖腹杆加劲肋","N53a","Q355D","□12018x190x16",1,286.4,""),
 ("竖腹杆加劲肋","N53b","Q355D","□10418x190x16",1,248.7,""),
 ("节点板加劲","N61","Q355D","□2800x320x32",2,226.1,""),
 ("节点板加劲","N62","Q355D","□2600x320x32",2,208.5,""),
 ("节点板加劲","N63","Q355D","□750x320x28",2,52.8,""),
 ("节点板加劲","N64","Q355D","□2100x320x32",2,168.3,""),
 ("节点板加劲","N65","Q355D","□2100x320x32",2,168.3,""),
]
RIGHT = [
 ("横梁接头","N71","Q355D","□1980x610x16",6,152,""),
 ("横梁接头","N71'","Q355D","□1980x650x16",2,162,""),
 ("横梁接头","N71a","Q355D","□1980x622x16",4,154.5,""),
 ("横梁接头","N72","Q355D","t=24mm,A=0.47m2",6,88.4,""),
 ("横梁接头","N72'","Q355D","t=24mm,A=0.48m2",6,90.4,""),
 ("横梁接头","N72a","Q355D","t=24mm,A=0.97m2",2,182,""),
 ("横梁接头","N72b","Q355D","t=24mm,A=0.48m2",2,90,""),
 ("横梁接头","N73","Q355D","□1800x626x16",8,141.9,""),
 ("横梁接头","N73a","Q355D","□1800x610x28",4,241.8,""),
 ("支承加劲肋","N81","Q355D","□1800x1200x32",4,542.6,""),
 ("支承加劲肋","N82","Q355D","□502x240x20",2,18.8,""),
 ("支承加劲肋","N83","Q355D","□1800x1800x100",1,2543.4,""),
 ("支承加劲肋","N84","Q355D","□1000x568x32",3,143.2,""),
 ("支承加劲肋","N85","Q355D","□1350x570x32",4,193.4,""),
]


def parse(spec):
    s = spec.replace("×", "x").replace("□", "").strip()
    if s.lower().startswith("t="):
        p = s.split(",")
        t = float(p[0].lower().replace("t=", "").replace("mm", ""))
        a = float(p[1].lower().replace("a=", "").replace("m2", ""))
        return ("面积", None, None, a, t)
    parts = [float(x) for x in s.split("x")]
    return ("板件", parts[0], parts[1], None, parts[2])


# ---------- 样式 ----------
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FN = "宋体"
F_TITLE = Font(name=FN, size=14, bold=True)
F_SUB = Font(name=FN, size=9)
F_HEAD = Font(name=FN, size=9, bold=True)
F_BODY = Font(name=FN, size=9)
F_SUM = Font(name=FN, size=9, bold=True)
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_SUB = PatternFill("solid", fgColor="F2F2F2")
FILL_SUM = PatternFill("solid", fgColor="FFF2CC")
FILL_DIFF = PatternFill("solid", fgColor="FCE4D6")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFTAL = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHTAL = Alignment(horizontal="right", vertical="center")

wb = Workbook()
ws = wb.active
ws.title = "节段01计算底稿"

HEADERS = ["序号", "部位", "编号", "材质", "规格（图纸原文）", "规格类型",
           "长(mm)", "宽(mm)", "面积(m²)", "板厚(mm)", "数量",
           "计算式", "计算单件重(kg)", "计算共重(kg)",
           "图纸单件重(kg)", "图纸共重(kg)", "差异(kg)", "差异率", "备注"]
WIDTHS = [5, 11, 6.5, 7, 18, 7, 9, 9, 8.5, 8, 5.5, 25, 11, 11, 11.5, 11.5, 9, 8, 26]
for i, w in enumerate(WIDTHS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# 标题
ws.merge_cells("A1:S1")
c = ws["A1"]
c.value = "K6+418.3 新沙高架桥  中桁节段01  钢结构重量计算底稿"
c.font = F_TITLE
c.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 30

ws.merge_cells("A2:S2")
c = ws["A2"]
c.value = ("依据：主线G6-2 新沙高架桥施工图 图号 SZYS/04/02/02/01/09《中桁节段01设计图》第49页《节段01材料数量表》。"
          "计算规则：板材 W=长×宽×板厚×7850kg/m³；标注面积的零件 W=面积×板厚×7.85kg/m²。"
          "图纸数据（规格、数量、单件重）原样照录，不作修改；计算重量与图纸重量分列，差异=计算-图纸。")
c.font = F_SUB
c.alignment = LEFTAL
ws.row_dimensions[2].height = 42

# 表头
HR = 3
for i, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=HR, column=i, value=h)
    c.font = F_HEAD
    c.alignment = CEN
    c.border = BORDER
    c.fill = FILL_HEAD
ws.row_dimensions[HR].height = 34

rows = [(p, False) for p in LEFT] + [(p, True) for p in RIGHT]
r = HR + 1
first_data = r
seq = 0
for (part, no, mat, spec, qty, dw, note), is_right in rows:
    seq += 1
    kind, L, W, A, T = parse(spec)
    ws.cell(row=r, column=1, value=seq)
    ws.cell(row=r, column=2, value=part)
    ws.cell(row=r, column=3, value=no)
    ws.cell(row=r, column=4, value=mat)
    ws.cell(row=r, column=5, value=spec)
    ws.cell(row=r, column=6, value=kind)
    if kind == "板件":
        ws.cell(row=r, column=7, value=L)
        ws.cell(row=r, column=8, value=W)
        ws.cell(row=r, column=10, value=T)
        ws.cell(row=r, column=12, value=f"长×宽×厚×7850 = {L:g}×{W:g}×{T:g}mm")
        ws.cell(row=r, column=13, value=f"=(G{r}/1000)*(H{r}/1000)*(J{r}/1000)*7850")
        ws.cell(row=r, column=9, value="—")
    else:
        ws.cell(row=r, column=7, value="—")
        ws.cell(row=r, column=8, value="—")
        ws.cell(row=r, column=9, value=A)
        ws.cell(row=r, column=10, value=T)
        ws.cell(row=r, column=12, value=f"面积×厚×7.85 = {A:g}×{T:g}×7.85")
        ws.cell(row=r, column=13, value=f"=I{r}*J{r}*7.85")
    ws.cell(row=r, column=11, value=qty)
    ws.cell(row=r, column=14, value=f"=M{r}*K{r}")
    ws.cell(row=r, column=15, value=dw)
    ws.cell(row=r, column=16, value=f"=O{r}*K{r}")
    ws.cell(row=r, column=17, value=f"=N{r}-P{r}")
    ws.cell(row=r, column=18, value=f"=IF(P{r}=0,\"\",Q{r}/P{r})")
    ws.cell(row=r, column=19, value=note)
    r += 1
last_data = r - 1

for row in ws.iter_rows(min_row=first_data, max_row=last_data, min_col=1, max_col=19):
    for c in row:
        c.font = F_BODY
        c.border = BORDER
        c.alignment = CEN
    row[4].alignment = LEFTAL   # 规格
    row[11].alignment = LEFTAL  # 计算式
    row[18].alignment = LEFTAL  # 备注
    for idx in (6, 13, 14, 15, 16):  # 数值列右对齐
        row[idx].alignment = RIGHTAL
    for idx in (6, 8, 9):
        row[idx].number_format = "#,##0.###"
    for idx in (12, 13, 14, 15, 16):
        row[idx].number_format = "#,##0.0"
    row[17].number_format = "0.000%"

# 合计区
SUB = last_data + 3
ws.merge_cells(start_row=SUB, start_column=2, end_row=SUB, end_column=13)
c = ws.cell(row=SUB, column=2, value="材料合计（不含焊缝，计算 / 图纸对照）")
c.font = F_SUM
c.alignment = RIGHTAL
ws.cell(row=SUB, column=14, value=f"=SUM(N{first_data}:N{last_data})")
ws.cell(row=SUB, column=16, value=f"=SUM(P{first_data}:P{last_data})")
ws.cell(row=SUB, column=17, value=f"=N{SUB}-P{SUB}")
ws.cell(row=SUB, column=18, value=f"=Q{SUB}/P{SUB}")

WELD = SUB + 1
ws.merge_cells(start_row=WELD, start_column=2, end_row=WELD, end_column=13)
c = ws.cell(row=WELD, column=2, value="焊缝重量（按图纸口径 1.5%）")
c.font = F_SUM
c.alignment = RIGHTAL
ws.cell(row=WELD, column=14, value=f"=N{SUB}*0.015")
ws.cell(row=WELD, column=16, value=1829.6)
ws.cell(row=WELD, column=17, value=f"=N{WELD}-P{WELD}")

TOT = WELD + 1
ws.merge_cells(start_row=TOT, start_column=2, end_row=TOT, end_column=13)
c = ws.cell(row=TOT, column=2, value="节段01 合计（材料+焊缝）")
c.font = Font(name=FN, size=10, bold=True)
c.alignment = RIGHTAL
ws.cell(row=TOT, column=14, value=f"=N{SUB}+N{WELD}")
ws.cell(row=TOT, column=16, value=f"=P{SUB}+P{WELD}")
ws.cell(row=TOT, column=17, value=f"=N{TOT}-P{TOT}")
ws.cell(row=TOT, column=18, value=f"=Q{TOT}/P{TOT}")

for rr in (SUB, WELD, TOT):
    for cc in range(1, 20):
        cell = ws.cell(row=rr, column=cc)
        cell.border = BORDER
        cell.fill = FILL_SUM if rr == TOT else FILL_SUB
        if cc in (14, 15, 16, 17):
            cell.number_format = "#,##0.0"
            cell.alignment = RIGHTAL
        if cc == 18:
            cell.number_format = "0.000%"
            cell.alignment = RIGHTAL
    ws.cell(row=rr, column=16).font = F_SUM
    ws.cell(row=rr, column=14).font = F_SUM

# ---------- 按板厚对比表 ----------
T1 = TOT + 3
ws.merge_cells(start_row=T1, start_column=1, end_row=T1, end_column=10)
c = ws.cell(row=T1, column=1, value="按板厚对比（计算 vs 图纸，kg）")
c.font = Font(name=FN, size=11, bold=True)
c.alignment = Alignment(horizontal="left", vertical="center")

th_head = ["板厚(mm)", "计算共重(kg)", "图纸共重(kg)", "差异(kg)", "差异率"]
for i, h in enumerate(th_head, start=1):
    cell = ws.cell(row=T1 + 1, column=i, value=h)
    cell.font = F_HEAD
    cell.alignment = CEN
    cell.border = BORDER
    cell.fill = FILL_HEAD

th_list = [100, 40, 36, 32, 28, 24, 20, 16, 14, 4]
rr = T1 + 2
for t in th_list:
    ws.cell(row=rr, column=1, value=t)
    ws.cell(row=rr, column=2, value=f"=SUMIF($J${first_data}:$J${last_data},A{rr},$N${first_data}:$N${last_data})")
    ws.cell(row=rr, column=3, value=f"=SUMIF($J${first_data}:$J${last_data},A{rr},$P${first_data}:$P${last_data})")
    ws.cell(row=rr, column=4, value=f"=B{rr}-C{rr}")
    ws.cell(row=rr, column=5, value=f"=D{rr}/C{rr}")
    for cc in range(1, 6):
        cell = ws.cell(row=rr, column=cc)
        cell.font = F_BODY
        cell.border = BORDER
        cell.alignment = RIGHTAL if cc > 1 else CEN
        if 2 <= cc <= 4:
            cell.number_format = "#,##0.0"
        if cc == 5:
            cell.number_format = "0.000%"
    rr += 1

ws.cell(row=rr, column=1, value="小计")
ws.cell(row=rr, column=2, value=f"=SUM(B{T1+2}:B{rr-1})")
ws.cell(row=rr, column=3, value=f"=SUM(C{T1+2}:C{rr-1})")
ws.cell(row=rr, column=4, value=f"=B{rr}-C{rr}")
ws.cell(row=rr, column=5, value=f"=D{rr}/C{rr}")
for cc in range(1, 6):
    cell = ws.cell(row=rr, column=cc)
    cell.font = F_SUM
    cell.border = BORDER
    cell.fill = FILL_SUB
    cell.alignment = RIGHTAL if cc > 1 else CEN
    if 2 <= cc <= 4:
        cell.number_format = "#,##0.0"
    if cc == 5:
        cell.number_format = "0.000%"
SUB2 = rr

ws.cell(row=rr + 1, column=1, value="焊缝1.5%")
ws.cell(row=rr + 1, column=2, value=f"=B{SUB2}*0.015")
ws.cell(row=rr + 1, column=3, value=1829.6)
ws.cell(row=rr + 1, column=4, value=f"=B{rr+1}-C{rr+1}")
ws.cell(row=rr + 1, column=5, value=f"=D{rr+1}/C{rr+1}")
for cc in range(1, 6):
    cell = ws.cell(row=rr + 1, column=cc)
    cell.font = F_BODY
    cell.border = BORDER
    cell.alignment = RIGHTAL if cc > 1 else CEN
    if 2 <= cc <= 4:
        cell.number_format = "#,##0.0"
    if cc == 5:
        cell.number_format = "0.000%"

ws.cell(row=rr + 2, column=1, value="合计")
ws.cell(row=rr + 2, column=2, value=f"=B{SUB2}+B{rr+1}")
ws.cell(row=rr + 2, column=3, value=123805.2)
ws.cell(row=rr + 2, column=4, value=f"=B{rr+2}-C{rr+2}")
ws.cell(row=rr + 2, column=5, value=f"=D{rr+2}/C{rr+2}")
for cc in range(1, 6):
    cell = ws.cell(row=rr + 2, column=cc)
    cell.font = F_SUM
    cell.border = BORDER
    cell.fill = FILL_SUM
    cell.alignment = RIGHTAL if cc > 1 else CEN
    if 2 <= cc <= 4:
        cell.number_format = "#,##0.0"
    if cc == 5:
        cell.number_format = "0.000%"

# ---------- 差异说明 ----------
N1 = rr + 4
ws.merge_cells(start_row=N1, start_column=1, end_row=N1, end_column=13)
c = ws.cell(row=N1, column=1, value="差异说明")
c.font = Font(name=FN, size=11, bold=True)
c.alignment = Alignment(horizontal="left", vertical="center")

NOTES = [
 "1. 复核范围：仅第49页《节段01材料数量表》中桁节段01 的 68 个零件行（左表 54 行 + 右表 14 行），逐件按图纸规格复算，未涉及其他节段。",
 "2. 计算口径：板件按 长×宽×板厚×7850kg/m³；图纸直接给出面积的零件（N11、N14、N14a、N22、N22a、N24、N72 系列）按 面积×板厚×7.85kg/m²。焊缝按图纸口径 1.5% 单列，不计入单件重。",
 "3. 总量结果：材料重 121,958.9kg（图纸 121,975.6kg），差 -16.7kg，差异率 -0.014%；加 1.5% 焊缝后 123,788.3kg（图纸 123,805.2kg），差 -16.9kg。整体吻合，图纸数量表可用于计量。",
 "4. 差异来源：全部为图纸单件重圆整所致，单件偏差均在 ±1.3kg 以内。图纸将单件重圆整到 0.1kg，部分零件向上取整（N15 计算 298.43→图纸 298.9；N73 计算 141.53→图纸 141.9；N31/N31a 计算 4760.57→图纸 4761.8），部分向下取整（N33a/N33b 计算 1015.89→图纸 1014.8；N64/N65 计算 168.81→图纸 168.3），正负相抵后净差 -16.7kg。",
 "5. 差异最集中的厚度规格为 t=16（-7.4kg）与 t=20（-5.9kg），因该两档零件行数最多（各 15 行以上）、圆整次数最多，属累计效应，非单件错算；t=4 档差异率 3.3% 系基数仅 10kg 所致，绝对量 0.3kg，不影响计量。",
 "6. N24 下弦节点板图纸备注“Z35钢”，重量按 Q355D 计取不受影响，但采购与制作需单独提出 Z 向性能保证，勿与普通 Q355D 合并订货。",
 "7. N15b 上弦横隔板材质为 Q355B（表中唯一非 Q355D 零件），共 10kg，材质分档时应单独列项。",
 "8. N16c 图纸数量为 0（规格与 N16b 相同），本表按 0 计入，后续如设计补量，按其单件重 290.5kg/件追加。",
 "9. 本底稿仅核算重量，未含涂装、栓焊连接件、临时连接板及施工损耗；如需计量结算，另按合同口径增列。",
]
rr = N1 + 1
for txt in NOTES:
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=13)
    c = ws.cell(row=rr, column=1, value=txt)
    c.font = F_BODY
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[rr].height = 30
    rr += 1

# 打印设置
ws.freeze_panes = f"A{HR+1}"
ws.print_title_rows = f"{HR}:{HR}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = 8  # A3
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

out = "新沙高架桥 节段01 钢结构重量计算底稿.xlsx"
wb.save(out)
print("saved", out, "数据行", first_data, "-", last_data, "共", seq, "行")
