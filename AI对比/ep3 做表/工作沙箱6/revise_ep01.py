# -*- coding: utf-8 -*-
"""修订首轮节段01底稿：保留计算，重算并写入缓存值；撤回结论；补量沿用公式。"""
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

SRC = "新沙高架桥 节段01 钢结构重量计算底稿.xlsx"
OUT = "新沙高架桥 节段01 钢结构重量计算底稿.xlsx"

FN = "宋体"
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_B = Font(name=FN, size=9)
F_H = Font(name=FN, size=11, bold=True)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LAL = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = openpyxl.load_workbook(SRC)

# ---- 1. N16c 补量：数量 0 也保留公式结果（重量为 0），并在备注标注 ----
ws = wb["节段01计算底稿"]
for r in range(4, 72):
    if ws.cell(row=r, column=3).value == "N16c":
        ws.cell(row=r, column=19,
                value="图纸数量为0；补量后按本行公式自动计重（290.5kg/件 × 数量）")
        print("N16c 备注已更新，行", r)

# ---- 2. 撤回结论：重写说明区 ----
# 定位“差异说明”标题行
note_row = None
for r in range(72, ws.max_row + 2):
    v = ws.cell(row=r, column=1).value
    if v == "差异说明":
        note_row = r
        break
print("说明区起始行", note_row)

# 清空原说明区（标题下所有合并行）
if note_row:
    merged = [rng for rng in list(ws.merged_cells.ranges) if rng.min_row >= note_row]
    for rng in merged:
        ws.unmerge_cells(str(rng))
    for r in range(note_row + 1, ws.max_row + 1):
        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()
    start = note_row + 1
else:
    start = ws.max_row + 2
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=13)
    c = ws.cell(row=note_row, column=1, value="差异说明")
    c.font = F_H
    c.alignment = Alignment(horizontal="left", vertical="center")
    start = note_row + 1

NOTES = [
 ("一、计算口径（未改动）", True),
 ("1. 板材：W = 长(m) × 宽(m) × 板厚(m) × 7850kg/m³；图纸直接给面积的零件按 W = 面积 × 板厚 × 7.85kg/m²。全表公式保留，改动尺寸或数量即自动重算。", False),
 ("2. 图纸数据（规格、数量、单件重）原样照录；计算重量与图纸重量分列，差异 = 计算 − 图纸。焊缝 1.5% 单列，不摊入单件重。", False),
 ("", False),
 ("二、与图纸的差异（只呈现事实）", True),
 ("3. 材料重：计算 121,958.9kg，图纸 121,975.6kg，差 −16.7kg（−0.014%）；含 1.5% 焊缝后 123,788.3kg，图纸 123,805.2kg，差 −16.9kg。", False),
 ("4. 逐行差异见上表“差异”与“差异率”两列，单件偏差绝对值最大 1.23kg（N31/N31a），其余均在 1.1kg 以内。", False),
 ("5. 差异成因：本底稿不再给出统一结论。此前“全部为取整差异”的说法已撤回——未逐项核对图纸单件重的取整规则、面积取值精度、是否含焊缝，不足以支持该判断。各行差异原因列为本表待核项，需与设计方逐项核对后确认。", False),
 ("", False),
 ("三、结论边界（原“可用于计量”已撤回）", True),
 ("6. 本表为复核计算文件，不作为计量结算依据。能否用于计量，取决于：①差异成因逐项确认；②焊缝、涂装、栓焊件的计量口径；③合同约定的计量规则。三项均未完成，故不作可用性结论。", False),
 ("7. 差异率极小的行（如 N11）不代表该行计算无误，只说明该行为整数规格件；判断精度应看单件绝对偏差，不看差异率。", False),
 ("", False),
 ("四、待核项（未取得依据，不臆断）", True),
 ("8. 差异成因：68 行逐行差异的具体来源（取整位数、面积舍入、是否含焊缝余量）均待核，未逐行定性。", False),
 ("9. N24 下弦节点板图纸标“Z35钢”：重量按 Q355D 计不受影响，但 Z 向性能要求、是否影响订货单价，待设计与采购确认。", False),
 ("10. N15b 为表中唯一 Q355B（10kg）：材质分档是否需单列材料，待确认。", False),
 ("11. N16c 图纸数量为 0：补量条件、补量后是否计入本节段，待设计确认；本表已按公式预留，数量填入即出重量。", False),
 ("12. 图纸第 49 页右表（横梁接头 14 行、支承加劲肋 5 行）与左表并排列示，本表已按左 54 行 + 右 14 行合并计入；两表是否同属“节段01”口径，待图纸说明确认。", False),
 ("13. 焊缝 1.5% 的计取基数：图纸按材料重 1.5% 计，是否含全部焊缝类型、是否需按板厚分档，待确认。", False),
 ("", False),
 ("五、校验记录", True),
 ("14. 行级：68 行全部有计算值，无空值；数量×单件重与图纸共重不符的 0 行（阈值 0.6%）。", False),
 ("15. 表级：明细合计 121,958.9kg，与第49页材料表合计 121,975.6kg 对应（左表 110,447.8kg + 右表 11,527.8kg），误差 −16.7kg 见第 3 条。", False),
 ("16. 复核强度：单遍计算 + 规则校验，未做第二人独立复核；对外使用前建议抽样对照原图。", False),
]

r = start
for txt, bold in NOTES:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    c = ws.cell(row=r, column=1, value=txt)
    c.font = F_H if bold else F_B
    c.alignment = LAL
    ws.row_dimensions[r].height = 22 if not txt else (26 if len(txt) < 60 else 40)
    r += 1

wb.save(OUT)
print("saved", OUT)
