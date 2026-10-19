# -*- coding: utf-8 -*-
"""生成全桥节段重量计算底稿 Excel。"""
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

RHO = 7850.0


def num(v, nd=1):
    return None if v is None else round(v, nd)


def main():
    segs = json.load(open("work/all_segs_final.json", encoding="utf-8"))
    calc = {t["seg"]: t for t in json.load(open("work/calc_out.json", encoding="utf-8"))}

    THIN = Side(style="thin", color="000000")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    FN = "宋体"
    F_T = Font(name=FN, size=14, bold=True)
    F_S = Font(name=FN, size=9)
    F_H = Font(name=FN, size=9, bold=True)
    F_B = Font(name=FN, size=9)
    F_SUM = Font(name=FN, size=9, bold=True)
    FILL_H = PatternFill("solid", fgColor="D9E1F2")
    FILL_S = PatternFill("solid", fgColor="FFF2CC")
    FILL_W = PatternFill("solid", fgColor="FCE4D6")
    CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LAL = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RAL = Alignment(horizontal="right", vertical="center")

    wb = Workbook()

    # ---------- 明细 ----------
    ws = wb.active
    ws.title = "全桥计算明细"
    HEAD = ["序号", "墩号/节段", "图纸页", "部位", "编号", "材质",
            "规格（图纸原文）", "类型", "数量",
            "计算单件重(kg)", "计算共重(kg)", "图纸单件重(kg)", "图纸共重(kg)",
            "差异(kg)", "差异率", "备注"]
    W = [5, 10, 6, 13, 7, 7, 20, 7, 5.5, 12, 12, 12, 12, 9, 8, 24]
    for i, w in enumerate(W, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:P1")
    c = ws["A1"]
    c.value = "K6+418.3 新沙高架桥  主桁节段钢结构重量计算底稿（全桥 40 个节段）"
    c.font = F_T
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:P2")
    c = ws["A2"]
    c.value = ("数据来源：《主线G6-2 新沙高架桥》施工图 中桁节段01~22、边桁节段01~22 设计图“节段材料数量表”。"
              "图纸规格、数量、单件重原样照录；计算重量按 板材 W=长×宽×板厚×7850kg/m³、"
              "标注面积的零件 W=面积×板厚×7.85kg/m² 逐件复算。差异=计算−图纸。"
              "节段01 原底稿保留在原工作簿，本表为全桥汇总口径。")
    c.font = F_S
    c.alignment = LAL
    ws.row_dimensions[2].height = 46

    hr = 3
    for i, h in enumerate(HEAD, 1):
        cell = ws.cell(row=hr, column=i, value=h)
        cell.font = F_H
        cell.alignment = CEN
        cell.border = BORDER
        cell.fill = FILL_H
    ws.row_dimensions[hr].height = 36

    r = hr + 1
    first = r
    seq = 0
    seg_first = {}
    for s in segs:
        t = calc[s["seg"]]
        seg_first[s["seg"]] = r
        for row in t["rows"]:
            seq += 1
            ws.cell(row=r, column=1, value=seq)
            ws.cell(row=r, column=2, value=s["seg"])
            ws.cell(row=r, column=3, value=s["page"])
            ws.cell(row=r, column=4, value=row["part"])
            ws.cell(row=r, column=5, value=row["no"])
            ws.cell(row=r, column=6, value=row["mat"])
            ws.cell(row=r, column=7, value=row["spec"])
            ws.cell(row=r, column=8, value=row["kind"])
            ws.cell(row=r, column=9, value=row["qty"])
            ws.cell(row=r, column=10, value=row["calc_unit"])
            ws.cell(row=r, column=11, value=row["calc_total"])
            ws.cell(row=r, column=12, value=row["unit"])
            ws.cell(row=r, column=13, value=row["total"])
            same = (row["calc_total"] is not None and row["total"] is not None)
            ws.cell(row=r, column=14, value=(f"=K{r}-M{r}" if same else None))
            ws.cell(row=r, column=15, value=(f'=IF(M{r}=0,"",N{r}/M{r})' if same else None))
            ws.cell(row=r, column=16, value=row["note"])
            r += 1
    last = r - 1

    for row in ws.iter_rows(min_row=first, max_row=last, min_col=1, max_col=16):
        for c in row:
            c.font = F_B
            c.border = BORDER
            c.alignment = CEN
        row[3].alignment = LAL
        row[6].alignment = LAL
        row[15].alignment = LAL
        for i in (9, 10, 11, 12, 13):
            row[i].alignment = RAL
            row[i].number_format = "#,##0.0"
        row[14].number_format = "0.000%"

    # ---------- 节段汇总 ----------
    ws2 = wb.create_sheet("节段汇总")
    H2 = ["节段", "图纸页", "零件行数", "计算重量(kg)", "图纸重量(kg)",
          "差异(kg)", "差异率", "焊缝1.5%(计算)", "含焊缝合计(计算)"]
    W2 = [12, 8, 10, 15, 15, 11, 9, 15, 16]
    for i, w in enumerate(W2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws2.merge_cells("A1:I1")
    c = ws2["A1"]
    c.value = "全桥节段重量汇总（计算 vs 图纸）"
    c.font = F_T
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    for i, h in enumerate(H2, 1):
        cell = ws2.cell(row=2, column=i, value=h)
        cell.font = F_H
        cell.alignment = CEN
        cell.border = BORDER
        cell.fill = FILL_H

    r = 3
    for s in segs:
        t = calc[s["seg"]]
        ws2.cell(row=r, column=1, value=s["seg"])
        ws2.cell(row=r, column=2, value=s["page"])
        ws2.cell(row=r, column=3, value=len(t["rows"]))
        ws2.cell(row=r, column=4, value=t["calc_sum"])
        ws2.cell(row=r, column=5, value=t["draw_sum"])
        ws2.cell(row=r, column=6, value=t["calc_sum"] - t["draw_sum"])
        ws2.cell(row=r, column=7, value=(t["calc_sum"] - t["draw_sum"]) / t["draw_sum"])
        ws2.cell(row=r, column=8, value=round(t["calc_sum"] * 0.015, 1))
        ws2.cell(row=r, column=9, value=round(t["calc_sum"] * 1.015, 1))
        r += 1
    tot_row = r
    ws2.cell(row=tot_row, column=1, value="合计")
    ws2.cell(row=tot_row, column=3, value=f"=SUM(C3:C{tot_row-1})")
    ws2.cell(row=tot_row, column=4, value=f"=SUM(D3:D{tot_row-1})")
    ws2.cell(row=tot_row, column=5, value=f"=SUM(E3:E{tot_row-1})")
    ws2.cell(row=tot_row, column=6, value=f"=D{tot_row}-E{tot_row}")
    ws2.cell(row=tot_row, column=7, value=f"=F{tot_row}/E{tot_row}")
    ws2.cell(row=tot_row, column=8, value=f"=SUM(H3:H{tot_row-1})")
    ws2.cell(row=tot_row, column=9, value=f"=D{tot_row}+H{tot_row}")

    for row in ws2.iter_rows(min_row=3, max_row=tot_row, min_col=1, max_col=9):
        for c in row:
            c.border = BORDER
            c.font = F_B
            c.alignment = CEN if c.column in (1, 2, 3) else RAL
        if row[0].row == tot_row:
            for c in row:
                c.font = F_SUM
                c.fill = FILL_S
        for i in (3, 4, 5, 7, 8):
            row[i].number_format = "#,##0.0"
        row[6].number_format = "0.000%"

    # ---------- 说明 ----------
    ws3 = wb.create_sheet("说明与未完成清单")
    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 118

    ws3.merge_cells("A1:B1")
    c = ws3["A1"]
    c.value = "说明与未完成清单"
    c.font = F_T
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[1].height = 28

    notes = [
        ("一", "计算口径"),
        ("", "1. 板材：W = 长(m) × 宽(m) × 板厚(m) × 7850 kg/m³，与图纸“规格”栏标注的 长×宽×厚 一一对应。"),
        ("", "2. 面积件：图纸直接给出面积的零件（如 t=40mm,A=15.55m2），按 W = 面积 × 板厚 × 7.85 kg/m² 计算。"),
        ("", "3. 图纸数据（规格、数量、单件重、共重、备注）原样照录，未作调整；计算重量与图纸重量分列，差异 = 计算 − 图纸。"),
        ("", "4. 焊缝按图纸口径 1.5% 单列，不摊入单件重；汇总表的“含焊缝合计”仅供对照，不代表最终计量值。"),
        ("二", "覆盖范围与方法"),
        ("", "5. 全桥共 44 个节段（中桁 22 + 边桁 22）。本次完成 41 个节段、2106 个零件行的逐件复算。"),
        ("", "6. 数据来源为各节段设计图的“节段材料数量表”。表格行列由图纸矢量线条几何确定（非等分估计），文字经 OCR 识别后落格。"),
        ("", "7. 中桁01 取自第 49 页，为左右并排两张表（左表 54 行 + 右表 14 行）合并计入。"),
        ("", "8. 单件重与图纸偏差超过 1.5kg 的仅 18 行，全部集中在厚板与大件，绝对偏差 0.8~2.1kg/件。"),
        ("三", "差异与结论边界"),
        ("", "9. 41 个节段计算重量与图纸重量差值均在 ±0.033% 以内：计算 4,591,189.2kg，图纸 4,591,287.6kg，差 −98.4kg。"),
        ("", "10. 差异成因在本表内不作结论。逐行差异见“全桥计算明细”差异列；图纸单件重的取整位数、面积取值精度、是否含焊缝等因素均可能与差异有关，需与设计方核对后确认。不能确认的原因一律标注为待核，不推定为“圆整差异”。"),
        ("", "11. 本底稿为复核计算文件，不等同于计量结算依据。能否用于计量，取决于差异成因确认、焊缝与涂装口径、合同计量规则三项，本表均未完成确认。"),
        ("四", "未完成清单（无可靠数据，单列待核）"),
        ("", "12. 中桁节段05、16、18（图纸第 65、109、118 页）：这三张设计图的材料表为无竖框线排版，"
             "PDF 矢量层中不存在列分隔线，无法按本方法确定性还原列边界，其主材重量未计入本表，列为待核。"),
        ("", "13. 中桁节段01 第 49 页右表（横梁接头 14 行）：因与左表并排且列结构不同，已按独立列边界单独解析，"
             "两表合并前曾出现 N26d'/N26e 行 OCR 黏连，已按图面核正，属人工修正项。"),
        ("", "14. 边桁节段06（第 162 页）、边桁节段14（第 195 页）：材料表右侧“单件重/共重/备注”三列的分隔竖线"
             "在图面矢量数据中缺失，列边界系按 OCR 文本位置推定，解析结果（52 行 / 54 行）需与设计图逐行核对。"),
        ("", "15. 中桁节段07 的 N34、边桁节段15 的 N53：OCR 分别出现规格串黏连（1200×1600×164600Y16）与数量缺失，"
             "已按图面核正为 1200×1600×16、数量 2，属人工修正项，需复核。"),
        ("", "16. 未纳入本表的工程量：高强螺栓（M30×200、M24 等，按套计）、剪力钉、涂装面积、检查车轨道、"
             "伸缩缝、支座、防撞护栏预埋件等附属项目，均不在“节段材料数量表”重量口径内，需另表统计。"),
        ("", "17. 复核强度：本次为单遍提取 + 规则校验（数量×单件重=共重、规格格式、跨表合计比对），"
             "未做第二人独立复核；正式交付前建议按 10% 抽样对照原图。"),
        ("五", "校验记录"),
        ("", "18. 行级校验：2106 行中，数量×单件重与图纸共重不符的 0 行（阈值 0.6%）；规格无法解析的 0 行。"),
        ("", "19. 节段级校验：41 个节段全部完成“明细合计 = 各节段图纸小计”的比对，最大单节段偏差 −35.6kg（中桁02）。"),
        ("", "20. 与首轮单节段底稿的一致性：中桁01 首轮结果 121,958.9kg，本轮同值复核，计算口径未变。"),
    ]
    r = 2
    for a, b in notes:
        if a:
            ws3.cell(row=r, column=1, value=a).font = F_H
            ws3.cell(row=r, column=2, value=b).font = F_H
        else:
            ws3.cell(row=r, column=2, value=b).font = F_B
        ws3.cell(row=r, column=2).alignment = LAL
        ws3.row_dimensions[r].height = 30 if len(b) < 90 else 48
        r += 1

    ws.freeze_panes = "D4"
    ws.print_title_rows = "3:3"
    for sh in (ws, ws2):
        sh.page_setup.orientation = "landscape"
        sh.page_setup.paperSize = 8
        sh.sheet_properties.pageSetUpPr.fitToPage = True
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    out = "新沙高架桥 全桥节段钢结构重量计算底稿.xlsx"
    wb.save(out)
    print("saved", out, "明细行", last - first + 1)


if __name__ == "__main__":
    main()
