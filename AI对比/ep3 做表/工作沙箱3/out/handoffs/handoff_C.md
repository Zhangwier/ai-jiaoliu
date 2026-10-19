# 阶段C Handoff：OCR / 矢量视觉表格抽取

## 已完成
- 逐页处理 `table_pages.pdf` 全部 **155 页**（0 抽样，全部跑完，无中断），页状态全部 `ok`。
- 用矢量线（`get_drawings()` 的 `l`/`re`/`qu`，长度≥15pt，排除底部图签带 y∈[752,860]）合并成行/列层级线，按互交双团构造单元格网格（每页 1~4 个表）：1 表 53 页 / 2 表 64 页 / 3 表 37 页 / 4 表 1 页。
- 单元格合并按“跨该边界的线是否真实覆盖本行/列（覆盖≥60%）”用并查集判定，输出 `rowspan`/`colspan`。
- 取文本优先级：文字层（`get_text('words')`，`source=text`，置信度=1.0）→ 表区域渲染 300dpi 送 RapidOCR（`source=ocr`，置信度取 OCR score，只填空格不覆盖文字层）。
- 表题：优先网格上方文字层；纯 OCR 表则 OCR 网格上方 30~40pt 条带取真实表题（保守过滤：CJK 占比/长度/数字占比），失败回退表头首格文字。
- 产出：`raw_tables.csv`(25191 行非空单元格)、`raw_tables_meta.json`、20 张网格复核 PNG。

## 产物路径
- `工作沙箱3\out\raw_tables.csv`（UTF-8 with BOM，14 列）
- `工作沙箱3\out\raw_tables_meta.json`
- `工作沙箱3\out\handoffs\handoff_C.md`（本文件）
- 复核图：`_ep3_scratch\c_peek\`（20 张 `pkg###_t#.png`，红框=识别网格）
- 脚本：`_ep3_scratch\stage_c_extract.py`（主）、`c_render_peek.py`（复核图）、`detect_grid.py`（网格检测）

## 关键字段
- CSV 列：`page`(原页号 page_index) / `pkg_page`(表页序 1..155) / `表序号`(页内 0,1,2…) / `行号` / `列号` / `文本` / `置信度` / `table_title` / `source`(text/ocr) / `cell_bbox`(x0,y0,x1,y1) / `merge_group` / `continued_from` / `rowspan` / `colspan`。仅输出非空单元格。
- 规模：**表格 296 个，单元格总数 33889，非空输出行 25191**（空格 8698，25.7%）。
- 来源：`text`=748(2.2%)，`ocr`=24443(72.1%)；**OCR 平均置信度 0.968**（267 个含 OCR 的表）。
- 表题：296 表中 204 个为中文表题（如 `节段04材料数量表`、`A1a桥面板钢筋数量表`、`节段10~节段11拼接数量表`）。
- 跨页续表组：**1 组**（`MG235`，pkg119/120，表题 `编号`，7 列，`continued_from` 串联）。
- meta：每页/每表含 表题、n_rows/n_cols、cell_total、filled、text/ocr 单元格数、text_coverage、ocr_mean_conf、low_quality、grid_bbox、status、merge_group、continued_from。

## 未决问题
- 155 页均检出矢量网格，**无页触发“无网格整页 OCR”回退**（无 `no_grid` 页）。
- 存在“绘图细节表”网格（焊材/孔位/符号表，如 pkg51 `R250`、翼缘尺寸表），多为小表；仅 1 张判 `low_quality`，其余保留但表题与置信度偏低，需下游甄别。
- 跨页续表仅 1 组：改用真实表题后，相邻页表题普遍不同（如 A1a/A1b…桥面板钢筋数量表），按“表名一致”不再合并；阶段B 用通用表头（“名称”“编号”）曾合出 8 组，**是否真实连续需人工确认**。MG235 两页表题仍为 OCR 失败的 `编号`。
- 保留 CAD 字体误识别原样（如 `<熔透`、`iizlA`、`0/O`），未改数值；个别密集小表存在跨格合并或列对齐偏差。
- 部分页表题取自网格上方矢量文字 OCR，个别仍为表头首格文字（`R250`、`规格（mm)`）。

## 下一步
- 阶段D：依 `template_map.json` 的 `field_columns`/`fill_plan`，将 `raw_tables.csv` 中“材料数量表/拼接数量表”类映射到模板字段；用 `cell_bbox`+`rowspan/colspan`+`行号/列号` 重建行列对齐再写值，禁止覆盖公式列。
- 建议下游按 `table_title`（含 材料数量/拼接数量/数量表）与 `merge_group` 过滤，忽略 `low_quality` 及焊材细节表；跨页续表按 `merge_group`+`continued_from` 并表。
