# 阶段A 交接 — 模板结构解析

## 已完成
- 以 openpyxl(data_only=False) 只读解析模板 `潭州高架桥（计算稿模板）.xlsx`，未修改模板。
- 5 个 sheet 全部解析：`合计数量`(A1:I18,18×9,10合并)、`主梁数量`(A1:W435,435×23,315合并)、`横隔板`(A1:AA142,142×27,362合并)、`梁底挡块`(A1:U8,8×21,24合并)、`附属预埋件`(A1:U11,11×21,24合并)。
- 产出结构映射：字段列、公式列（含原样公式与归一化模式 {r}）、按厚度汇总列与条件值行、样式样例行、跨表引用、合并单元格、颜色图例、fill_plan。

## 产物路径
- 主产物：`d:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out\template_map.json`（124~137KB，合法 UTF-8 JSON）
- 本交接：`d:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out\handoffs\handoff_A.md`
- 解析脚本（scratch）：`d:\MG\program\ai交流\AI对比\ep3 做表\_ep3_scratch\gen_map.py`

## 关键字段
- 版本：`meta` 标注 openpyxl/data_only=False，`{r}` 表示随行变化的相对行号。
- `sheets`：各表 dims/max_row/max_column/merged_count。
- `field_columns`（需填）：主梁数量 A/B/C/D/E/F/G/I/K；横隔板 A/B/C/D/E/F/G/H/J/L；梁底挡块、附属预埋件同横隔板；合计数量仅 D(图纸数量Q355)/E(图纸数量Q420)/I备注。
- `formula_columns`（禁止覆盖，含 `formulas_by_row` 原样公式）：主梁 H`=(E{r}/1000)*(F{r}/1000)*(G{r}/1000)*7850`(起始4,413条)、J`=H{r}*I{r}`、L`=A{r}`；横隔 I`=(E{r}/1000)*(F{r}/1000)*(G{r}/1000)*7850*H{r}`、K`=I{r}*J{r}`；挡块/预埋 I、K 同横隔；合计 F`=D{r}+E{r}`、G(跨表引用)、H`=G{r}-F{r}`。
- `summary_columns`：主梁数量复核块 M–S 为 `=SUMIF(G{start}:G{end-1},$X$3,J...)`（T 合计、U 图纸数量、V 差异），条件值行 row3：M3=12,N3=14,O3=16,P3=20,Q3=24,R3=30,S3=40,T3='合计'；`block_summary_rows` 给出 16 块块首行公式；横隔 21 块；合计数量 G/H 为汇总列。
- `cross_sheet_refs`：13 条，如 `合计数量!G5 = '=主梁数量!M432+横隔板!O139'`。
- `style_sample_rows`：表头行=2、单位行=3、数据起始行=4；含 font 宋体/9、border thin、alignment center、number_format（H=`0.00_ `, M/T=`0_ `）。
- `color_legend`：主梁红底 FFFF0000×48、横隔×11、合计×8 红 + A15/A16 黄 FFFFFF00；均 `inferred:true`。
- `fill_plan`：逐表给出数据来源、需填列、禁止覆盖的公式列、数据起始/结束行。
- `extras`：主梁数据行区间 [4,431]（16块，小计行在每块末行）；M/N 覆盖 4~431（每块 SUMIF 覆盖本块数据行，M432=SUM(M4:M431)）；`summary_row_formulas` 各表末汇总行公式。

## 未决问题
- 模板内 `横隔板!D140` 公式为 `=#REF!`（坏引用），fill 时需规避/修复。
- 颜色含义无显式图例文字，红/黄含义为分布推测（`inferred:true`），建议人工确认。
- 块汇总列公式跨块并不统一（横隔 M–X 多为跨单元格引用 `=K{n}` 而非 SUMIF），下游按 `block_summary_rows` 逐块处理更稳妥。
- 合计数量 G 列是跨表引用公式，其引用的目标行（如 横隔板!N139、主梁数量!M432）依赖下游明细填满后才有值。

## 下一步
- 阶段B/C 按 `fill_plan` 从 PDF 材料表/工程数量表提取数据，仅写 `field_columns` 对应列，绝不覆盖 `formula_columns` 与 `summary_columns`。
- 依 `style_sample_rows` 复用样式；行区间用 `extras.主梁数量_data_row_range` 与各表 data_row_start/end 对齐。
