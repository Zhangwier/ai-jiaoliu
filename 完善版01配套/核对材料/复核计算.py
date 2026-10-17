"""基于原PDF逐格读取的参数重算；不读取任何被测模型生成的Excel作为真值。
范围：图内公式与表列长度、分项汇总，不作结构合格判定。
运行 python 复核计算.py，在同目录生成可复核CSV与JSON。
"""
from pathlib import Path
import math, csv, json
from decimal import Decimal

ROOT = Path(__file__).resolve().parent
# 列顺序：部位 B0 B1 B2 B3 alpha gamma 六种表列长度 Φ22小计 Φ10小计 C35
# 来源：桥墩挡块.pdf 第2页，左表自上而下5组，右表自上而下5组；第3页参数另作目视交叉核对。
RAW = [
('1#左',200,47.5,48.7,46.0,89.3,88.5,[260.8,164,208,208,484.2,475.2],253.7,50.3,.45),
('1#右',200,46.1,45.0,47.2,90.7,91.2,[259.5,164,208,208,479.1,470.1],252.9,49.9,.44),
('2#小里程左',100,45.0,46.7,48.4,88.0,92.0,[260.1,164,108,108.1,279.8,270.8],126.7,28.3,.22),
('2#大里程左',100,54.0,56.6,59.1,87.1,92.9,[270.0,164,108,108.1,298.7,289.7],129.6,29.8,.27),
('2#右',200,48.9,45.0,49.6,92.2,92.6,[261.2,164,208,208.2,482.1,473.1],254.0,50.1,.46),
('3#小里程左',100,45.1,48.0,51.0,86.6,93.4,[261.4,164,108,108.2,281.3,272.3],127.0,28.4,.23),
('3#大里程左',100,53.4,57.1,60.8,85.7,94.3,[270.5,164,108,108.2,298.7,289.7],129.8,29.9,.28),
('3#右',200,51.4,45.0,52.6,93.6,94.3,[263.1,164,208,208.5,484.9,475.9],255.1,50.3,.48),
('4#左',200,45.1,53.4,44.1,85.3,84.7,[260.9,164,208,208.7,487.2,478.2],253.8,50.5,.44),
('4#右',200,53.2,45.0,54.3,94.7,95.3,[264.2,164,208,208.7,486.9,477.9],255.8,50.5,.49),
]
def sin(x): return math.sin(math.radians(x))
def cot(x): return 1/math.tan(math.radians(x))
rows=[]; params=[]
for name,b0,b1,b2,b3,a,g,listed,m22,m10,vol in RAW:
    h=28.; beta=90.; n0=17 if b0==200 else 7
    lengths=[221-h+(b1+b2+b3)/3-9/sin(beta)+math.sqrt(h*h+(9/sin(beta))**2),
             164.,20+(b0-12)/sin(beta),
             20+(b0/2-6)/sin(a)+(b0/2-6)/sin(g),
             (b0/2-6)/sin(a)+(b0/2-6)/sin(g)+(b0-12)/sin(beta)+b1+b2+12,
             (b0/2-6)/sin(a)+(b0/2-6)/sin(g)+(b0-12)/sin(beta)+b1+b2-9/sin(beta)+12]
    for j,(given,calc) in enumerate(zip(listed,lengths),1):
        rows.append([name,f'{j}a',given,round(calc,6),round(calc-given,6), '按印刷总长公式，不代表大样分段正确'])
    params.append([name,50,h,b0,b1,b2,b3,a,beta,g,n0,9,m22,m10,vol])
def save(name,head,values):
    with (ROOT/name).open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f);w.writerow(head);w.writerows(values)
save('原图参数逐格记录.csv',['部位','H_cm','h_cm','B0_cm','B1_cm','B2_cm','B3_cm','alpha_deg','beta_deg','gamma_deg','n0','m0_cm','Phi22_subtotal_kg','Phi10_subtotal_kg','C35_m3'],params)
save('60项表列长度与印刷公式复算.csv',['部位','钢筋编号','图表长度_cm','按印刷公式复算_cm','复算减图表_cm','边界'],rows)
sum22=sum(Decimal(str(r[8])) for r in RAW)
sum10=sum(Decimal(str(r[9])) for r in RAW)
sumvol=sum(Decimal(str(r[10])) for r in RAW)
summary={
 'scope':'印刷总长公式与表列长度一致性，不等于构造/下料公式正确性',
 'records':len(rows),'max_abs_length_delta_cm':max(abs(r[4]) for r in rows),
 'length_deltas_over_0_05cm':[r[:5] for r in rows if abs(r[4])>.050001],
 'sum_Phi22_subtotals_kg':str(sum22),'drawing_Phi22_kg':'2038.3',
 'sum_Phi10_subtotals_kg':str(sum10),'drawing_Phi10_kg':'418.1',
 'sum_C35_m3':str(sumvol),'drawing_C35_m3':'3.8',
 'one_a_segment_constant':117+117+22+22-12,'one_a_printed_constant':221,
 'one_a_conditional_difference_cm':117+117+22+22-12-221,
 'one_a_bar_count':sum((20 if r[1]==200 else 10) for r in RAW),
 'one_a_conditional_total_length_m':sum((20 if r[1]==200 else 10) for r in RAW)*.45,
 'rubber_count':10*2,'rubber_volume_m3':10*2*.3*.2*.02,
 'geometry_4left_if_B2_midpoint_cm':45.1+100*cot(85.3),
 'geometry_4left_if_B3_endpoint_cm':53.4-100*cot(84.7),
 'geometry_caution':'按分段平面几何假设的反推，支持定义待澄清，不授权交换参数或修改设计',
 'units':'长度cm、重量kg、混凝土m3。所有差值不直接认定安全或计量后果。'
}
(ROOT/'复算摘要.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(summary,ensure_ascii=False,indent=2))
