import pandas as pd
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
df = pd.read_csv(OUT + r'\normalized_tables.csv', encoding='utf-8-sig')
hz = df[df['构件类别'] == '汇总(厚度)']
print('汇总(厚度) rows', len(hz))
print(hz[['梁板编号','板件名称','位置','材质','长','宽','厚','数量','单件重','全桥数量','备注','来源页','表名']].to_string())

print('\n=== 其他/附属/预埋/钢筋 类别 抽样 ===')
oth = df[df['构件类别'].isin(['其他','附属','预埋件附属','钢筋','桥面板'])]
print(oth[['板件名称','构件类别','厚','数量','单件重','来源页','表名']].head(30).to_string())
