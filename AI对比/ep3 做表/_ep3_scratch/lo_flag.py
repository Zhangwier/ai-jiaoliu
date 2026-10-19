import openpyxl, shutil, os
OUT = r'd:\MG\program\ai交流\AI对比\ep3 做表\工作沙箱3\out'
src = OUT + r'\v2_edit.xlsx'
# set recalc-on-load flag on the INTERMEDIATE (not on the LO output)
wb = openpyxl.load_workbook(src)
try:
    wb.calculation.fullCalcOnLoad = True
except Exception as e:
    print('calcflag', e)
wb.save(src)
print('flagged', os.path.getsize(src))

dst = r'C:\temp\ep3recalc'
os.makedirs(dst, exist_ok=True)
shutil.copy(src, os.path.join(dst, 'v2_edit.xlsx'))
# also keep an untouched v1 copy for the diff later
shutil.copy(OUT + r'\final.xlsx', os.path.join(dst, 'v1_final.xlsx'))
print('copied to', dst)
