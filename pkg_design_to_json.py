from openpyxl import load_workbook
import json

wb = load_workbook('new_one_month.xlsx')

sheetnames = wb.sheetnames

pkg_design = {}

for name in sheetnames:
    if name not in ['Summary', ' back_to_manual', 'TCID']:
        tmp = []
        for case in wb[name].iter_rows(max_col=1, values_only=True):
            if case[0] and case[0].lower()[:3] == 'tc_':
                tmp.append(case[0].lower())
        pkg_design[name] = tmp

print(pkg_design)