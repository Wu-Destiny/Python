# -*- coding: UTF-8 -*-
import openpyxl
print ('Opening excel')
content = openpyxl.load_workbook('gk.xlsx')
sheet = content.get_sheet_by_name('中央国家行政机关直属机构(省级及以下)')
print ('Reading rows')
d = {}
key = ""
for row in range(9271,9549):
    State = sheet['M'+str(row)].value
    if key in State:
        if "2016" in sheet['W'+str(row)].value:
            print (key,row,sheet['C'+str(row)].value,sheet['W'+str(row)].value)