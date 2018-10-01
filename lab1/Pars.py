import xlrd, re

wb = xlrd.open_workbook('gazizov_minibaev\lab1\RIS.xlsx')

sheet = wb.sheet_by_index(0)
for s in wb.sheets():
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value = (s.cell(row,col).value)
            try : value = str(int(value))
            except : pass
            if (value != ''):
                #print(value, end="\n")
                col_value.append(value)
        values.append(col_value)
# print(values)
res = re.sub(r'','',str(values))
res = re.sub(r'[,\[\]\']|(\s){2,}|(\\xa0)','',res)
print(res)