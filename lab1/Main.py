import xlrd, xlwt, re, json
#открываем файл
wb = xlrd.open_workbook('./RIS.xlsx')

#выбираем активный лист
sheet = wb.sheet_by_index(0)

for s in wb.sheets():
#print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value = (s.cell(row,col).value)
            try :
                value = str(int(value))
            except : pass
            if (value != ''):
                col_value.append(value)
                values.append(col_value)

result = re.sub(r'(\[\], )|[\[\]\']|(\s){2,}', '', str(values))
result = re.sub(r'(, ,)', ',', result)
result = re.sub(r'\\n|\\xa0', ' ', result)
print(result)