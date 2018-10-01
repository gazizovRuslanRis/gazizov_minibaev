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
    #print(result)

    days = ['понедельник',
            'вторник',
            'среда',
            'четверг',
            'пятница',
            'суббота']

    times = ['8:00' ,
             '9:40' ,
             '11:30',
             '13:20',
             '15:00',
             '16:40']

    dict = {'понедельник' : '',
            'вторник'     : '',
            'среда'       : '',
            'четверг'     : '',
            'пятница'     : '',
            'суббота'     : ''}

    dict_times = {'8:00' : '',
                 '9:40'  : '',
                 '11:30' : '',
                 '13:20' : '',
                 '15:00' : '',
                 '16:40' : ''}

    for d in range(len(days)):
        if d == (len(days)-1):
            temp = re.search(r'(?<=%s, ).+' % (days[d]), result)
            str = ''.join(temp[0])
        else:
            temp = re.search(r'(?<=%s, ).+(?=, %s)' % (days[d], days[d+1]), result)
            str = ''.join(temp[0])

        for t in range(len(times)):
            if t == (len(times)-1):
                tmp = re.findall(r'(?<=%s, ).+' % (times[t]), str)
                dict_times[times[t]] = tmp
            else:
                tmp = re.findall(r'(?<=%s, ).+(?=, %s)' % (times[t], times[t + 1]), str)
                dict_times[times[t]] = tmp
        dict[days[d]] = dict_times.copy()

    for key in dict:
        print('\n' + key)
        for subkey in dict[key]:
            print('%s -> %s' % (subkey, dict[key][subkey]))

    with open('data.json', 'w') as f:
        json.dump(dict, f, ensure_ascii=False, indent=2)