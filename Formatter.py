import pandas as pd 
import openpyxl

# Читаем файл excel
df = pd.read_excel('1.xlsx', sheet_name='Лист1')
# Создаем словарь для записи данных
dicti = {}

# Перебираем строки в файле excel и записываем данные в словарь
for i in df.values:
    if i[1] not in dicti:
        dicti[i[1]] = {'T': [], 'e\'': [], 'e"': []}
    else:
        dicti[i[1]]['T'].append(i[0])
        dicti[i[1]]['e\''].append(i[2])
        dicti[i[1]]['e"'].append(i[3])


# Вставляем значения T,e',e"
writer = pd.ExcelWriter('output.xlsx')

freq = []
for i in dicti:
    freq.append(i)

for i,j in enumerate(dicti):
    data = pd.DataFrame(dicti[j])
    for k,l in enumerate(data):
        data.to_excel(writer, startcol = i*len(dicti[j]), startrow = 1 , index = False)

#Сохраняем в файл
writer._save()

# Вставляем значения частоты
wb = openpyxl.load_workbook('output.xlsx')
ws = wb.active

for i, j in enumerate(dicti):
    ws.cell(row=1, column=i*len(dicti[j]) + 1 , value = j)

#Сохраняем в файл
wb.save('output.xlsx')