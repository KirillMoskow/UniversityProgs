import pandas as pd 
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilenames

# функция выбора нескольких файлов с помощью диалогового окна
def open_files():
    root = Tk()
    root.withdraw()
    root.files = askopenfilenames(title="Choose a file", filetypes=[('Excel Files', '*.xlsx')])
    return root.files

files = open_files

for file in files:
    # Читаем файл excel
    df = pd.read_excel('1.xlsx', sheet_name = 'Лист1')
    # Создаем словарь для записи данных
    dicti = {}
    # Перебираем строки в файле excel и записываем данные в словарь
    for i in df.values:
        if i[1] not in dicti:
            dicti[i[1]] = {'T': [], 'e\'': [], 'e"': [], 'tg' : []}
            dicti[i[1]]['T'].append(i[0])
            dicti[i[1]]['e\''].append(i[2])
            dicti[i[1]]['e"'].append(i[3])
            dicti[i[1]]['tg'].append(i[4])
        else:
            dicti[i[1]]['T'].append(i[0])
            dicti[i[1]]['e\''].append(i[2])
            dicti[i[1]]['e"'].append(i[3])
            dicti[i[1]]['tg'].append(i[4])


    # Вставляем значения T,e',e"
    writer = pd.ExcelWriter('output.xlsx')

    for i,j in enumerate(dicti):
        T = dicti[j]['T']
        e_1 = dicti[j]['e\'']
        e_2 = dicti[j]['e"']
        tg = dicti[j]['tg']
        T_data = pd.DataFrame(T)
        e_1_data = pd.DataFrame(e_1)
        e_2_data = pd.DataFrame(e_2)
        tg_data = pd.DataFrame(tg)
        T_data = T_data.rename(columns={0 : 'T'})
        e_1_data = e_1_data.rename(columns={0 : 'e\''})
        e_2_data = e_2_data.rename(columns={0 : 'e"'})
        tg_data = tg_data.rename(columns={0 : 'tg'})

        for k,l in enumerate(T_data):
            T_data.to_excel(writer,sheet_name='e\' to T', startcol = i*2, startrow = 1 , index = False)
            e_1_data.to_excel(writer,sheet_name='e\' to T', startcol = i*2 + 1, startrow = 1 , index = False)
        for k,l in enumerate(T_data):
            T_data.to_excel(writer,sheet_name='e" to T', startcol = i*2, startrow = 1 , index = False)
            e_2_data.to_excel(writer,sheet_name='e" to T', startcol = i*2 + 1, startrow = 1 , index = False)
        for k,l in enumerate(T_data):
            T_data.to_excel(writer,sheet_name='tg to T', startcol = i*2, startrow = 1 , index = False)
            tg_data.to_excel(writer,sheet_name='tg to T', startcol = i*2 + 1, startrow = 1 , index = False)
        


    #Сохраняем в файл
    writer._save()

    # Вставляем значения частоты
    wb = openpyxl.load_workbook('output.xlsx')
    for k in ['e\' to T', 'e" to T','tg to T']:
        ws = wb[k]
        for i, j in enumerate(dicti):
            ws.cell(row=1, column=i*2 + 1 , value = j)

    #Сохраняем


    wb.save('output.xlsx')