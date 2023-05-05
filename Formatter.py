import pandas as pd 
import openpyxl
from tkinter import filedialog, Tk
import os

# создаем основное окно приложения
root = Tk()
root.withdraw() # скрываем окно

# открываем диалоговое окно для выбора файлов
files = filedialog.askopenfilenames(title="Выберите файлы",  # files - кортеж выбранных файлов
                                    filetypes=(("Excel Файлы", "*.xlsx"),
                                               ("все файлы", "*.*")))


file_names = []

#Записываем пути файлов в список
for file in files:
    file_names.append(file)

names = []

#Выделяем из пути только название файла и записываем в список
for file in file_names:
    filename = os.path.basename(file)
    names.append(filename)

#Используем полученные названия в цикле
for file in names:

    # Читаем файл excel
    df = pd.read_excel(file)

    #Исправление температуры
    for i in range(1, len(df) - 1):
        if df.loc[i, 'T'] == 300: # .loc[Строки, столбцы]
            df.loc[i, 'T'] = (df.loc[i-1, 'T'] + df.loc[i+1, 'T']) / 2
        elif df.loc[len(df) - 1, 'T'] == 300:
            continue

    # Создаем словарь для записи данных
    dicti = {}

    # Перебираем строки в файле excel и записываем данные в словарь
    for i in df.values:
        dicti.setdefault(i[1],{'T': [], 'e\'': [], 'e"': [], 'tg': []})
        dicti[i[1]]['T'].append(i[0])
        dicti[i[1]]['e\''].append(i[2])
        dicti[i[1]]['e"'].append(i[3])
        dicti[i[1]]['tg'].append(i[4])

    #Создаем файл Excel в который запишем нужные данные
    writer = pd.ExcelWriter('out_' + file)

    #Запись значений в соответствующие листы
    for i,j in enumerate(dicti):
        #Записываем списки в переменные
        T = dicti[j]['T']
        e_1 = dicti[j]['e\'']
        e_2 = dicti[j]['e"']
        tg = dicti[j]['tg']

        #Записываем каждую переменную в DF
        T_data = pd.DataFrame(T)
        e_1_data = pd.DataFrame(e_1)
        e_2_data = pd.DataFrame(e_2)
        tg_data = pd.DataFrame(tg)

        #Переименовываем колонки
        T_data = T_data.rename(columns={0 : 'T'})
        e_1_data = e_1_data.rename(columns={0 : 'e\''})
        e_2_data = e_2_data.rename(columns={0 : 'e"'})
        tg_data = tg_data.rename(columns={0 : 'tg'})

        #Записываем значения полученных DF в нужные листы
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
    wb = openpyxl.load_workbook('out_' + file)
    for k in ['e\' to T', 'e" to T','tg to T']:
        ws = wb[k]
        for i, j in enumerate(dicti):
            ws.cell(row=1, column=i*2 + 1 , value = j)

    #Сохраняем
    wb.save('out_' + file)