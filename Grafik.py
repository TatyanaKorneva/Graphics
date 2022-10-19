import pandas
import openpyxl
from datetime import datetime, timedelta


def update_spreadsheet(path: str, _df, starcol: int = 1, startrow: int = 1, sheet_name: str = "ToUpdate"):
    '''

    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param starcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return:
    '''
    wb = openpyxl.load_workbook(path)
    for ir in range(0, len(_df)):
        for ic in range(0, len(_df.iloc[ir])):
            wb[sheet_name].cell(startrow + ir, starcol + ic).value = _df.iloc[ir][ic]
    wb.save(path)

df = pandas.read_excel('C:/Users/79042/PycharmProjects/pythonProject/graphics/Graphics_2023.xlsx', sheet_name='График', header=1)

# print whole sheet data
# print(excel_data_df['ФИО'])

count=0
s=[]
spisok_1=[]
spisok_2=[]
spisok_vs=[]
slovar = dict()
spisok_3=[]
for j in range(len(df)):
    # print(df.iloc[j]['ФИО'])
    for columnIndex, value in df.iloc[j].items():
        # print(columnIndex, value)
        if value==1 and columnIndex!='№':
            count += 1
            # print(columnIndex, value)
        else:
            if count!=0:
                s.append(count)
                h1=columnIndex - timedelta(days=count)
                h2=columnIndex - timedelta(days=1)
                spisok_1.append(h1.date())
                spisok_2.append(h2.date())
                count=0
                spisok_vs.append(df.iloc[j]['ФИО'])

    spisok_3 = spisok_1.copy()
    spisok_4 = spisok_2.copy()
    intermediate_dictionary = {'ФИО':spisok_vs}
    intermediate_dictionary2 = {'Дата начала': spisok_1, 'Дата окончания': spisok_2}
    df1 = pandas.DataFrame(intermediate_dictionary)
    df2 = pandas.DataFrame(intermediate_dictionary2)


    # writer = pandas.ExcelWriter('C:/Users/79042/PycharmProjects/pythonProject/graphics/Graphics_2023(ch).xlsx', engine='openpyxl')
    # pandas_dataframe.to_excel(writer, sheet_name='Sheet1', header=None, index=False, startcol=7,startrow=6)
    # writer.close()

    # df.to_excel(path, sheet_name="sheet1")
    # df.sample(10).to_excel('C:/Users/79042/OneDrive/Рабочий стол/График отпусков_2023.xlsx', sheet_name='Sheet1')

update_spreadsheet('C:/Users/79042/PycharmProjects/pythonProject/graphics/Graphics_2023.xlsx', df1, 4, 23,"График по форме")
update_spreadsheet('C:/Users/79042/PycharmProjects/pythonProject/graphics/Graphics_2023.xlsx', df2, 10, 23,"График по форме")



print(df1, df2)