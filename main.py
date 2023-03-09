import matplotlib.pyplot as plt
import openpyxl.reader.excel
from openpyxl import Workbook
from datetime import datetime, date

def getDate(str):
    dateElem = str.split('.')
    return datetime(int(dateElem[2]), int(dateElem[1]), int(dateElem[0]))

def getNamesSchool(list):
    schools = []
    for key in сounter.keys():
        str = key.split(' ')
        start = 0
        for i in range(len(str)):
            for j in str[i]:
                if start != 0:
                    break
                if ord(j) < 64 or ord(j) == 171:
                    start = i
                    break
        schools.append(' '.join(str[start:]))
    return schools


def sortedName(names):
    male, female = [], []
    for name in names:
        if name.endswith(("а", "я")):
            female.append(name)
        else:
            male.append(name)
    return male, female


wb = openpyxl.reader.excel.load_workbook(filename='Копия медалисты (2).xlsx')
wb.active = 0
sheet = wb.active
mapRegionCountSchool = {
    'Алексеевский район': {'school': [0], "name": [], "date": []},
    'Белгородский район': {'school': [0], "name": [], "date": []},
    'Борисовский район': {'school': [0], "name": [], "date": []},
    'Валуйский район': {'school': [0], "name": [], "date": []},
    'Вейделевский район': {'school': [0], "name": [], "date": []},
    'Волоконовский район': {'school': [0], "name": [], "date": []},
    'Грайворонский район': {'school': [0], "name": [], "date": []},
    'Губкинский район': {'school': [0], "name": [], "date": []},
    'Ивнянский район': {'school': [0], "name": [], "date": []},
    'Корочанский район': {'school': [0], "name": [], "date": []},
    'Красненский район': {'school': [0], "name": [], "date": []},
    'Красногвардейский район': {'school': [0], "name": [], "date": []},
    'Краснояружский район': {'school': [0], "name": [], "date": []},
    'Новооскольский район': {'school': [0], "name": [], "date": []},
    'Прохоровский район': {'school': [0], "name": [], "date": []},
    'Ракитянский район': {'school': [0], "name": [], "date": []},
    'Ровеньский район': {'school': [0], "name": [], "date": []},
    'Старооскольский район': {'school': [0], "name": [], "date": []},
    'Чернянский район': {'school': [0], "name": [], "date": []},
    'Шебекинский район': {'school': [0], "name": [], "date": []},
    'Яковлевский район': {'school': [0], "name": [], "date": []},
}
for i in range(4, 846):
    school = sheet['F' + str(i)].value.lower().replace('белгородской области', '')
    name = sheet['D' + str(i)].value
    date = sheet['E' + str(i)].value
    for key in mapRegionCountSchool.keys():
        if school.find(key[:-10].lower()) != -1:
            mapRegionCountSchool[key]['school'][0] += 1
            mapRegionCountSchool[key]['school'].append(school)
            mapRegionCountSchool[key]['name'].append(name)
            mapRegionCountSchool[key]['date'].append(date)
            break
        if key == 'Яковлевский район':
            if school.find('новый оскол') != -1:
                mapRegionCountSchool['Новооскольский район']['school'][0] += 1
                mapRegionCountSchool['Новооскольский район']['school'].append(school)
                mapRegionCountSchool['Новооскольский район']['name'].append(name)
                mapRegionCountSchool['Новооскольский район']['date'].append(date)
                break
            if school.find('средняя общеобразовательная школа № 5') != -1:
                mapRegionCountSchool['Шебекинский район']['school'][0] += 1
                mapRegionCountSchool['Шебекинский район']['school'].append(school)
                mapRegionCountSchool['Шебекинский район']['name'].append(name)
                mapRegionCountSchool['Шебекинский район']['date'].append(date)
                break
            if school.find('слобожанщина') != -1:
                mapRegionCountSchool['Краснояружский район']['school'][0] += 1
                mapRegionCountSchool['Краснояружский район']['school'].append(school)
                mapRegionCountSchool['Краснояружский район']['name'].append(name)
                mapRegionCountSchool['Краснояружский район']['date'].append(date)
                break
            if school.find('пятницкая') != -1:
                mapRegionCountSchool['Волоконовский район']['school'][0] += 1
                mapRegionCountSchool['Волоконовский район']['school'].append(school)
                mapRegionCountSchool['Волоконовский район']['name'].append(name)
                mapRegionCountSchool['Волоконовский район']['date'].append(date)
                break
            if school.find('новоуколовская') != -1:
                mapRegionCountSchool['Красненский район']['school'][0] += 1
                mapRegionCountSchool['Красненский район']['school'].append(school)
                mapRegionCountSchool['Красненский район']['name'].append(name)
                mapRegionCountSchool['Красненский район']['date'].append(date)
                break
            if school.find('верхопенская') != -1:
                mapRegionCountSchool['Ивнянский район']['school'][0] += 1
                mapRegionCountSchool['Ивнянский район']['school'].append(school)
                mapRegionCountSchool['Ивнянский район']['name'].append(name)
                mapRegionCountSchool['Ивнянский район']['date'].append(date)
                break
            if school.find('бирюч') != -1 or school.find('белогорский') != -1:
                mapRegionCountSchool['Красногвардейский район']['school'][0] += 1
                mapRegionCountSchool['Красногвардейский район']['school'].append(school)
                mapRegionCountSchool['Красногвардейский район']['name'].append(name)
                mapRegionCountSchool['Красногвардейский район']['date'].append(date)
                break
            if school.find('строитель') != -1:
                mapRegionCountSchool['Яковлевский район']['school'][0] += 1
                mapRegionCountSchool['Яковлевский район']['school'].append(school)
                mapRegionCountSchool['Яковлевский район']['name'].append(name)
                mapRegionCountSchool['Яковлевский район']['date'].append(date)
                break
            if school.find('стригуновская') != -1:
                mapRegionCountSchool['Борисовский район']['school'][0] += 1
                mapRegionCountSchool['Борисовский район']['school'].append(school)
                mapRegionCountSchool['Борисовский район']['name'].append(name)
                mapRegionCountSchool['Борисовский район']['date'].append(date)
                break
            if school.find('шухов') != -1 or school.find('алгоритм успеха') != -1 or \
                    school.find('средняя общеобразовательная школа № 11') != -1 or \
                    school.find('искорка') != -1:
                mapRegionCountSchool['Белгородский район']['school'][0] += 1
                mapRegionCountSchool['Белгородский район']['school'].append(school)
                mapRegionCountSchool['Белгородский район']['name'].append(name)
                mapRegionCountSchool['Белгородский район']['date'].append(date)
                break
            if school.find('средняя общеобразовательная школа № 40') != -1 or \
                    school.find(
                        'средняя общеобразовательная школа № 28 с углубленным изучением отдельных предметов имени а.а.угарова') != -1 or \
                    school.find('средняя общеобразовательная школа №21') != -1 or \
                    school.find('средняя общеобразовательная школа №30') != -1 or \
                    school.find('средняя политехническая школа №33') != -1 or \
                    school.find('образовательный комплекс "озерки"') != -1 or \
                    school.find(
                        'муниципальное бюджетное общеобразовательное учреждение "центр образования "перспектива"') != -1 or \
                    school.find('роговат') != -1 or \
                    school.find('средняя общеобразовательная школа №16') != -1 or \
                    school.find('князя александра невского') != -1 or \
                    school.find('старого оскола') != -1 or \
                    school.find('средняя общеобразовательная школа № 14') != -1 or \
                    school.find('средняя общеобразовательная школа № 12') != -1:
                mapRegionCountSchool['Старооскольский район']['school'][0] += 1
                mapRegionCountSchool['Старооскольский район']['school'].append(school)
                mapRegionCountSchool['Старооскольский район']['name'].append(name)
                mapRegionCountSchool['Старооскольский район']['date'].append(date)
                break
    # if sheet['F' + str(i)].value in mapRegionCountSchool.keys():
    #     mapCountSchool[sheet['F' + str(i)].value] += 1
    # else:
    #     mapCountSchool[sheet['F' + str(i)].value] = 1
    # print(sheet['E' + str(i)].value, sheet['F' + str(i)].value)

from collections import Counter

array = mapRegionCountSchool['Яковлевский район']['school']
print(array)
сounter = Counter(array[1:])
print((сounter.items()))
fig = plt.figure(figsize=(32, 10))
ax = fig.add_subplot()
print(mapRegionCountSchool.keys())
# x = [f'{key[:30]}\n{key[30:60]}\n{key[60:90]}' for key in getNamesSchool(сounter.keys())]
# y = [value for value in сounter.values()]
x = [f'{key.split(" ")[0][:-4]}.\n{key.split(" ")[1]}' for key in mapRegionCountSchool.keys()]
y = []
y1 = []
# for key in mapRegionCountSchool.keys():
#     y.append(len(Counter(mapRegionCountSchool[key]['school'][1:])))
# countFemale, countMale = [], []
# for key in mapRegionCountSchool.keys():
#     male, female = sortedName(mapRegionCountSchool[key]['name'])
#     countFemale.append(len(female))
#     countMale.append(len(male))
avarage_date = []
for key in mapRegionCountSchool.keys():
    age_list = [(datetime.now()-getDate(age)).days/365 for age in mapRegionCountSchool[key]['date']]
    avarage_date.append(sum(age_list)/len(age_list))
print(avarage_date)
# y = [len(Counter(value[1:])) for value in mapRegionCountSchool.values()]
# y1 = [value[0] for value in mapRegionCountSchool.values()['school']]
for key in mapRegionCountSchool.keys():
    y1.append(mapRegionCountSchool[key]['school'][0])
ax.bar(x, avarage_date, label='Средний возраст', width=0.4)
# ax.bar([i + 0.4 for i in range(len(x))], countMale, label='количество мальчиков', width=0.4)
# ax.barh(x, y1, label='количество учеников', color='r')
ax.set_ylabel('Возраст')
ax.legend()
ax.grid()
plt.title('Средний возраст учеников по районам Белгородской области')
plt.show()
