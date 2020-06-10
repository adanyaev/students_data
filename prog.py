import datetime
import xlrd
import matplotlib.pyplot as plt
import pylab
import matplotlib.dates
import numpy as np
from collections import OrderedDict
import matplotlib.ticker as ticker
from matplotlib.figure import Figure
#D:\19ПИ-3.xlsx - это путь к файлу
class STUDENT(object):
    def __init__(self, name='', group='', results=[]):
        self.name = name
        self.group = group
        self.results = results
def Current_student_grades(*file_names):
    students = []
    for file_num in range(len(file_names)):
        file_name = file_names[file_num]
        try:
            xlrd.open_workbook(file_name, formatting_info=False)
        except FileNotFoundError:
            #print(str("The file is not found: " + str(file_name)))
            continue
        rb = xlrd.open_workbook(file_name, formatting_info=False)
        sheet = rb.sheet_by_index(0)            #выбираем активный лист
        row_number = sheet.nrows
        col_number = sheet.ncols
        if row_number > 0:
            group = sheet.cell(rowx=0, colx=col_number-1).value
            subject = sheet.cell(rowx=1, colx=col_number-1).value
            i = 2
            kolvo = 0
            while sheet.cell(rowx=0, colx=i).value != 'Средний балл' and sheet.cell(rowx=1, colx=i).value != 'Middle point':
                kolvo += 1
                i += 1
            k = 0
            for i in range(1, row_number):
                #array = [0]*kolvo
                j = 1
                k = 0
                dict_ = {}
                while sheet.cell(rowx=0, colx=j + 1).value != 'Средний балл' and sheet.cell(rowx=1,
                                                                                            colx=j).value != 'Middle point':
                    cell = sheet.cell(rowx=0, colx=j + 1)
                    j += 1
                    if cell.ctype == 3:
                        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(cell.value, rb.datemode)
                        value = str(datetime.date(year, month, day))
                    else:
                        value = cell.value
                    mark = str(sheet.cell(rowx=i, colx=j).value)
                    if mark == '':
                        continue
                    else:
                        dict_.update({str(value): mark})
               # dict_.update({sheet.cell(rowx=0, colx=kolvo + 2).value: sheet.cell(rowx=i, colx=kolvo + 2).value})
               # dict_.update({sheet.cell(rowx=0, colx=kolvo + 3).value: sheet.cell(rowx=i, colx=kolvo + 3).value})

                # отсортирует по возрастанию ключей словаря
                OrderedDict(sorted(dict_.items(), key=lambda t: t[0]))
                flag = 0
                if students != []:
                    for m in range(len(students)):
                        if students[m].name == sheet.cell(rowx=i, colx=1).value:
                            flag = 1
                            students[m].results.append(subject)
                            students[m].results.append(dict_)
                if flag == 0:
                    results = []
                    results.append(subject)
                    results.append(dict_)
                    Student = STUDENT(sheet.cell(rowx=i, colx=1).value, group, results)
                    students.append(Student)
        else:
            #print('Empty file detected: ' + file_name)
            continue
    return students


def Graph_Of_Current_Grades(students, number_of_st):
    fig = plt.figure(figsize=(10, 6))         #размер окна задается тут
    ax = fig.add_subplot()
    plt.ymin = 0
    plt.ymax = 10
    plt.style.use('seaborn-bright')
    subjects = []
    dates = []
    all_dates = []
    grades = []
    all_grades = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    nums = []

    for i in range(len(students[number_of_st].results)):
        if i % 2 == 0:
            subjects.append(students[number_of_st].results[i])
            nums.append(i)
    for i in range(len(nums)):
        num = nums[i]                  #i - номер предмета
        grades.append(list(students[number_of_st].results[num+1].values()))
        dates.append(list(students[number_of_st].results[num+1].keys()))
        for k in range(len(dates[i])):
                dates[i][k] = datetime.datetime.strptime(dates[i][k], '%Y-%m-%d')
                dates[i][k] = dates[i][k].strftime('%d.%m')
                all_dates.append(dates[i][k])
        for j in range(len(grades[i])):
            grades[i][j] = int(float(grades[i][j]))
    all_dates = list(set(all_dates))
    all_dates.sort(key=lambda date: datetime.datetime.strptime(date, '%d.%m'))
    grades_gr = []
    grades_gr_ = []
    for i in range(len(dates)):
        k = 0
        for j in range(len(all_dates)):
            if all_dates[j] == dates[i][k]:
                grades_gr_.append(grades[i][k])
                if k < len(dates[i])-1:
                    k += 1
            else:
                grades_gr_.append(None)
        grades_gr.append(list(grades_gr_))
        grades_gr_.clear()

    for i in range(len(all_dates)):
        all_dates[i] = datetime.datetime.strptime(all_dates[i], '%d.%m')

    xdata_float = matplotlib.dates.date2num(all_dates)
    pylab.subplot().xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%d.%m"))
    all_dates = np.array(all_dates)

    for i in range(len(grades_gr)):
        series1 = np.array(grades_gr[i]).astype(np.double)
        mask = np.isfinite(series1)
        plt.plot(xdata_float[mask], series1[mask], label=subjects[i], marker='o')

    ax.grid(c='#BFEFEF', ls='-', lw=1)
    ax.set_ylim([0, 11])
    ax.set_title("Grades of " + str(students[number_of_st].name))
    ax.set_xlabel('Dates')
    ax.set_ylabel('Marks')
    plt.yticks(all_grades)
    plt.xticks(all_dates)
    #  Настраиваем вид основных тиков:
    ax.tick_params(axis='x',  # Применяем параметры к оси x
                   which='major',  # Применяем параметры к основным делениям
                   pad=3,  # Расстояние между черточкой и ее подписью
                   labelsize=8,  # Размер подписи
                   bottom=True,  # Рисуем метки снизу
                   top=False,  # сверху
                   left=True,  # слева
                   right=False,  # и справа
                   labelbottom=True,  # Рисуем подписи снизу
                   labeltop=False,  # сверху
                   labelleft=True,  # слева
                   labelright=False,  # и справа
                   labelrotation=35)  # Поворот подписей

    plt.legend()
    plt.show()



arr = []
mytuple = ("D:\\19ПИ-3.xlsx", "D:\\Empty.xlsx", "D:\\New.xls", "D:\\vsc.xlsx")     #пример входных данных
arr = Current_student_grades(*mytuple)                                 #вызов моей функции, передаем кортеж в качестве аргумента
for i in range(len(arr)):                                              #вывод результата работы функции
   print(str(arr[i].name) + " " + str(arr[i].group) + " ")
   print(arr[i].results)
Graph_Of_Current_Grades(students=arr, number_of_st=4)


