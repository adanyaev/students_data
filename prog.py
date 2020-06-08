import datetime
import xlrd
#D:\19ПИ-3.xlsx - это путь к файлу
class STUDENT(object):
    def __init__(self, name='', group='', results=[]):
        self.name = name
        self.group = group
        self.results = results
def Current_student_grades(*file_names):
    students = []
    for file_num in range(int(len(file_names))):
        file_name = file_names[file_num]
        try:
            xlrd.open_workbook(file_name, formatting_info=False)
        except FileNotFoundError:
            print(str("The file is not found: " + str(file_name)))
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
                        value = datetime.datetime(year, month, day)
                    else:
                        value = cell.value
                    mark = str(sheet.cell(rowx=i, colx=j).value)
                    if mark == '':
                        continue
                    else:
                        dict_.update({str(value): mark})
                dict_.update({sheet.cell(rowx=0, colx=kolvo + 2).value: sheet.cell(rowx=i, colx=kolvo + 2).value})
                dict_.update({sheet.cell(rowx=0, colx=kolvo + 3).value: sheet.cell(rowx=i, colx=kolvo + 3).value})
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
            print('Empty file detected: ' + file_name)
            continue
    return students
#arr = []
#mytuple = ("D:\\19ПИ-3.xlsx", "D:\\Empty.xlsx", "D:\\New.xls", "D:\\vcs.xlsx")     #пример входных данных
#arr = Current_student_grades(*mytuple)                                 #вызов моей функции, передаем кортеж в качестве аргумента
#for i in range(len(arr)):                                              #вывод результата работы функции
#   print(str(arr[i].name) + " " + str(arr[i].group) + " ")
#   print(arr[i].results)