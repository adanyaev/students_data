import xlrd
import numpy as np
import matplotlib.pyplot as plt


def Pars(name_of_file):
    class Student(object):
        as10 = None
        as5 = None
        studbil = None
        name = None
        number = None

    # Shkala5 = {
    #     'неуд': 1,
    #     'уд': 2,
    #     'хор': 3,
    #     'отл': 4
    # }

    fil = xlrd.open_workbook(name_of_file, formatting_info=True)
    sheet = fil.sheet_by_index(0)

    rownum = 0

    for rr in range(sheet.nrows):
        for cc in range(sheet.ncols):
            if sheet.cell(rr, cc).value == 'Номер студенческого билета':
                rownum = rr + 2

    rr = rownum
    rend = 0
    while (True):
        if sheet.cell(rr, 2).value == '':
            rend = rr
            break
        rr += 1
    rend -= 1

    students = []

    sl = rend - rownum

    while rownum <= rend:
        j = Student()

        j.number = int(sheet.cell(rownum, 1).value)
        j.studbil = sheet.cell(rownum, 2).value
        j.name = sheet.cell(rownum, 3).value
        j.as10 = int(sheet.cell(rownum, 9).value)
        j.as5 = sheet.cell(rownum, 10).value

        students.append(j)
        rownum += 1

    # faculty = sheet.cell(1, 1).value
    # ed_progr = sheet.cell(2, 1).value
    # course = int(sheet.cell(6, 3).value)
    # module = sheet.cell(6, 6).value
    # year = sheet.cell(6, 10).value
    # group = sheet.cell(7, 3).value
    # discipline = sheet.cell(10, 3).value
    # teacher = sheet.cell(11, 6).value
    # date = sheet.cell(12, 6).value
    # ed_progr = ed_progr[28:]
    # module = module[0:1]

    # print(faculty)
    # print('Образовательная программа:', ed_progr,
    #       '\nКурс -', course,
    #       '\nМодуль -', module,
    #       '\nУчебный год -', year,
    #       '\nГруппа -', group,
    #       '\nДисциплина -', discipline,
    #       '\nПреподаватель -', teacher,
    #       '\nУченики:')

    scores = []
    allsc = []
    k = 0
    while k < 10:
        scores.append(0)
        k += 1

    n = 0
    while n <= sl:
        # print(students[n].number, students[n].studbil, students[n].name, students[n].as10,
        #       students[n].as5)
        scores[students[n].as10 - 1] += 1
        allsc.append(students[n].as10)
        n += 1
    # print()

    # print(scores)

    # print(
    #     'Введите [1 / 2 / 3 / 4 / 5] для сортировки учеников по ' +
    #     '\n[номеру / студ. билету / ФИО / 10-бальной шкале / 5-бальной '
    #     'шкале]')
    # print(
    #     'Если хотите увидеть' +
    #     '\n[факультет / обр.программу / курс / модуль / учебный год / группу / дисциплину / преподавателя],' +
    #     '\nВведите [f / edpr / course / module / year / gr / disc / t]' +
    #     '\nЛибо 0 для выхода')
    # d = input()

    students.sort(key=lambda Student: Student.number)

    # while d != '0':
    #     if d == '1':
    #         students.sort(key=lambda Student: Student.number)
    #     elif d == '2':
    #         students.sort(key=lambda Student: Student.studbil)
    #     elif d == '3':
    #         students.sort(key=lambda Student: Student.name)
    #     elif d == '4':
    #         students.sort(key=lambda Student: Student.as10, reverse=True)
    #     elif d == '5':
    #         students.sort(key=lambda Student: Shkala5[Student.as5], reverse=True)
    #     if (d <= '5') and (d >= '1'):
    #         n = 0
    #         while n <= sl:
    #             print(students[n].number, students[n].studbil, students[n].name, students[n].as10,
    #                   students[n].as5)
    #             n += 1
    #     if d == 'f':
    #         print(faculty)
    #     elif d == 'edpr':
    #         print(ed_progr)
    #     elif d == 'course':
    #         print(course)
    #     elif d == 'module':
    #         print(module)
    #     elif d == 'year':
    #         print(year)
    #     elif d == 'gr':
    #         print(group)
    #     elif d == 'disc':
    #         print(discipline)
    #     elif d == 't':
    #         print(teacher)
    #     print(
    #         '\nВведите [1 / 2 / 3 / 4 / 5] для сортировки учеников по ' +
    #         '\n[номеру / студ. билету / ФИО / 10-бальной шкале / 5-бальной '
    #         'шкале]')
    #     print(
    #         'Если хотите увидеть' +
    #         '\n[факультет / обр.программу / курс / модуль / учебный год / группу / дисциплину / преподавателя],' +
    #         '\nВведите [f / edpr / course / module / year / gr / disc / t]' +
    #         '\nЛибо 0 для выхода')
    #     d = input()
    return scores, sl, allsc


def Gist():
    plt.title = "Распределение количчества оценок"
    plt.xlabel("Оценки по десятибальной шкале")
    plt.ylabel("Количество учеников, получивших соответсвующую оценку")
    i = 0
    allsccopy = allsc
    allsccopy.sort()
    x = np.arange(0.5, 11.5, 1)
    plt.hist(allsccopy, bins=x, edgecolor="black",rwidth=0.95, color="red")
    plt.xlim(0.5, 10.5)
    plt.ylim(0, max(scrs) + 0.5)
    plt.show()


name_of_file = "19pi_example.xls"
scrs, numofst, allsc = Pars(name_of_file)
numofst += 1
# print(scrs)
# print(allsc)
# print(numofst)
Gist()
