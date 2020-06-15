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

    scores = []
    allsc = []
    k = 0
    while k < 10:
        scores.append(0)
        k += 1

    n = 0
    while n <= sl:
        scores[students[n].as10 - 1] += 1
        allsc.append(students[n].as10)
        n += 1

    students.sort(key=lambda Student: Student.number)

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
Gist()