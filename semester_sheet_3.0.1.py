import xlrd
import numpy as np
import matplotlib.pyplot as plt


def Pars(name_of_file):
    fil = xlrd.open_workbook(name_of_file)
    sheet = fil.sheet_by_index(0)

    rownum = 0

    for rr in range(sheet.nrows):
        for cc in range(sheet.ncols):
            if sheet.cell(rr, cc).value == 'Номер студенческого билета':
                rownum = rr + 2

    rr = rownum

    while (True):
        if sheet.cell(rr, 2).value == '':
            rend = rr
            break
        rr += 1
    rend -= 1

    as10 = []

    sl = rend - rownum

    while rownum <= rend:
        as10.append(int(sheet.cell(rownum, 9).value))
        rownum += 1

    scores = []
    k = 0
    while k < 10:
        scores.append(0)
        k += 1

    n = 0
    while n <= sl:
        scores[as10[n] - 1] += 1
        n += 1

    print(as10)
    print(scores)

    return scores, as10


def Gist():
    plt.title = "Распределение количчества оценок"
    plt.xlabel("Оценки по десятибальной шкале")
    plt.ylabel("Количество учеников, получивших соответсвующую оценку")
    i = 0
    allsccopy = as10.copy()
    allsccopy.sort()
    x = np.arange(0.5, 11.5, 1)
    plt.hist(allsccopy, bins=x, edgecolor="black", rwidth=0.95, color="red")
    plt.xlim(0.5, 10.5)
    plt.ylim(0, max(scrs) + 0.5)
    plt.show()


name_of_file = input()
scrs, as10 = Pars(name_of_file)
Gist()
