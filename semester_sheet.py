import xlrd
import numpy as np
import matplotlib.pyplot as plt


def Pars(name_of_file):
    fil = xlrd.open_workbook(name_of_file)
    sheet = fil.sheet_by_index(0)

    rowbeg = 0

    for rr in range(sheet.nrows):
        for cc in range(sheet.ncols):
            if sheet.cell(rr, cc).value == 'Номер студенческого билета':
                rowbeg = rr + 2

    rr = rowbeg

    while (True):
        if sheet.cell(rr, 2).value == '':
            rend = rr
            break
        rr += 1
    rend -= 1

    as10 = []

    while rowbeg <= rend:
        as10.append(int(sheet.cell(rowbeg, 9).value))
        rowbeg += 1

    scores = []
    k = 0
    while k < 10:
        scores.append(0)
        k += 1

    k = 0
    while k < len(as10):
        scores[as10[k] - 1] += 1
        k += 1

    return scores, as10


def Gist(scrs, as10):
    fig = plt.figure(figsize=(10, 6))
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
#    plt.show()
    return fig


if __name__ == '__main__':
    pass
