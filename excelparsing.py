import xlrd
import matplotlib.pyplot as plt


def parsing():
    # "D:\\Python\\PY PROJECTS\\19pi.xlsx"
    file_path = input()
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(book.nsheets - 1)
    data = {}
    students = []
    rownum = sheet.nrows
    colnum = sheet.ncols

    def findStartInTable():
        startrow = None
        startcol = None
        for i in range(rownum):
            for j in range(colnum):
                name = sheet.cell_value(i, j)
                if name == "№ п/п":
                    startrow = i
                    startcol = j
        if startrow == None or startcol == None:
            print("Error")
        return startrow, startcol

    def makeStudent(i):
        ordinal = sheet.cell_value(i, startcol)
        if ordinal == "":
            return None
        tmp_student = {}
        name = sheet.cell_value(i, startcol + 2)
        if sheet.cell_value(i, startcol + 8) != "":
            mark1 = int(sheet.cell_value(i, startcol + 8))
        else:
            mark1 = 0
        if sheet.cell_value(i, startcol + 9) != "":
            mark2 = int(sheet.cell_value(i, startcol + 9))
        else:
            mark2 = 0
        if sheet.cell_value(i, startcol + 10) != "":
            mark3 = int(sheet.cell_value(i, startcol + 10))
        else:
            mark3 = 0
        if sheet.cell_value(i, startcol + 11) != "":
            mark4 = int(sheet.cell_value(i, startcol + 11))
        else:
            mark4 = 0
        tmp_student.update({"ФИО": name})
        tmp_student.update({subject1: mark1})
        tmp_student.update({subject2: mark2})
        tmp_student.update({subject3: mark3})
        tmp_student.update({subject4: mark4})
        return tmp_student


    # Get positions
    startrow, startcol = findStartInTable()

    subject1 = sheet.cell_value(startrow, startcol + 8)
    subject2 = sheet.cell_value(startrow, startcol + 9)
    subject3 = sheet.cell_value(startrow, startcol + 10)
    subject4 = sheet.cell_value(startrow, startcol + 11)
    for i in range(startrow + 1, rownum):
        tmp = makeStudent(i)
        if tmp != None:
            students.append(tmp)

    # Get data from excel
    for i in range(startcol, colnum):
        tmp_list = []
        tmp_name = sheet.cell_value(startrow, i)
        for j in range(startrow + 1, rownum):
            tmp_value = sheet.cell_value(j, i)
            if tmp_value == "":
                continue
            tmp_list.append(tmp_value)
        data.update({tmp_name: tmp_list})

    return data, students


def drawHystogramm(subject):
    plt.xlabel("Баллы ЕГЭ")
    plt.ylabel("Количество сдавших")
    plt.title("Распределение баллов дисциплины: " + subject)
    ages = main_data[subject]
    ages.sort(key=int)
    bins = [20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105, 110]
    plt.hist(ages, bins=bins, edgecolor='black')
    plt.xlim(min(ages)-10, 110)
    plt.show()


def topFive(subject):
    students_list.sort(reverse=True, key=lambda n: n[subject])
    print("Top 5 list in subject: " + subject)
    print("----------------------------------")
    for i in range(5):
        print(students_list[i]["ФИО"] + " - " + str(students_list[i][subject]))
    print("----------------------------------\n")


main_data, students_list = parsing()
drawHystogramm("Информатика и ИКТ")

#topFive("Математика")
#topFive("Информатика и ИКТ")
#topFive("Русский язык")
#topFive("Экзамен 4")







