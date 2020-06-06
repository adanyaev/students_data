import xlrd
import matplotlib.pyplot as plt
import math


def start_data_parsing(file_path):      # Принимает путь к файлу (только к одному), который нужно считать
    # "D:\\Python\\PY PROJECTS\\19pi.xlsx"
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


def start_data_drawPiechart():              # Распределение конкурсных баллов по всем предметам
    values = sorted(main_data["Сумма конкурсных баллов"], key=int)
    colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
    slices = []
    if max(values) < 320:
        limits = [200, 250, 290, 320]
        labels = ["0-200", "201-250", "251-290", "291-320"]
    else:
        limits = [200, 280, 350, 420]
        labels = ["0-190", "191-280", "281-350", "351-420"]
    len_ = len(values)
    j = 0
    for i in range(0, 4):
        counter = 0
        while j < len_ and values[j] <= limits[i]:
            counter += 1
            j += 1
        slices.append(counter)
    total = sum(slices)
    plt.title("Распределение конкурсных баллов")
    plt.pie(slices, shadow=True, colors=colors, wedgeprops={"edgecolor": "black"},  autopct=lambda p: '{:.0f}'.format(p * total / 100), startangle=180)
    plt.legend(labels, loc="best")
    plt.axis('equal')
    plt.tight_layout()
    plt.show()


def start_data_drawHistogramm(subject):     # Принимает название предмета (Информатика и ИКТ, Математика или Русский язык)
    plt.xlabel("Баллы ЕГЭ")
    plt.ylabel("Количество сдавших")
    plt.title("Распределение баллов дисциплины: " + subject)
    axes = main_data[subject]
    axes.sort(key=int)
    for i in range(len(axes)):
        axes[i] -= 0.01
    bins = [20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105, 110]
    plt.hist(axes, bins=bins, edgecolor='black')
    plt.xlim(min(axes)-10, 110)
    plt.show()


def start_data_topFive(subject):            # Принимает название предмета, возвращает словарь
    students_list.sort(reverse=True, key=lambda n: n[subject])
    new_dict = {'ФИО': [], subject: []}
    for i in range(0, 5):
        new_dict['ФИО'].append(students_list[i]['ФИО'])
        new_dict[subject].append(students_list[i][subject])
    return new_dict

main_data, students_list = start_data_parsing("D:\\Python\\PY PROJECTS\\19pi.xlsx")

#   new_dict = start_data_topFive('Математика')
#   new_dict содержит два ключа: ФИО и название предмета, значениями являются списки из имен и оценок, порядок сохраняется
#   Пример получения этих данных
#for i in new_dict:
#   print(new_dict[i])




