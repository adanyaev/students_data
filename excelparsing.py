import xlrd
import matplotlib.pyplot as plt


def start_data_parsing(file_path):      # Принимает путь к файлу (только к одному), который нужно считать
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

    startrow, startcol = findStartInTable()         # Находим позицию, с которой начинаем парсить

    for i in range(startcol, colnum):
        tmp_list = []
        tmp_name = sheet.cell_value(startrow, i)
        for j in range(startrow + 1, rownum):
            tmp_value = sheet.cell_value(j, i)
            tmp_list.append(tmp_value)
        data.update({tmp_name: tmp_list})

    subject1 = sheet.cell_value(startrow, startcol + 8)
    subject2 = sheet.cell_value(startrow, startcol + 9)
    subject3 = sheet.cell_value(startrow, startcol + 10)
    subject4 = sheet.cell_value(startrow, startcol + 11)
    students_num = rownum - startrow - 1
    for i in range(students_num):
        tmp_student = {'ФИО': data['Фамилия, имя, отчество'][i]}
        if data[subject1][i] != "":
            mark1 = int(data[subject1][i])
        else:
            mark1 = 0
        if data[subject2][i] != "":
            mark2 = int(data[subject2][i])
        else:
            mark2 = 0
        if data[subject3][i] != "":
            mark3 = int(data[subject3][i])
        else:
            mark3 = 0
        if data[subject4][i] != "":
            mark4 = int(data[subject4][i])
        else:
            mark4 = 0
        tmp_student.update({subject1: mark1})
        tmp_student.update({subject2: mark2})
        tmp_student.update({subject3: mark3})
        tmp_student.update({subject4: mark4})
        students.append(tmp_student)

        subject_list = [subject1, subject2, subject3]
        if data[subject4][0] != "":
            subject_list.append(subject4)

    return data, students, subject_list


def start_data_drawPiechart(main_data):              # Распределение конкурсных баллов по всем предметам
    fig = plt.figure(figsize=(10, 6))
    values = sorted(main_data["Сумма конкурсных баллов"], key=int)
    colors = ['red', 'lightcoral', 'yellowgreen', 'lightskyblue', 'gold']
    slices = []
    if max(values) < 320:
        limits = [200, 230, 260, 290, 320]
        labels = ["0-200", "201-230", "231-260", "261-290", "291-320"]
    else:
        limits = [200, 255, 310, 365, 420]
        labels = ["0-200", "201-255", "256-310", "311-365", "366-420"]
    len_ = len(values)
    j = 0
    for i in range(0, 5):
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
    #plt.show()
    return fig


def start_data_drawHistogramm(subject, main_data):     # Принимает название предмета (Информатика и ИКТ, Математика или Русский язык)
    fig = plt.figure()
    plt.xlabel("Баллы ЕГЭ")
    plt.ylabel("Количество сдавших")
    plt.title("Распределение баллов дисциплины: " + subject)
    axes = list(filter(lambda x: isinstance(x, (int, float)), main_data[subject]))
    axes.sort()
    for i in range(len(axes)):
        axes[i] -= 0.01
    bins = [20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105, 110]
    plt.hist(axes, bins=bins, edgecolor='black')
    plt.xlim(min(axes)-10, 110)
    #plt.show()
    return fig


def start_data_topTen(subject):            # Принимает название предмета
    print("Топ", str(top_num),  "людей по дисциплине:", subject)
    students_list.sort(reverse=True, key=lambda n: n[subject])
    for i in range(top_num):
        print(students_list[i]['ФИО'], "-", str(students_list[i][subject]))
    return


if __name__ == '__main__':
    top_num = 10
    main_data, students_list, subjects = start_data_parsing()
   # print(subjects)
   # print(main_data)
    print(subjects)
   # start_data_drawHistogramm('Информатика и ИКТ')
    #start_data_drawPiechart()
    #start_data_topTen('Математика')

