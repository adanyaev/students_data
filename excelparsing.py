import xlrd

book = xlrd.open_workbook("D:\\Python\\PY PROJECTS\\19pi.xlsx")
sheet = book.sheet_by_index(book.nsheets - 1)
data = {}
students = {}
rownum = sheet.nrows
colnum = sheet.ncols

#------------------------------- Definitions --------------------------------#

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

def makeStudent(i, startcol):
    ordinal = sheet.cell_value(i, startcol)
    if ordinal == "":
        return None
    name = str(sheet.cell_value(i, startcol + 2))
    info = []
    info.append(sheet.cell_value(i,startcol + 1))
    for j in range(startcol + 3, colnum):
        info.append(sheet.cell_value(i, j))
    stmp = Student(name, ordinal, info)
    return stmp

class Student:
   def __init__(self, name, ordinal, info):
       self.name = name
       self.ordinal = ordinal
       self.info = info.copy()
   def showInfo(self):
       print("ФИО: " + self.name)
       print("Регистрационный номер: " + str(self.info[0]))
       print("Подлинник/Копия документа об образовании: " + str(self.info[1]))
       print("Право поступления без вступительных испытаний: " + str(self.info[2]))
       print("Поступление на места в рамках квоты для лиц, имеющих особое право: " + str(self.info[3]))
       print("Поступление на места в рамках квоты целевого приема: " + str(self.info[4]))
       print("Наличие согласия на зачисление: " + str(self.info[5]))
       print("Математика: " + str(self.info[6]))
       print("Информатика и ИКТ: " + str(self.info[7]))
       print("Русский язык: " + str(self.info[8]))
       print("Экзамен 4: " + str(self.info[9]))
       print("Балл за итоговое сочинение: " + str(self.info[10]))
       print("Балл за иные достижения: " + str(self.info[11]))
       print("Итоговая сумма баллов по индивидуальным достижениям: " + str(self.info[12]))
       print("Сумма конкурсных баллов: " + str(self.info[13]))
       print("Форма обучения: " + str(self.info[14]))
       print("Преимущественное право: " + str(self.info[15]))
       print("Требуется общежитие на время обучения: " + str(self.info[16]))
       print("Возврат документов: " + str(self.info[17]))


#---------------------------------- Main code --------------------------------#

startrow, startcol = findStartInTable()

# Get students information
for i in range(startrow + 1, rownum):
    stmp = makeStudent(i, startcol)
    if stmp != None:
        students.update({stmp.name : stmp})
        

# Get data from excel
for i in range(startcol, colnum):
    tmp_list = []
    tmp_name = sheet.cell_value(startrow, i)
    for j in range(startrow + 1, rownum):
        tmp_value = sheet.cell_value(j, i)
        tmp_list.append(tmp_value)
    data.update({tmp_name : tmp_list})


while 1:
    print("Enter operation ('data' to show all info, 'show' to see individual info): ")
    inp = input()
    if inp == "data":
        print(data)
    elif inp == "show":
        print("Enter student name: ")
        inp2 = input()
        print(students[inp2].showInfo())
    else:
        print("Error")


