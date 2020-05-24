import xlrd

#"D:\\Python\\PY PROJECTS\\19pi.xlsx"
file_path = input()
book = xlrd.open_workbook(file_path)
sheet = book.sheet_by_index(book.nsheets - 1)
data = {}
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

#---------------------------------- Main code --------------------------------#

startrow, startcol = findStartInTable()

# Get data from excel
for i in range(startcol, colnum):
    tmp_list = []
    tmp_name = sheet.cell_value(startrow, i)
    for j in range(startrow + 1, rownum):
        tmp_value = sheet.cell_value(j, i)
        tmp_list.append(tmp_value)
    data.update({tmp_name : tmp_list})

print(data)


