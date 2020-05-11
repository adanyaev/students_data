import xlrd

book = xlrd.open_workbook(r'D:\other\code\PyCode\excel\module3.1.xls', formatting_info=True)

sheet = book.sheet_by_index(0)
places = []
i = 0
while sheet.cell_value(i, 0).find("Корпус") == -1:
	i += 1
start = i
i = 0
while i < sheet.ncols:
	if sheet.cell_value(start, i) == "":
		i += 1
		continue
	xfx = sheet.cell_xf_index(start, i)
	xf = book.xf_list[xfx]
	bgx = xf.background.pattern_colour_index
	pattern_colour = book.colour_map[bgx]
	if pattern_colour is None:
		pattern_colour = (255, 255, 255)
	places.append({"name": sheet.cell_value(start, i), "color": pattern_colour})
	i += 1

isMerged = [[False] * sheet.ncols for i in range(sheet.nrows)]
for crange in sheet.merged_cells:
    rlo, rhi, clo, chi = crange
    for row in range(rlo, rhi):
        for col in range(clo, chi):
        	isMerged[row][col] = (rlo, clo)

table = []
column = 2
while sheet.cell_value(start+1, column) != "":
	group = {"name": sheet.cell_value(start+1, column)}
	schedule = []
	for i in range (start + 2 , start + 2 + (6*6), 6):
		day = {"name": sheet.cell_value(i, 0)}
		xfx = sheet.cell_xf_index(i, column)
		xf = book.xf_list[xfx]
		bgx = xf.background.pattern_colour_index
		pattern_colour = book.colour_map[bgx]
		if pattern_colour is None:
			pattern_colour = (255, 255, 255)
		for k in places:
			if pattern_colour == k["color"]:
				day["place"] = k["name"]
				break
		upper_week = []
		lower_week = []
		for j in range (i, i + 6):
			lesson = {}
			
			lesson["subject"] = sheet.cell_value(j, column).strip()
			lesson["cab"] = sheet.cell_value(j, column+1)
			
			if lesson["subject"] == "" and isMerged[j][column] == False:
				lesson["subject"] = "Нет пары"
				lesson["cab"] = "Нет пары"
			elif isMerged[j][column] != False:
				lesson["subject"] = sheet.cell_value(isMerged[j][column][0], isMerged[j][column][1])
				c =  isMerged[j][column][1]
				while isMerged[j][column] == isMerged[isMerged[j][column][0]][c]:
					c += 1
				lesson["cab"] = sheet.cell_value(isMerged[j][column][0], c)
			lesson["subject"] = lesson["subject"].strip()
			
			if isMerged[j][column] == False:
				xfx = sheet.cell_xf_index(j, column)
			else:
				xfx = sheet.cell_xf_index(isMerged[j][column][0], isMerged[j][column][1])
			xf = book.xf_list[xfx]
			if xf.border.diag_down == 1:
				text = lesson["subject"].split("\n", maxsplit=1)
				if len(text) > 1:
					upper_week.append({"subject": text[0].strip(), "cab": lesson["cab"]})
					lower_week.append({"subject": text[1].strip(), "cab": lesson["cab"]})
				else:
					align = xf.alignment.hor_align
					if align == 1:
						lower_week.append(lesson)
						upper_week.append({"subject": "Нет пары", "cab": "Нет пары"})
					elif align == 3:
						upper_week.append(lesson)
						lower_week.append({"subject": "Нет пары", "cab": "Нет пары"})
			else:
				lower_week.append(lesson)
				upper_week.append(lesson)
		day["upper_week"] = upper_week
		day["lower_week"] = lower_week
		schedule.append(day)
	group["schedule"] = schedule
	table.append(group)
	column += 2

print (table)
print (table[1]["name"])
print (table[1]["schedule"][0]["name"])
print (table[1]["schedule"][0]["place"])
print (table[1]["schedule"][0]["lower_week"][0])
print (table[1]["schedule"][0]["upper_week"][0])




