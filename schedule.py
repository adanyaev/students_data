import xlrd
import matplotlib.pyplot as plt
import numpy as np


def get_schedule(paths):
    tables = []
    for path in paths:
        book = xlrd.open_workbook(path, formatting_info=True)
        sheet = book.sheet_by_index(0)
        module = {}

        i = 0
        while sheet.cell_value(i, 0).find("Факультет") == -1:
            i += 1
        txt = sheet.cell_value(i, 0)
        pos1 = txt.find('.')
        pos2 = txt.find('.', pos1 + 1)
        start = int(txt[pos1 - 2:pos1]) + (int(txt[pos1 + 1:pos1 + 3]) - 1) * 30
        end = int(txt[pos2 - 2:pos2]) + (int(txt[pos2 + 1:pos2 + 3]) - 1) * 30
        module['module_length'] = end - start

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

        is_merged = [[None] * sheet.ncols for i in range(sheet.nrows)]
        for crange in sheet.merged_cells:
            rlo, rhi, clo, chi = crange
            for row in range(rlo, rhi):
                for col in range(clo, chi):
                    is_merged[row][col] = (rlo, clo)

        table = []
        column = 2
        while sheet.cell_value(start + 1, column) != "":
            group = {"name": sheet.cell_value(start + 1, column)}
            schedule = []
            for i in range(start + 2, start + 2 + (6 * 6), 6):
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
                for j in range(i, i + 6):
                    lesson = {}
                    lesson["subject"] = sheet.cell_value(j, column).strip()
                    lesson["cab"] = sheet.cell_value(j, column + 1)
                    if lesson["subject"] == "" and is_merged[j][column] is None:
                        lesson["subject"] = "Нет пары"
                        lesson["cab"] = "Нет пары"
                    elif is_merged[j][column]:
                        lesson["subject"] = sheet.cell_value(is_merged[j][column][0], is_merged[j][column][1])
                        c = is_merged[j][column][1]
                        while is_merged[j][column] == is_merged[is_merged[j][column][0]][c]:
                            c += 1
                        lesson["cab"] = sheet.cell_value(is_merged[j][column][0], c)
                    lesson["subject"] = lesson["subject"].strip()
                    if is_merged[j][column] is None:
                        xfx = sheet.cell_xf_index(j, column)
                    else:
                        xfx = sheet.cell_xf_index(is_merged[j][column][0], is_merged[j][column][1])
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
        module['table'] = table
        tables.append(module)
    return tables


def print_chart(tables, module, group):
    schedule = tables[module]['table'][group]['schedule']
    module_length = tables[module]['module_length']

    working_hours = {'lower_week': [0 for i in range(6)],
                     'upper_week': [0 for i in range(6)]}

    windows = {'lower_week': [0 for i in range(6)],
               'upper_week': [0 for i in range(6)]}

    for i in range(6):
        first = -1
        last = 0
        for k in range(6):
            if schedule[i]['lower_week'][k]['subject'] != "Нет пары":
                working_hours['lower_week'][i] = working_hours['lower_week'][i] + 1
                last = k
                if first == -1:
                    first = k
        for k in range(first, last):
            if schedule[i]['lower_week'][k]['subject'] == "Нет пары":
                windows['lower_week'][i] = windows['lower_week'][i] + 1
        first = -1
        last = 0
        for k in range(6):
            if schedule[i]['upper_week'][k]['subject'] != "Нет пары":
                working_hours['upper_week'][i] = working_hours['upper_week'][i] + 1
                last = k
                if first == -1:
                    first = k
        for k in range(first, last):
            if schedule[i]['upper_week'][k]['subject'] == "Нет пары":
                windows['upper_week'][i] = windows['upper_week'][i] + 1

    total_work = [0 for i in range(6)]
    total_win = [0 for i in range(6)]

    k = module_length // 14
    for j in range(6):
        total_work[j] = (working_hours['upper_week'][j] + working_hours['lower_week'][j]) * k
        total_win[j] = (windows['upper_week'][j] + windows['lower_week'][j]) * k
    k = module_length % 14
    if k >= 7:
        for j in range(6):
            total_work[j] = total_work[j] + working_hours['upper_week'][j]
            total_win[j] = total_win[j] + windows['upper_week'][j]
        for j in range(k % 7):
            total_work[j] = total_work[j] + working_hours['lower_week'][j]
            total_win[j] = total_win[j] + windows['lower_week'][j]
    else:
        for j in range(k):
            total_work[j] = total_work[j] + working_hours['upper_week'][j]
            total_win[j] = total_win[j] + windows['upper_week'][j]

    days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
    data = [total_work, total_win]
    x = np.arange(6)
    fig, ax = plt.subplots()
    ax.bar(x + 0.25, data[0], color='b', width=0.25, label="Учебные часы")
    ax.bar(x + 0.5, data[1], color='g', width=0.25, label="Окна")
    ax.set_xlabel('Дни недели')
    ax.set_ylabel('Академические часы')
    ax.set_xticks(x + 0.375)
    ax.set_xticklabels(days)
    ax.legend()

    plt.show()

    return


if __name__ == "__main__":
    path1 = r'D:\other\code\PyCode\excel\module2.1.xls'
    path2 = r'D:\other\code\PyCode\excel\module3.1.xls'
    paths = [path1, path2]
    tables = get_schedule(paths)
    table = tables[0]['table']
    print(table)
    print(table[1]["name"])
    print(table[1]["schedule"][0]["name"])
    print(table[1]["schedule"][0]["place"])
    print(table[1]["schedule"][0]["lower_week"][0])
    print(table[1]["schedule"][0]["upper_week"][0])
    print(tables[1]['module_length'])
    print_chart(tables, 1, 0)
