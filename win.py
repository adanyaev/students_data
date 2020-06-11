from PyQt5 import QtCore
from PyQt5.QtWidgets import *
import sys
import subprocess
from random import choice
from PyQt5.QtWidgets import QSizePolicy, QMessageBox
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5 import NavigationToolbar2QT as NavigationToolbar
from Canvas import Graphic
import prog
from copy import deepcopy

from matplotlib.figure import Figure

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.InitUI()
        self.MakeFrames()
        self.CreateMenu()

    def CreateGraphic(self):
        self.graphic.Plot()
        if self.chart_combo.currentText() == 'Bar chart':
            print('bar')
        elif self.chart_combo.currentText() == 'Curve chart':
            print('curve')
        elif self.chart_combo.currentText() == 'Pie chart':
            print('Pie')



    def InitUI(self):
        self.setStyleSheet(open('style.css', 'r').read())
        self.setWindowTitle('title')
        self.setGeometry(500, 500, 500, 500)
        self.showMaximized()
        self.fig = Figure(figsize=(10, 10), dpi=10)
        mytuple = (r'C:\Users\79527\Desktop\test_data\19PI-3.xlsx', r'C:\Users\79527\Desktop\test_data\New.xls',
                   r'C:\Users\79527\Desktop\test_data\19PI-1-2.xlsx', r'C:\Users\79527\Desktop\test_data\19PI-3.xlsx',
                   r'C:\Users\79527\Desktop\test_data\vsc.xlsx')
       # self.graphic = Graphic(self.fig)
        arr = prog.Current_student_grades(*mytuple)  # вызов моей функции, передаем кортеж в качестве аргумента
       # for i in range(len(arr)):  # вывод результата работы функции
        #    print(str(arr[i].name) + " " + str(arr[i].group) + " ")
         #   print(arr[i].results)
        #self.fig = prog.Graph_Of_Current_Grades(students=arr, number_of_st=5)
        self.figure = Graphic(self.fig)



        self.table = QTableWidget()
        self.graphic_frame = QFrame()
        self.button_lay = QHBoxLayout()
        self.graphic_lay = QVBoxLayout()

        self.graphic_lay.addWidget(self.figure)

        self.table_frame = QFrame()
        self.table_lay = QVBoxLayout()
        self.table_frame.setLayout(self.table_lay)

        self.bar_chart_frame = QFrame()
        self.bar_chart_lay = QVBoxLayout()
        self.bar_chart_frame.setLayout(self.bar_chart_lay)

        self.pie_chart_frame = QFrame()
        self.pie_chart_lay = QVBoxLayout()
        self.pie_chart_frame.setLayout(self.pie_chart_lay)

#        self.button_lay.addWidget(self.build_graphic_btn)
 #       self.graphic_lay.addItem(self.button_lay)
  #      self.graphic_lay.addWidget(self.graphic)
   #     self.graphic_frame.setLayout(self.graphic_lay)


        self.txt_field = QTextEdit()
        self.txt_field2 = QTextEdit()
        self.txt_field.setStyleSheet(open('style.css', 'r').read())
        self.file_tree = QTreeView()
        self.manage_tabs = QTabWidget()
        self.show_info_tabs = QTabWidget()
        self.img_label = QLabel()
        self.toolbar = QToolBar()
        self.scroll_info = QScrollArea()
        self.scroll_info.setWidget(self.show_info_tabs)
        self.scroll_info.setWidgetResizable(True)
        self.group_box = QGroupBox('MainWindow')
        self.setCentralWidget(self.group_box)

        self.show_info_tabs.addTab(self.graphic_frame, 'CurveChart')
        self.show_info_tabs.addTab(self.bar_chart_frame, 'BarChart')
        self.show_info_tabs.addTab(self.pie_chart_frame, 'PieChart')
        self.show_info_tabs.addTab(self.table_frame, 'Table')

        self.manage_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.manage_tabs.setAutoFillBackground(True)
        self.show_info_tabs.setStyleSheet(open('style.css', 'r').read())
        self.process_btn = QPushButton('Process')
        self.process_btn.clicked.connect(self.ImportIntoTable)
        self.build_graphic_btn = QPushButton('Create graphic')
        self.build_graphic_btn.clicked.connect(self.CreateCurveChart)

        self.build_bar_chart_btn = QPushButton('Create bar chart')
        self.build_bar_chart_btn.clicked.connect(self.CreateBarChart)
        self.bar_chart_lay.addWidget(self.build_bar_chart_btn)

        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(['Ведомости', 'Данные абитуриентов', 'Оценки', 'Расписание'])
        self.chart_combo = QComboBox()
        self.chart_combo.addItems(['Curve chart', 'Pie chart', 'Bar chart'])

        self.student_combo = QComboBox()

        self.button_lay.addWidget(self.build_graphic_btn)
        self.button_lay.addWidget(self.chart_combo)
        self.button_lay.addWidget(self.student_combo)
        self.lay = QHBoxLayout()
        self.graphic_lay.addItem(self.button_lay)
        #self.graphic_lay.addWidget(self.graphic)
        self.graphic_frame.setLayout(self.graphic_lay)
        #self.graphic_frame.setLayout(self.button_lay)
        self.table_lay.addWidget(self.table)

    #        self.txt_field.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
#        self.txt_field.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)


    def ImportIntoTable(self):
        if self.file_type_combo.currentText() == 'Ведомости':
            self.data = {}
            self.table.setRowCount(100)
            self.table.setColumnCount(100)
            self.table.horizontalHeader().setVisible(True)
            self.table.verticalHeader().setVisible(True)
            self.table.verticalHeader().sectionResizeMode(QHeaderView.Fixed)
          #  self.table.resizeColumnsToContents()

            chosen_files = self.file_tree.selectionModel().selectedIndexes()
            self.all_chosen_files = set()
            files = []
            for i in chosen_files:
                self.all_chosen_files.add(self.file_model.fileInfo(i).absoluteFilePath())

            for i in self.all_chosen_files:
                files.append(i)
            files = tuple(files)
            cnt = 0
            counter = 0
            mark_cnt = 2
            main_data = []
            for student in prog.Current_student_grades(*files):
                self.table.setItem(cnt, 0, QTableWidgetItem(str(student.name)))
                self.table.setSpan(cnt, 0, 3, 1)
                self.student_combo.addItem(str(student.name))

                main_data.append(student.group)
                main_data.append(student.results)
                self.data[deepcopy(student.name)] = main_data
                main_data = []
                print(self.data)

                for i in range(len(student.results)):
                    if type(student.results[i]) == str:
                        self.table.setItem(counter, 1, QTableWidgetItem(str(student.results[i])))
                    else:
                        for j in student.results[i]:
                            self.table.setItem(counter - 1, mark_cnt, QTableWidgetItem(str(j) + '/' + str(student.results[i][j])))
                            mark_cnt += 1
                    mark_cnt = 2
                    counter += 1
                cnt += 3
                counter = cnt
        else:
            chosen_files = self.file_tree.selectionModel().selectedIndexes()
            self.all_chosen_files = set()
            for i in chosen_files:
                self.all_chosen_files.add(self.file_model.fileInfo(i).absoluteFilePath())

    def CreateCurveChart(self):
        if self.file_type_combo.currentText() == 'Ведомости':
            mytuple = self.all_chosen_files
            arr = prog.Current_student_grades(*mytuple)  # вызов моей функции, передаем кортеж в качестве аргумента
            self.fig = prog.Graph_Of_Current_Grades(students=arr, number_of_st=self.student_combo.currentIndex())
            self.graphic_lay.removeWidget(self.figure)
            self.figure = Graphic(self.fig)
            self.graphic_lay.addWidget(self.figure)
        else:
            pass

    def CreateBarChart(self):
        if self.file_type_combo.currentText() == 'Оценки':
            for name in self.all_chosen_files:
                name_of_file = name #r'C:\Users\79527\Desktop\19pi_example.xls'
                scrs, as10 = Pars(name_of_file)
                self.bar_chart = Gist(scrs, as10)
                try:
                    self.bar_chart_lay.removeWidget(self.bar)
                except:
                    self.bar = Graphic(self.bar_chart)
                    self.bar_chart_lay.addWidget(self.bar)
                finally:
                    self.bar_chart_lay.addWidget(self.bar)
        else:
            pass


    def MakeFrames(self):
        form_frame = QFrame()
        form_frame.setFrameShape(QFrame.StyledPanel)
        form_lay = QVBoxLayout()
        inside_lay = QHBoxLayout()

        form_lay.addWidget(self.manage_tabs)
        form_lay.addItem(inside_lay)
        inside_lay.addWidget(self.process_btn)
        inside_lay.addWidget(self.file_type_combo)
        form_frame.setLayout(form_lay)
        ver_frame = QFrame()
        ver_frame.setFrameShape(QFrame.StyledPanel)

        ver_box = QVBoxLayout()
        ver_box.setContentsMargins(25, 20, 25, 25)
        ver_box.addWidget(self.scroll_info)
        ver_frame.setLayout(ver_box)

        horizontal_splitter = QSplitter(QtCore.Qt.Horizontal)
        horizontal_splitter.addWidget(form_frame)
        horizontal_splitter.addWidget(ver_frame)

        vertical_splitter = QSplitter(QtCore.Qt.Vertical)
        vertical_splitter.addWidget(horizontal_splitter)

        vbox = QVBoxLayout()
        vbox.addWidget(vertical_splitter)
        ver_frame.setStyleSheet(open('style.css', 'r').read())
        form_frame.setStyleSheet(open('style.css', 'r').read())
        self.group_box.setLayout(vbox)

    def CreateMenu(self):
        menubar = self.menuBar()
        menubar.setStyleSheet(open('style.css', 'r').read())
        file_menu = menubar.addMenu('File')
        import_project = QAction('New', self)
        import_project.setShortcut('Ctrl+Q')

        menubar.setStyleSheet(open('style.css', 'r').read())
        import_project.triggered.connect(self.CreateFileTree)
        file_menu.addAction(import_project)

        save_curve_chart = QAction('Save curve chart', self)
        save_curve_chart.triggered.connect(self.SaveCurveChart)

        save_bar_chart = QAction('Save bar chart', self)
        save_bar_chart.triggered.connect(self.SaveBarChart)

        file_menu.addAction(save_curve_chart)
        file_menu.addAction(save_bar_chart)


        edit = menubar.addMenu('Edit')
        view = menubar.addMenu('View')
        navigate = menubar.addMenu('Navigate')

    def SaveCurveChart(self):
        self.fig.savefig(str(self.student_combo.currentText()))

    def SaveBarChart(self):
        self.bar_chart.savefig('barchart')

    def CreateFileTree(self):
        options = QFileDialog.DontResolveSymlinks
        dir_cur = QtCore.QDir.currentPath()
        selected_directory = QFileDialog.getExistingDirectory(None, 'Find Files', dir_cur, options)
        self.file_tree = QTreeView()
        self.file_model = QFileSystemModel()
        self.file_model.setReadOnly(False)
        self.file_model.setRootPath(selected_directory)
        self.file_tree.setModel(self.file_model)
        self.file_tree.setAnimated(True)
        self.file_tree.setRootIndex(self.file_model.index(selected_directory))
        self.manage_tabs.addTab(self.file_tree, 'Project')
        self.file_tree.setSelectionMode(QAbstractItemView.MultiSelection)


    def OpenFile(self):
        chosen_files = self.file_tree.selectionModel().selectedIndexes()
        self.all_chosen_files = set()
        files = []
        for i in chosen_files:
            self.all_chosen_files.add(self.file_model.fileInfo(i).absoluteFilePath())

        for i in self.all_chosen_files:
            files.append(i)
        files = tuple(files)

        #self.data = {}
        main_data = []
        for student in prog.Current_student_grades(*files):
            main_data.append(student.group)
            main_data.append(student.results)
            self.data[student.name] = main_data
            main_data = []
        print(self.data)

        #self.txt_field.setText(info_from_file)
        #file.close()

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    app.exec_()

if __name__ == '__main__':
    main()
