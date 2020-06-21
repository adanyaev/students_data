from PyQt5 import QtCore
from PyQt5.QtWidgets import *
import sys
from PyQt5.QtWidgets import QSizePolicy, QMessageBox
from Canvas import Graphic
from prog import *
from copy import deepcopy
from semester_sheet import *
from excelparsing import *
from matplotlib.figure import Figure

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.InitUI()
        self.CreateFrames()
        self.CreateMenu()

    def InitUI(self):
        self.setGeometry(500, 500, 500, 500)
        self.showMaximized()
        self.fig = Figure(figsize=(10, 10), dpi=10)
        self.pie = Figure(figsize=(10, 10), dpi=10)
        self.figure = Graphic(self.fig)
        self.pie_chart = Graphic(self.pie)


        self.table = QTableWidget()
        self.graphic_frame = QFrame()
        self.button_lay = QHBoxLayout()
        self.graphic_lay = QVBoxLayout()

        self.table_frame = QFrame()
        self.table_lay = QVBoxLayout()
        self.table_frame.setLayout(self.table_lay)

        self.bar_chart_frame = QFrame()
        self.bar_chart_lay = QVBoxLayout()
        self.bar_btn_lay = QHBoxLayout()
        self.build_bar_chart_btn = QPushButton('Create bar chart')
        self.build_bar_chart_btn.clicked.connect(self.CreateBarChart)
        self.bar_btn_lay.addWidget(self.build_bar_chart_btn, 0, QtCore.Qt.AlignTop)
        self.bar_chart_lay.addItem(self.bar_btn_lay)
        self.bar_chart_frame.setLayout(self.bar_chart_lay)


        self.build_pie_chart_btn = QPushButton('Create pie chart')
        self.build_pie_chart_btn.clicked.connect(self.CreatePieChart)
        self.build_hist_chart_btn = QPushButton('Create histogram')
        self.build_hist_chart_btn.clicked.connect(self.CreateHist)

        self.file_idx_combo = QComboBox()
        self.pie_chart_btn_lay = QHBoxLayout()
        self.hist_chart_lay = QVBoxLayout()
        self.hist_chart_frame = QFrame()
        self.hist_chart_frame.setLayout(self.hist_chart_lay)
        self.subject_combo = QComboBox()
        self.pie_chart_btn_lay.addWidget(self.build_pie_chart_btn, 0, QtCore.Qt.AlignTop)
        self.pie_chart_btn_lay.addWidget(self.file_idx_combo, 1, QtCore.Qt.AlignTop)
        self.hist_chart_btn_lay = QHBoxLayout()
        self.hist_chart_btn_lay.addWidget(self.build_hist_chart_btn, 0, QtCore.Qt.AlignTop)
        self.hist_chart_btn_lay.addWidget(self.subject_combo, 1, QtCore.Qt.AlignTop)
        self.hist_chart_lay.addLayout(self.hist_chart_btn_lay)

        self.pie_chart_frame = QFrame()
        self.pie_chart_lay = QVBoxLayout()
        self.pie_chart_lay.addLayout(self.pie_chart_btn_lay)
        self.pie_chart_frame.setLayout(self.pie_chart_lay)

        self.schedule_frame = QFrame()
        self.schedule_lay = QVBoxLayout()
        self.schedule_button_lay = QHBoxLayout()
        self.build_schedule_btn = QPushButton('Create Schedule Chart')
        self.build_schedule_btn.clicked.connect(self.CreateScheduleChart)
        self.students_group_combo = QComboBox()
        self.module_combo = QComboBox()
        #self.schedule_lay.addWidget(self.build_schedule_btn, 0, QtCore.Qt.AlignTop)
        #self.schedule_lay.addWidget(self.schedule_combo, 1, QtCore.Qt.AlignTop)
        self.schedule_button_lay.addWidget(self.build_schedule_btn, 0, QtCore.Qt.AlignTop)
        self.schedule_button_lay.addWidget(self.students_group_combo, 1, QtCore.Qt.AlignTop)
        self.schedule_button_lay.addWidget(self.module_combo, 2, QtCore.Qt.AlignTop)
        self.schedule_lay.addItem(self.schedule_button_lay)
        self.schedule_frame.setLayout(self.schedule_lay)

        self.table.setRowCount(100)
        self.table.setColumnCount(100)
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().sectionResizeMode(QHeaderView.Fixed)

        self.txt_field = QTextEdit()
        self.txt_field2 = QTextEdit()
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
        self.show_info_tabs.addTab(self.hist_chart_frame, 'Hist')
        self.show_info_tabs.addTab(self.schedule_frame, 'Schedule')
        self.show_info_tabs.addTab(self.table_frame, 'Table')

        self.manage_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.manage_tabs.setAutoFillBackground(True)
        self.process_btn = QPushButton('Process')
        self.process_btn.clicked.connect(self.ProcessData)
        self.build_graphic_btn = QPushButton('Build curve chart')
        self.build_graphic_btn.clicked.connect(self.CreateCurveChart)

        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(['Ведомости', 'Данные абитуриентов', 'Оценки', 'Расписание'])
        self.student_combo = QComboBox()

        self.button_lay.addWidget(self.build_graphic_btn, QtCore.Qt.AlignTop)
        #self.button_lay.addWidget(self.chart_combo, QtCore.Qt.AlignTop)
        self.button_lay.addWidget(self.student_combo, QtCore.Qt.AlignTop)
        self.lay = QHBoxLayout()
        self.graphic_lay.addItem(self.button_lay)
        self.graphic_lay.addWidget(self.figure)
        #self.graphic_lay.addWidget(self.graphic)
        self.graphic_frame.setLayout(self.graphic_lay)
        #self.graphic_frame.setLayout(self.button_lay)
        self.table_lay.addWidget(self.table)


    def ProcessData(self):
        if self.file_type_combo.currentText() == 'Ведомости':
            self.data = {}
            self.GetHighlightedPaths()
            files = []
            for i in self.all_chosen_files:
                files.append(i)
            files = tuple(files)
            cnt = 0
            counter = 0
            mark_cnt = 2
            main_data = []
            for student in Current_student_grades(*files):
                self.table.setItem(cnt, 0,
                QTableWidgetItem(str(student.name)))
                self.table.setSpan(cnt, 0, 3, 1)
                self.student_combo.addItem(str(student.name))

                main_data.append(student.group)
                main_data.append(student.results)
                self.data[deepcopy(student.name)] = main_data
                main_data = []

                for i in range(len(student.results)):
                    if type(student.results[i]) == str:
                        self.table.setItem(counter, 1,
                        QTableWidgetItem(str(student.results[i])))
                    else:
                        for j in student.results[i]:
                            self.table.setItem(counter - 1, mark_cnt,
                            QTableWidgetItem(str(j) + '/' + str(student.results[i][j])))
                            mark_cnt += 1
                    mark_cnt = 2
                    counter += 1
                cnt += 3
                counter = cnt
        elif self.file_type_combo.currentText() == 'Данные абитуриентов':
            self.GetHighlightedPaths()
            self.start_data = []
            self.main_data = []
            self.students_list = []
            self.subjects = []
            for i in self.all_chosen_files:
                self.file_idx_combo.addItem(str(i.split('/')[-1]))
                self.start_data.append(i)
                self.main_data, self.students_list, self.subjects = start_data_parsing(i)
                self.subject_combo.addItems(self.subjects)

        elif self.file_type_combo.currentText() == 'Расписание':
            self.GetHighlightedPaths()
            paths = []
            for i in self.all_chosen_files:
                paths.append(i)
            self.tables = get_schedule(paths)
            table = self.tables[1]['table']
            for i in range(len(table)):
                self.students_group_combo.addItem(table[i]['name'])
            for i in range(len(self.tables)):
                self.module_combo.addItem('module ' + str(i + 1))
        else:
            self.GetHighlightedPaths()

    def GetHighlightedPaths(self):
        chosen_files = self.file_tree.selectionModel().selectedIndexes()
        self.all_chosen_files = set()
        for i in chosen_files:
            self.all_chosen_files.add(self.file_model.fileInfo(i).absoluteFilePath())

    def CreateCurveChart(self):
        if self.file_type_combo.currentText() == 'Ведомости':
            mytuple = self.all_chosen_files
            arr = Current_student_grades(*mytuple)
            self.fig = Graph_Of_Current_Grades(students=arr, number_of_st=self.student_combo.currentIndex())
            self.graphic_lay.removeWidget(self.figure)
            self.figure = Graphic(self.fig)
            self.graphic_lay.addWidget(self.figure)
        else:
            pass

    def CreateScheduleChart(self):
        if self.file_type_combo.currentText() == 'Расписание':
            if self.schedule_lay.count() > 1:
                self.schedule_lay.removeWidget(self.hist_chart)
            self.hist = print_chart(self.tables, self.module_combo.currentIndex(), self.students_group_combo.currentIndex())
            self.hist_chart = Graphic(self.hist)
            self.schedule_lay.addWidget(self.hist_chart, 1, QtCore.Qt.AlignTop)



    def CreatePieChart(self):
        if self.file_type_combo.currentText() == 'Данные абитуриентов':
            top_num = 10

            self.main_data, self.students_list, self.subjects = start_data_parsing(self.start_data[self.file_idx_combo.currentIndex()])
            if self.pie_chart_lay.count() > 1:
                self.pie_chart_lay.removeWidget(self.pie_chart)
            self.pie = start_data_drawPiechart(self.main_data)
            self.pie_chart = Graphic(self.pie)
            self.pie_chart_lay.addWidget(self.pie_chart)

    def CreateHist(self):
        if self.file_type_combo.currentText() == 'Данные абитуриентов':
            self.main_data, self.students_list, self.subjects = start_data_parsing(self.start_data[self.file_idx_combo.currentIndex()])
            if self.hist_chart_lay.count() > 1:
                self.hist_chart_lay.removeWidget(self.exams_hist_chart)
            self.exams_hist = start_data_drawHistogramm(str(self.subject_combo.currentText()), self.main_data)
            self.exams_hist_chart = Graphic(self.exams_hist)
            self.hist_chart_lay.addWidget(self.exams_hist_chart)

    def CreateBarChart(self):
        if self.file_type_combo.currentText() == 'Оценки':
            for name in self.all_chosen_files:
                name_of_file = name
                scrs, as10 = Pars(name_of_file)
                self.bar_chart = Gist(scrs, as10)
                if self.bar_chart_lay.count() > 1:
                    self.bar_chart_lay.removeWidget(self.bar)
                self.bar = Graphic(self.bar_chart)
                self.bar_chart_lay.addWidget(self.bar)
        else:
            pass


    def CreateFrames(self):
        form_frame = QFrame()
        form_frame.setFrameShape(QFrame.StyledPanel)
        form_lay = QVBoxLayout()
        inside_lay = QHBoxLayout()

        form_lay.addWidget(self.manage_tabs)
        form_lay.addItem(inside_lay)
        inside_lay.addWidget(self.process_btn)
        inside_lay.addWidget(self.file_type_combo)
        form_frame.setLayout(form_lay)
        vertical_main_frame = QFrame()
        vertical_main_frame.setFrameShape(QFrame.StyledPanel)

        ver_box = QVBoxLayout()
        ver_box.setContentsMargins(25, 20, 25, 25)
        ver_box.addWidget(self.scroll_info)
        vertical_main_frame.setLayout(ver_box)

        horizontal_splitter = QSplitter(QtCore.Qt.Horizontal)
        horizontal_splitter.addWidget(form_frame)
        horizontal_splitter.addWidget(vertical_main_frame)

        vertical_splitter = QSplitter(QtCore.Qt.Vertical)
        vertical_splitter.addWidget(horizontal_splitter)

        vertical_main_box = QVBoxLayout()
        vertical_main_box.addWidget(vertical_splitter)
        self.group_box.setLayout(vertical_main_box)

    def CreateMenu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        import_project = QAction('New', self)
        import_project.setShortcut('Ctrl+Q')

        import_project.triggered.connect(self.CreateFileTree)
        file_menu.addAction(import_project)

        save_curve_chart = QAction('Save curve chart', self)
        save_curve_chart.triggered.connect(self.SaveCurveChart)

        save_bar_chart = QAction('Save bar chart', self)
        save_bar_chart.triggered.connect(self.SaveBarChart)

        file_menu.addAction(save_curve_chart)
        file_menu.addAction(save_bar_chart)

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

        main_data = []
        for student in Current_student_grades(*files):
            main_data.append(student.group)
            main_data.append(student.results)
            self.data[student.name] = main_data
            main_data = []
        print(self.data)

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    app.exec_()

if __name__ == '__main__':
    main()
