from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtWidgets import *
import sys
import subprocess
from random import choice
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5 import NavigationToolbar2QT as NavigationToolbar
from Canvas import Graphic

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.InitUI()
        self.MakeFrames()
        self.CreateMenu()
        self.gen_dict()
        self.ImportIntoTable()

    def gen_dict(self):
        lst = ['d', 'g', 'e', 'a', 'u']
        numbers = [1, 3, 5, 2, 6, 7, 8, 9]
        dictionary = {}
        values = (1, 3, 5)
        for i in range(20):
            dictionary[choice(lst)] = values
        print(dictionary)
        return dictionary

    def CreateGraphic(self):
        self.graphic.Plot()

    def InitUI(self):
        self.setStyleSheet(open('style.css', 'r').read())
        self.setWindowTitle('title')
        self.setGeometry(500, 500, 500, 500)
        self.showMaximized()
        self.graphic = Graphic(5, 5, 100)
        self.table = QTableWidget()
        self.graphic_frame = QFrame()
        self.button_lay = QHBoxLayout()
        self.graphic_lay = QVBoxLayout()

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
        self.show_info_tabs.addTab(self.table, 'Text Tab')
        self.show_info_tabs.addTab(self.graphic_frame, 'Text tab')
        self.manage_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.manage_tabs.setAutoFillBackground(True)
        self.show_info_tabs.setStyleSheet(open('style.css', 'r').read())
        self.process_btn = QPushButton('Process')
        self.build_graphic_btn = QPushButton('Create graphic')
        self.build_graphic_btn.clicked.connect(self.CreateGraphic)
        self.file_type_combo = QComboBox()

        self.student_combo = QComboBox()

        self.button_lay.addWidget(self.build_graphic_btn)
        self.button_lay.addWidget(self.student_combo)
        self.graphic_lay.addItem(self.button_lay)
        self.graphic_lay.addWidget(self.graphic)
        self.graphic_frame.setLayout(self.graphic_lay)

    #        self.txt_field.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
#        self.txt_field.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)


    def ImportIntoTable(self):
        self.data = self.gen_dict()
        self.table.setRowCount(len(self.data.keys())*3)
        self.table.setColumnCount(3)
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().sectionResizeMode(QHeaderView.Fixed)
        self.table.resizeColumnsToContents()
        self.table.scrollToBottom()

        names = self.data.keys()
        marks = []

        for i in self.data.values():
            for j in i:
                marks.append(j)
        marks.reverse()

        for i in range(len(names)*3):
            self.table.setSpan(i, 0, 3, 1)
        cnt = 0
        var = 0
        for name in names:
            name_item = QTableWidgetItem()
            subject_item = QTableWidgetItem()
            mark_item = QTableWidgetItem()
            name_item.setText(str(name))
            mark_item.setText(str(marks.pop()))
            self.table.setItem(cnt, 0, name_item)
            self.table.setItem(var, 1, mark_item)
            cnt += 3
            var += 1


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

        show_bar_btn = QAction('Some command', self)
        show_bar_btn.setShortcut('Ctrl+W')
        file_menu.addAction(show_bar_btn)

        edit = menubar.addMenu('Edit')
        view = menubar.addMenu('View')
        navigate = menubar.addMenu('Navigate')

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
        self.file_tree.clicked.connect(self.OpenFile)

    def OpenFile(self, index):
        file_path = self.file_model.fileInfo(index).absoluteFilePath()
        image_path = file_path.split('/')[-1]

        if image_path.split('/')[-1] == 'exe':
            subprocess.call(file_path)
            return
        if image_path.split('.')[-1] == 'png':
            self.img_label.setPixmap(QtGui.QPixmap(file_path))
            self.img_slider = QSlider(QtCore.Qt.Horizontal, self.img_label)
            self.setGeometry(60, 50, 100, 500)
            return
        file = open(file_path, 'r')
        info_from_file = file.read()

        self.txt_field.setText(info_from_file)
        file.close()

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    app.exec_()

if __name__ == '__main__':
    main()
