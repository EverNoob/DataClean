__author__ = 'Rowbot'

import PySide
from PySide.QtCore import QRect, QMetaObject, QObject
from PySide.QtGui  import (QApplication, QMainWindow, QWidget,
                           QGridLayout, QTabWidget, QPlainTextEdit,
                           QMenuBar, QMenu, QStatusBar, QAction,
                           QIcon, QFileDialog, QMessageBox, QFont)

import dataClean

class Scrubber(QMainWindow):
    def __init__(self, parent=None):
        super(Scrubber, self).__init__(parent)
        self.resize(800, 600)
        self.filename = None
        self.filetuple = None
        self.dirty = False
        centralwidget = QWidget(self)
        gridLayout = QGridLayout(centralwidget)
        self.tabWidget = QTabWidget(centralwidget)
        self.tab = QWidget()
        font = QFont()
        font.setFamily("Courier 10 Pitch")
        font.setPointSize(12)
        self.tab.setFont(font)
        gridLayout_3 = QGridLayout(self.tab)
        self.plainTextEdit = QPlainTextEdit(self.tab)
        gridLayout_3.addWidget(self.plainTextEdit, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        self.tab_2.setFont(font)
        gridLayout_2 = QGridLayout(self.tab_2)
        self.plainTextEdit_2 = QPlainTextEdit(self.tab_2)
        gridLayout_2.addWidget(self.plainTextEdit_2, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab_2, "")
        gridLayout.addWidget(self.tabWidget, 0, 0, 1, 1)
        self.setCentralWidget(centralwidget)
        menubar = QMenuBar(self)
        menubar.setGeometry(QRect(0, 0, 800, 29))
        menu_File = QMenuBar(self)
        self.menu_Clean = QMenu(menubar)
        self.menu_Help = QMenu(menubar)
        self.setMenuBar(menubar)
        self.statusbar = QStatusBar(self)
