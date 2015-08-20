import openpyxl
import dataClean

__author__ = 'Rowbot'

import os
import sys
import scrubdatamodel

from PySide.QtCore import QRect
from PySide.QtGui import (QApplication, QMainWindow, QWidget,
                          QGridLayout, QTabWidget, QTableView,
                          QMenuBar, QMenu, QAction,
                          QFont, QVBoxLayout, QFileDialog)


class CleanDisplay(QMainWindow):
    def __init__(self, parent=None):
        super(CleanDisplay, self).__init__(parent)
        self.resize(800, 600)
        self.filename = None
        centralwidget = QWidget(self)
        centralwidget_layout = QGridLayout(centralwidget)
        self.tabContainer = QTabWidget(centralwidget)
        self.revisionTab = QWidget()
        font = QFont("Helvetica [Cronyx]", 12)
        self.revisionTab.setFont(font)
        revisiontab_layout = QVBoxLayout(self.revisionTab)
        self.revision_TableView = QTableView()
        revisiontab_layout.addWidget(self.revision_TableView)
        self.revision_TableView.setSortingEnabled(True)
        self.tabContainer.addTab(self.revisionTab, "Revision")
        ## Results Tab
        self.resultsTab = QWidget()
        self.resultsTab.setFont(font)
        results_tab_layout = QVBoxLayout(self.resultsTab)
        self.results_TableView = QTableView()
        results_tab_layout.addWidget(self.results_TableView)
        self.tabContainer.addTab(self.resultsTab, "Results")
        centralwidget_layout.addWidget(self.tabContainer, 0, 0, 1, 1)
        self.setCentralWidget(centralwidget)
        # Menu Bar/ Menus
        menubar = QMenuBar(self)
        menubar.setGeometry(QRect(0, 0, 800, 29))
        self.menu_File = QMenu(menubar)
        self.menu_File.setTitle("&File")
        menubar.addMenu(self.menu_File)
        self.menu_Help = QMenu(menubar, title='Help')
        menubar.addMenu(self.menu_Help)
        self.setMenuBar(menubar)
        self.action_open = QAction("&Open", self)
        self.action_open.triggered.connect(self.fileopen)
        self.menu_File.addAction(self.action_open)

    def fileopen(self):
        dir = (os.path.dirname(self.filename)
               if self.filename is not None else ".")
        self.filetuple = QFileDialog.getOpenFileName(self,
                                                     "Open File", dir,
                                                     "Data (*.xlsx)\nAll Files (*.*)")
        self.filename = self.filetuple[0]
        fname = self.filename
        # QFileDialog returns a tuple x with x[0] = file name and
        #  x[1] = type of filter.
        if fname:
            print(fname)
            self.loadfile(fname)
            self.filename = fname

    def loadfile(self, fname=None):
        wb = openpyxl.load_workbook(fname)
        sh = wb.get_active_sheet()
        raw_data_list = dataClean.build_raw_list(sh)
        self.revision_TableView.setModel(scrubdatamodel.ScrubDataModel(self, raw_data_list,
                                                                       scrubdatamodel.ScrubDataModel.dirtyheader))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = CleanDisplay()
    frame.show()
    sys.exit(app.exec_())