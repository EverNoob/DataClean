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
                          QFont, QVBoxLayout, QFileDialog, QPushButton, QToolBar, QIcon)


class CleanDisplay(QMainWindow):
    def __init__(self, parent=None):
        super(CleanDisplay, self).__init__(parent)
        self.resize(800, 600)
        self.setWindowTitle("DATA CLEAN")
        self.filename = None
        self.input_filetuple = None
        self.output_filetuple = None
        self.raw_data = []
        self.cleaned_data = []
        centralwidget = QWidget(self)
        centralwidget_layout = QGridLayout(centralwidget)
        self.tabContainer = QTabWidget(centralwidget)
        # Raw Ledger Tab
        self.rawledger_tab = QWidget()
        font = QFont("Helvetica [Cronyx]", 12)
        self.rawledger_tab.setFont(font)
        rawledgertab_layout = QVBoxLayout(self.rawledger_tab)
        self.rawledger_TableView = QTableView()
        rawledgertab_layout.addWidget(self.rawledger_TableView)
        self.rawledger_TableView.setSortingEnabled(True)
        self.tabContainer.addTab(self.rawledger_tab, "Raw Ledger Data")
        self.process_button = QPushButton("&Process File", self)
        self.process_button.clicked.connect(self.process)
        rawledgertab_layout.addWidget(self.process_button)
        # Revision Tab
        self.revision_tab = QWidget()
        self.revision_tab.setFont(font)
        revision_tab_layout = QGridLayout(self.revision_tab)
        self.revision_tableview = QTableView()
        revision_tab_layout.addWidget(self.revision_tableview)
        self.tabContainer.addTab(self.revision_tab, "Revision")
        self.revision_tableview.setSortingEnabled(True)
        ## Results Tab
        self.results_tab = QWidget()
        self.results_tab.setFont(font)
        results_tab_layout = QVBoxLayout(self.results_tab)
        self.results_tableview = QTableView()
        results_tab_layout.addWidget(self.results_tableview)
        self.tabContainer.addTab(self.results_tab, "Results")
        centralwidget_layout.addWidget(self.tabContainer, 0, 0, 1, 1)
        self.setCentralWidget(centralwidget)
        self.export_button = QPushButton("&Export", self)
        self.export_button.clicked.connect(self.export_results)
        results_tab_layout.addWidget(self.export_button)
        self.results_tableview.setSortingEnabled(True)
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
        self.action_process = QAction("&Process Raw Ledger Data", self)
        self.action_process.triggered.connect(self.process)
        self.menu_File.addAction((self.action_process))
        self.action_save = QAction("&Save Results", self)
        self.action_save.triggered.connect(self.export_results)
        self.menu_File.addAction(self.action_save)
        #toolbar
        self.toolbar = QToolBar()
        self.toolbar.addActions(self.action_open.setIcon(), self.action_process, self.action_save)

        #Process

    def fileopen(self):
        dir = (os.path.dirname(self.filename)
               if self.filename is not None else "./Input")
        self.input_filetuple = QFileDialog.getOpenFileName(self,
                                                     "Open File", dir,
                                                     "Data (*.xlsx)\nAll Files (*.*)")
        self.filename = self.input_filetuple[0]
        fname = self.filename
        print('filename: ', fname)
        print('dir: ', dir)
        # QFileDialog returns a tuple x with x[0] = file name and
        #  x[1] = type of filter.
        if fname:
            print(fname)
            self.loadfile(fname)
            self.filename = fname

    def loadfile(self, fname=None):
        wb = openpyxl.load_workbook(fname)
        sh = wb.get_active_sheet()
        self.raw_data = dataClean.build_raw_list(sh)
        self.rawledger_TableView.setModel(scrubdatamodel.ScrubDataModel(self, self.raw_data,
                                                                       scrubdatamodel.ScrubDataModel.dirtyheader))

    def process(self):
        self.cleaned_data = dataClean.generate_clean_data_list(self.raw_data)
        self.results_tableview.setModel(scrubdatamodel.ScrubDataModel(self, self.cleaned_data, scrubdatamodel.ScrubDataModel.cleanheader))

    def export_results(self):
        dir = (os.path.dirname(self.filename)
               if self.filename is not None else ".")
        self.output_filetuple = QFileDialog.getSaveFileName(self, "Save File", dir, "Data (*.xlsx)\nAllFiles (*.*)")
        filename = self.output_filetuple[0]
        dataClean.write_to_excel(filename, self.cleaned_data)
        print('filename: ', filename)
        print('dir: ', dir)

    def editAction(self, action, slot=None, shortcut=None, icon=None,
                     tip=None):
        '''This method adds to action: icon, shortcut, ToolTip,\
        StatusTip and can connect triggered action to slot '''
        if icon is not None:
            action.setIcon(QIcon(":/%s.png" % (icon)))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            action.triggered.connect(slot)
        return action

if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = CleanDisplay()
    frame.show()
    sys.exit(app.exec_())