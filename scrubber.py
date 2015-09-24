from PySide import QtCore

__author__ = 'Rowbot'


import threading
import os
import sys
import scrubdatamodel
import lookuptools
import dataClean

from PySide.QtCore import QRect, QFile, QTextStream
from PySide.QtGui import (QApplication, QMainWindow, QWidget,
                          QGridLayout, QTabWidget, QTableView,
                          QMenuBar, QMenu, QAction,
                          QFont, QVBoxLayout, QFileDialog, QPushButton, QToolBar, QIcon, QHBoxLayout, QTextEdit,
                          QDialog)
# abspath = os.path.abspath(__file__)
# dname = os.path.dirname(abspath)
# os.chdir(dname)


class CleanDisplay(QMainWindow):
    def __init__(self, parent=None):
        super(CleanDisplay, self).__init__(parent)
        self.cleanerbot = dataClean.DataCleaner()
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
        rawledger_hbox = QHBoxLayout()
        rawledger_hbox.addStretch(1)
        rawledger_hbox.addWidget(self.process_button)
        rawledger_hbox.addStretch(1)
        rawledgertab_layout.addLayout(rawledger_hbox)
        ## Results Tab
        self.results_tab = QWidget()
        self.results_tab.setFont(font)
        results_tab_layout = QVBoxLayout(self.results_tab)
        self.results_tableview = QTableView()
        results_tab_layout.addWidget(self.results_tableview)
        self.tabContainer.addTab(self.results_tab, "Results")
        centralwidget_layout.addWidget(self.tabContainer, 0, 0, 1, 1)
        self.setCentralWidget(centralwidget)
        self.export_button = QPushButton("&Export/Save", self)
        self.export_button.clicked.connect(self.export_results)
        self.process_button2 = QPushButton('&Process', self)
        self.process_button2.setFixedHeight(30)
        self.process_button2.setFixedWidth(120)
        self.process_button2.clicked.connect(self.process)
        results_hbox = QHBoxLayout()
        results_hbox.addStretch(1)
        results_hbox.addWidget(self.process_button2)
        results_hbox.addWidget(self.export_button)
        results_vbox = QVBoxLayout()
        results_vbox.addLayout(results_hbox)
        results_tab_layout.addLayout(results_vbox)
        self.results_tableview.setSortingEnabled(True)
        # Revision Tab
        self.revision_tab = QWidget()
        self.revision_tab.setFont(font)
        revision_tab_layout = QGridLayout(self.revision_tab)
        self.revision_tableview = QTableView()
        revision_tab_layout.addWidget(self.revision_tableview)
        self.tabContainer.addTab(self.revision_tab, "Revision")
        self.revision_tableview.setSortingEnabled(True)
        # Menu Bar/ Menus
        menubar = QMenuBar(self)
        menubar.setGeometry(QRect(0, 0, 800, 29))
        self.menu_File = QMenu(menubar)
        self.menu_File.setTitle("&File")
        self.setMenuBar(menubar)
        self.action_open = QAction("&Open", self)
        self.action_open.triggered.connect(self.fileopen)
        self.menu_File.addAction(self.action_open)
        self.action_process = QAction("&Process Raw Ledger Data", self)
        self.action_process.triggered.connect(self.process)
        self.menu_File.addAction((self.action_process))
        self.action_export = QAction("&Export/Save Results", self)
        self.action_export.triggered.connect(self.export_results)
        self.menu_File.addAction(self.action_export)
        self.action_exit = QAction("&Exit", self)
        self.action_exit.triggered.connect(self.program_exit)
        self.menu_File.addAction(self.action_exit)
        self.menu_Update = QMenu(menubar, title='Update')
        self.action_update_database = QAction("&Update Database", self)
        self.action_update_database.triggered.connect(self.update_main_db)
        self.editAction(self.action_update_database, None, None, 'updatedb', 'Update Main Database')
        self.menu_Update.addAction(self.action_update_database)
        self.action_update_brand_manager_lookup = QAction("&Update Brand Manager Lookup", self)
        self.action_update_brand_manager_lookup.triggered.connect(self.update_brand_manager_lookup)
        self.editAction(self.action_update_brand_manager_lookup, None, None, 'brandmanager', 'Update Brand Manager Lookup')
        self.menu_Update.addAction(self.action_update_brand_manager_lookup)
        self.help = QTextEdit(self)
        self.help_text = open('Help/help.txt').read()
        self.help.setPlainText(self.help_text)
        self.help.resize(700, 700)
        self.help.setWindowFlags(QtCore.Qt.Dialog)
        self.action_help = QAction("&Help", self)
        self.action_help.triggered.connect(self.show_help)
        self.menu_Help = QMenu(menubar, title="Help")
        self.menu_Help.addAction(self.action_help)
        menubar.addMenu(self.menu_File)
        menubar.addMenu(self.menu_Update)
        menubar.addMenu(self.menu_Help)
        #toolbar
        self.toolbar = QToolBar()
        self.editAction(self.action_open, None, 'ctrl+o', 'open', 'Open')
        self.editAction(self.action_process, None, 'ctrl+p', 'process', 'Process Raw Ledger Data')
        self.editAction(self.action_export, None, 'ctrl+E', 'save', 'Export')
        self.editAction(self.action_exit, None, 'alt+f4', 'exit', 'Exit')
        #toolbar
        self.toolbar = QToolBar()
        self.editAction(self.action_open, None, 'ctrl+o', 'open', 'Open')
        self.editAction(self.action_process, None, 'ctrl+p', 'process', 'Process Raw Ledger Data')
        self.editAction(self.action_export, None, 'ctrl+E', 'save', 'Export')
        self.toolbar.addActions((self.action_open, self.action_process, self.action_export))

        self.addToolBar(self.toolbar)

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
        self.raw_data = self.cleanerbot.build_raw_list(filename=fname)
        self.rawledger_TableView.setModel(scrubdatamodel.ScrubDataModel(self, self.raw_data,
                                                                       scrubdatamodel.ScrubDataModel.dirtyheader))

    def process(self):
        self.cleaned_data = self.cleanerbot.generate_clean_data_list(self.raw_data)
        self.results_tableview.setModel(scrubdatamodel.ScrubDataModel(self, self.cleaned_data, scrubdatamodel.ScrubDataModel.cleanheader))
        self.tabContainer.setCurrentIndex(1)

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
            action.setIcon(QIcon("Icons/actions/%s.png" % (icon)))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            action.triggered.connect(slot)
        return action

    def program_exit(self):
        quit()

    def update_main_db(self):
        t = threading.Thread(target=lookuptools.build_brand_lookup)
        t.start()
        self.cleanerbot.update_brand_dict()

    def update_brand_manager_lookup(self):
        lookuptools.build_brand_manager_lookups()
        self.cleanerbot.update_brand_manager_by_brand()

    def update_size_dict(self):
        lookuptools.build_size_lookup()
        self.cleanerbot.update_size_dict()

    def show_help(self):
        self.help.show()


def main():
    app = QApplication(sys.argv)
    frame = CleanDisplay()
    frame.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()