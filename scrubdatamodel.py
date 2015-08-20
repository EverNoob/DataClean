import operator
import openpyxl

__author__ = 'Rowbot'

from PySide.QtGui import QApplication, QVBoxLayout, QFont, QTableView, QWidget
from PySide.QtCore import QAbstractTableModel, Qt, SIGNAL

import dataClean


class MyWindow(QWidget):
    def __init__(self, data_list, header, *args):
        QWidget.__init__(self, *args)
        # setGeometry(x_pos, y_pos, width, height)
        self.setGeometry(300, 200, 570, 450)
        self.setWindowTitle("Click on column title to sort")
        table_model = ScrubDataModel(self, data_list, header)
        table_view = QTableView()
        table_view.setModel(table_model)
        # set font
        font = QFont("Courier New", 14)
        table_view.setFont(font)
        # set column width to fit contents (set font first!)
        table_view.resizeColumnsToContents()
        # enable sorting
        table_view.setSortingEnabled(True)
        layout = QVBoxLayout(self)
        layout.addWidget(table_view)
        self.setLayout(layout)


class ScrubDataModel(QAbstractTableModel):
    cleanheader = ['Item code w/o vintage', 'Brand Code', 'Brand', 'Varietal Code', 'Varietal', 'Distributor',
                   'State', 'Sales Rep', 'Item ID', 'Item', 'SKU Tag', 'Item Pre', 'Size', 'Month', 'Year', 'Date',
                   'Document Type', 'Warehouse', 'Document #', 'Opposite #', 'Sales $', 'Total Cases', 'Vintage',
                   'Portfolio', 'Category', 'Sales/Key Acct Rep', 'ISM', 'IBM', 'Customer ID', 'Sales FOB', 'SKU Cost',
                   'SKU DA', 'Total DA$', 'GP$/CASE', 'Total GP$']

    dirtyheader = ['Customer Name', 'Ship-to State', 'Salesperson Code', 'Item No.', 'Description',
                   'Product Group Code', 'Posting Month', 'Posting Date', 'Year', 'Document Type', 'Location Code',
                   'Document No.', 'Quantity', 'Sales Amount (Actual)', 'Quantity (positive)', 'Vintage',
                   'Customer No.', 'Brand Code', 'Varietal Code']

    def __init__(self, parent, dict_list=None, header=None, *args):
        QAbstractTableModel.__init__(self, parent, *args)
        self.dict_list = dict_list
        self.header = header

    def setHeaderData(self, header, *args, **kwargs):
        self.header = header

    def rowCount(self, parent):
        return len(self.dict_list)

    def columnCount(self, parent):
        return len(self.dict_list[0])

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        elif role == Qt.DisplayRole or role == Qt.EditRole:
            try:
                return self.dict_list[index.row()][self.header[index.column()]]
            except IndexError as ie:
                print(ie)
                print('row:', index.row(), ', column:', index.column())
        else:
            print('None returned to data() call')
            return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and section in range(len(self.header)):
            return self.header[section]
        if orientation == Qt.Vertical and role == Qt.DisplayRole and section in range(len(self.dict_list)):
            return section + 1
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid():
            print('setData Failed: invalid index')
            return False
        elif role != Qt.EditRole:
            print('setData Failed: incorrect role entered')
            return False
        else:
            self.dict_list[index.row()][self.header[index.column()]] = value
            print('Dataset successful')
            return True

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEditable | Qt.ItemIsEnabled

    def sort(self, col, order):
        """sort table by given column number col"""
        self.emit(SIGNAL("layoutAboutToBeChanged()"))
        self.dict_list = sorted(self.dict_list, key=operator.itemgetter(self.header[col]))
        if order == Qt.DescendingOrder:
            self.dict_list.reverse()
        self.emit(SIGNAL("layoutChanged()"))



if __name__ == "__main__":


    wbtest = openpyxl.load_workbook('Input/Item Ledger Precept - 7.21.15.xlsx')
    shtest = wbtest.get_active_sheet()
    raw_data_list = dataClean.build_raw_list(shtest)
    clean_data = dataClean.generate_clean_data_list(raw_data_list)

    app = QApplication([])
    win = MyWindow(raw_data_list, ScrubDataModel.dirtyheader)
    win.show()
    app.exec_()