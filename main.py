from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from myform import Ui_MainWindow
from TableCreator import TableCreator

import sys


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.mainTable.setColumnCount(1)
        self.ui.mainTable.setRowCount(1)
        self.ui.plusColButon.clicked.connect(self.plus_col)
        self.ui.plusRowButon.clicked.connect(self.plus_row)
        self.ui.clearBut.clicked.connect(self.clear)
        self.ui.deleteBut.clicked.connect(self.delete)
        self.ui.action.triggered.connect(self.save)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Delete:
            self.delete()

    def clear(self):
        selected = self.ui.mainTable.selectedIndexes()
        for i in selected:
            self.ui.mainTable.setItem(i.row(), i.column(), None)

    def delete(self):
        rows = []
        for model_index in self.ui.mainTable.selectionModel().selectedRows():
            rows.append(model_index.row())
        rows.reverse()
        for i in rows:
            self.ui.mainTable.removeRow(i)
        cols = []
        for model_index in self.ui.mainTable.selectionModel().selectedColumns():
            cols.append(model_index.column())
        cols.reverse()
        for i in cols:
            self.ui.mainTable.removeColumn(i)

    def plus_col(self):
        self.ui.mainTable.insertColumn(self.ui.mainTable.columnCount())

    def plus_row(self):
        self.ui.mainTable.insertRow(self.ui.mainTable.rowCount())

    def save(self):
        options = QtWidgets.QFileDialog.Options()
        filename, form = QtWidgets.QFileDialog.getSaveFileName(self, "Save To File", "untitled",
                                                            "Text Files (*.odt);;Text Files (*.docx)",
                                                                 options=options)
        if filename:
            TableCreator.save_table(filename, form, self.ui.mainTable)


app = QtWidgets.QApplication([])
application = MainWindow()
application.show()

sys.exit(app.exec())