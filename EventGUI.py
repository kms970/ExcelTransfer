import sys
import os
from PySide2.QtWidgets import *
from PySide2 import QtUiTools
from PySide2.QtWidgets import QMainWindow
import MakeExcel


class MainView(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):
        global UI_set

        UI_set = QtUiTools.QUiLoader().load(resource_path("mainframeGUI.ui"))
        UI_set.loadex.clicked.connect(self.setLoadexClicked)
        UI_set.saveex.clicked.connect(self.setSaveexClicked)
        UI_set.changeEx.clicked.connect(self.setChangeexClicked)

        self.setCentralWidget(UI_set)
        self.setWindowTitle("엑셀변환기")
        self.resize(529, 525)
        self.show()

    def setLoadexClicked(self):
        global loadFilename
        global loadFilenames
        if UI_set.fromSell.isChecked():
            loadFilename = QFileDialog.getOpenFileName(self, self.tr("파일 열기"), "./", self.tr("Data Files (*.csv *.xls *.xlsx);; All Files(*.*)"))
            UI_set.lookloadex.setText(loadFilename[0])

        elif UI_set.fromDel.isChecked():
            loadFilenames = QFileDialog.getOpenFileNames(self, self.tr("파일 열기"), "./", self.tr("Data Files (*.csv *.xls *.xlsx);; All Files(*.*)"))
            UI_set.lookloadex.setText(str(loadFilenames[0]))

    def setSaveexClicked(self):
        global fileDirectory
        fileDirectory = QFileDialog.getExistingDirectory(self, self.tr("Open Data files"), "./", QFileDialog.ShowDirsOnly)
        UI_set.looksaveex.setText(fileDirectory)

    def setChangeexClicked(self):
        qT = UI_set.quantityType.text()
        qE = UI_set.quantityEdit.text()
        fDE = UI_set.fareDivisionEdit.text()
        sFN = UI_set.saveFileName.text()
        if UI_set.radioNaver.isChecked() and UI_set.fromSell.isChecked():
            MakeExcel.MakeNaverToDel(loadFilename[0], fileDirectory, qT, qE, fDE, sFN)

        elif UI_set.radioCoupang.isChecked() and UI_set.fromSell.isChecked():
            MakeExcel.MakeCoupangToDel(loadFilename[0], fileDirectory, qT, qE, fDE, sFN)

        elif UI_set.fromDel.isChecked():
            MakeExcel.CutDeliveryEx(loadFilenames[0], fileDirectory, sFN)


# 파일 경로

# pyinstaller로 원파일로 압축할때 경로 필요함

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)

    return os.path.join(os.path.abspath("."), relative_path)