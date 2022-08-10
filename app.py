import sys, os
import layout
from PyQt6 import QtWidgets, QtCore
import pandas as pd
from docxtpl import DocxTemplate


class App(QtWidgets.QMainWindow, layout.Ui_MainWindow):
    tupleTemplate = ('AAAA', 'BBBB', 'CCCC', 'DDDD', 'EEEE', 'FFFF', 'GGGG', 'HHHH', 'IIII',
                    'JJJJ', 'KKKK', 'LLLL', 'MMMM', 'NNNN', 'OOOO', 'PPPP', 'QQQQ', 'RRRR',
                    'SSSS', 'TTTT', 'UUUU', 'VVVV', 'WWWW', 'XXXX', 'YYYY', 'ZZZZ')

    dictionaryTemplate = { 'AAAA': None, 'BBBB': None, 'CCCC': None, 'DDDD': None, 'EEEE': None, 'FFFF': None, 'GGGG': None, 'HHHH': None, 'IIII': None,
                'JJJJ': None, 'KKKK': None, 'LLLL': None, 'MMMM': None, 'NNNN': None, 'OOOO': None, 'PPPP': None, 'QQQQ': None, 'RRRR': None,
                'SSSS': None, 'TTTT': None, 'UUUU': None, 'VVVV': None, 'WWWW': None, 'XXXX': None, 'YYYY': None, 'ZZZZ': None }
    
    docxPath = ""
    xlsxPath = ""
    currentPage = 0

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton.clicked.connect(self.browseDocx)
        self.pushButton_2.clicked.connect(self.browseXlsx)
        self.pushButton_3.clicked.connect(self.saveData)
        self.pushButton_4.clicked.connect(self.pageBack)
        self.pushButton_5.clicked.connect(self.pageForward)

    def browseDocx(self):
        self.docxPath, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать шаблон", "","Документ Word (*.docx)")
        if self.docxPath:
            print(self.docxPath)
            self.lineEdit_2.setText(QtCore.QFileInfo(self.docxPath).fileName())
            self.lineEdit_2.textChanged = True
            self.buttonStateChangeCheck()

    def browseXlsx(self):
        self.xlsxPath, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать базу данных", "","Книга Excel (*.xlsx)")
        if self.xlsxPath:
            print(self.xlsxPath)
            self.currentPage = 0
            self.lineEdit_3.setText(QtCore.QFileInfo(self.xlsxPath).fileName())
            self.lineEdit_3.textChanged = True
            self.buttonStateChangeCheck()
            self.loadExcelData()

    def pageBack(self):
        self.currentPage -= 1
        self.loadExcelData()

    def pageForward(self):
        self.currentPage += 1
        self.loadExcelData()

    def loadExcelData(self):
        #df = pd.read_excel(fileName, QtCore.QFileInfo(fileName).fileName())
        xl = pd.ExcelFile(self.xlsxPath)
        sheetList = xl.sheet_names

        if self.currentPage < 0:
            self.currentPage = 0
        if self.currentPage >= len(xl.sheet_names):
            self.currentPage = len(xl.sheet_names) - 1

        df = pd.read_excel(self.xlsxPath, sheetList[self.currentPage])
        self.label.setText(sheetList[self.currentPage])
        if df.size == 0:
            self.label.setText("Пустая страница!")
            self.tableWidget.clear()
            return
        df.fillna('', inplace=True)
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        #self.tableWidget.setHorizontalHeaderLabels(df.columns) # выдаёт ошибку если хэдер не str

        # return panda array object
        for row in df.iterrows():
            values = row[1]
            for col_index, value in enumerate(values):
                tableItem = QtWidgets.QTableWidgetItem(str(value))
                self.tableWidget.setItem(row[0], col_index, tableItem)

    def saveData(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите место сохранения файлов")
        if directory:  # не продолжать выполнение, если пользователь не выбрал директорию
            doc = DocxTemplate(f"{self.docxPath}") # нужно передать путь к шаблону
            print(directory)
            print(self.tableWidget.rowCount())
            print(self.tableWidget.columnCount())
            for j in range(self.tableWidget.rowCount()):
                for i in range(self.tableWidget.columnCount()):
                    self.dictionaryTemplate[f"{self.tupleTemplate[i]}"] = self.tableWidget.item(j, i).text()
                    #dictionaryTemplate.update(i = self.tableWidget.item(j, i).text())
                doc.render(self.dictionaryTemplate)
                doc.save(f"{directory}/{j + 1}_output.docx")
    
    def buttonStateChangeCheck(self):
        if (self.lineEdit_2.textChanged == True) and (self.lineEdit_3.textChanged == True):
            self.pushButton_3.setEnabled(1)
            self.pushButton_4.setEnabled(1)
            self.pushButton_5.setEnabled(1)

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = App()
    window.show()
    app.exec()

if __name__ == '__main__':
    main()