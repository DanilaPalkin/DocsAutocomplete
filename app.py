import sys, os
import layout
from PyQt6 import QtWidgets, QtCore
import pandas as pd
from docxtpl import DocxTemplate

tupleTemplate = ('AAAA', 'BBBB', 'CCCC', 'DDDD', 'EEEE', 'FFFF', 'GGGG', 'HHHH', 'IIII',
                'JJJJ', 'KKKK', 'LLLL', 'MMMM', 'NNNN', 'OOOO', 'PPPP', 'QQQQ', 'RRRR',
                'SSSS', 'TTTT', 'UUUU', 'VVVV', 'WWWW', 'XXXX', 'YYYY', 'ZZZZ')

dictionaryTemplate = { 'AAAA': None, 'BBBB': None, 'CCCC': None, 'DDDD': None, 'EEEE': None, 'FFFF': None, 'GGGG': None, 'HHHH': None, 'IIII': None,
            'JJJJ': None, 'KKKK': None, 'LLLL': None, 'MMMM': None, 'NNNN': None, 'OOOO': None, 'PPPP': None, 'QQQQ': None, 'RRRR': None,
            'SSSS': None, 'TTTT': None, 'UUUU': None, 'VVVV': None, 'WWWW': None, 'XXXX': None, 'YYYY': None, 'ZZZZ': None }

class App(QtWidgets.QMainWindow, layout.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton.clicked.connect(self.browseDocx)
        self.pushButton_2.clicked.connect(self.browseXlsx)
        self.pushButton_3.clicked.connect(self.saveData)

    def browseDocx(self):
        docxName, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать шаблон", "","Документ Word (*.docx)")
        if docxName:
            print(docxName)
            self.label_2.setText(QtCore.QFileInfo(docxName).fileName())

    def browseXlsx(self):
        xlsxName, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать базу данных", "","Книга Excel (*.xlsx)")
        if xlsxName:
            print(xlsxName)
            self.label.setText(QtCore.QFileInfo(xlsxName).fileName())
            self.loadExcelData(xlsxName)

    def loadExcelData(self, fileName):
        #df = pd.read_excel(fileName, QtCore.QFileInfo(fileName).fileName())
        xl = pd.ExcelFile(fileName)
        sheetList = xl.sheet_names
        df = pd.read_excel(fileName, sheetList[0])
        if df.size == 0:
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

    def saveData(self, docxName):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите место сохранения файлов")
        if directory:  # не продолжать выполнение, если пользователь не выбрал директорию
            doc = DocxTemplate("шаблон.docx")
            print(self.tableWidget.rowCount())
            print(self.tableWidget.columnCount())
            for j in range(self.tableWidget.rowCount()):
                for i in range(self.tableWidget.columnCount()):
                    dictionaryTemplate[f"{tupleTemplate[i]}"] = self.tableWidget.item(j, i).text()
                    #dictionaryTemplate.update(i = self.tableWidget.item(j, i).text())
                doc.render(dictionaryTemplate)
                doc.save(f"{j + 1}_output.docx")   

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = App()
    window.show()
    app.exec()

if __name__ == '__main__':
    main()