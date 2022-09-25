import sys, os
import layout, infowindow
from PyQt6 import QtWidgets, QtCore
import pandas as pd
from docxtpl import DocxTemplate

class InfoWindow(QtWidgets.QWidget, infowindow.Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

class App(QtWidgets.QMainWindow, layout.Ui_MainWindow):
    tupleTemplate = ('AAAA', 'BBBB', 'CCCC', 'DDDD', 'EEEE', 'FFFF', 'GGGG', 'HHHH', 'IIII',
                    'JJJJ', 'KKKK', 'LLLL', 'MMMM', 'NNNN', 'OOOO', 'PPPP', 'QQQQ', 'RRRR',
                    'SSSS', 'TTTT', 'UUUU', 'VVVV', 'WWWW', 'XXXX', 'YYYY', 'ZZZZ')

    dictionaryTemplate = { 'AAAA': None, 'BBBB': None, 'CCCC': None, 'DDDD': None, 'EEEE': None, 'FFFF': None, 'GGGG': None, 'HHHH': None, 'IIII': None,
                'JJJJ': None, 'KKKK': None, 'LLLL': None, 'MMMM': None, 'NNNN': None, 'OOOO': None, 'PPPP': None, 'QQQQ': None, 'RRRR': None,
                'SSSS': None, 'TTTT': None, 'UUUU': None, 'VVVV': None, 'WWWW': None, 'XXXX': None, 'YYYY': None, 'ZZZZ': None }
    
    docxPath = None
    xlsxPath = None
    currentPage = 0

    def __init__(self):
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton.clicked.connect(self.browseDocx)
        self.pushButton_2.clicked.connect(self.browseXlsx)
        self.pushButton_3.clicked.connect(self.saveData)
        self.pushButton_4.clicked.connect(self.pageBack)
        self.pushButton_5.clicked.connect(self.pageForward)
        self.pushButton_6.clicked.connect(self.showInfo)
        self.lineEdit.textChanged.connect(self.search)

    def showInfo(self):
        self.dialog = InfoWindow()
        self.dialog.show()

    def browseDocx(self):
        self.docxPath, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать шаблон", "","Документ Word (*.docx)")
        if self.docxPath:
            self.label_2.setText(QtCore.QFileInfo(self.docxPath).fileName())
            self.buttonStateChangeCheck()

    def browseXlsx(self):
        self.xlsxPath, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Выбрать базу данных", "","Книга Excel (*.xlsx)")
        if self.xlsxPath:
            self.currentPage = 0
            self.label_3.setText(QtCore.QFileInfo(self.xlsxPath).fileName())
            self.tableWidget.setRowCount(0) # необходимо, когда
            self.tableWidget.setColumnCount(0) # загружаешь новую таблицу
            self.tableWidget.clear() # при self.tableWidget.setSortingEnabled(True)
            self.loadExcelData()
            self.buttonStateChangeCheck()

    def pageBack(self):
        self.currentPage -= 1
        self.loadExcelData()

    def pageForward(self):
        self.currentPage += 1
        self.loadExcelData()

    def loadExcelData(self):
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
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            self.tableWidget.clear()
            self.pushButton_3.setEnabled(0)
            return
        else: self.buttonStateChangeCheck()
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

            for j in range(self.tableWidget.rowCount()):
                k = 0
                for i in range(self.tableWidget.columnCount()):
                    # работаем только с ячейками, которые выделены и не спрятаны фильтром поиска
                    if self.tableWidget.item(j, i).isSelected() and not self.tableWidget.isRowHidden(j):
                        self.dictionaryTemplate[f"{self.tupleTemplate[k]}"] = self.tableWidget.item(j, i).text()
                        k += 1
                if self.dictionaryTemplate["AAAA"] != None: # убедимся, что контекст не пустой
                    doc.render(self.dictionaryTemplate)
                    doc.save(f"{directory}/{j + 1}_output.docx")
                self.dictionaryTemplate = self.dictionaryTemplate.fromkeys(self.dictionaryTemplate, None)
    
    def buttonStateChangeCheck(self):
        if (self.docxPath != None) and (self.xlsxPath != None) and (self.tableWidget.rowCount() != 0):
            self.pushButton_3.setEnabled(1)
            self.pushButton_4.setEnabled(1)
            self.pushButton_5.setEnabled(1)

    def search(self, string):
        # очистка текущего выбора
        self.tableWidget.setCurrentItem(None)

        if not string:
            # пустая строка, не искать, показать всю таблицу
            for j in range(self.tableWidget.rowCount()):
                self.tableWidget.showRow(j)
            return

        # спрячем таблицу, будем показывать только подходящие под фильтр строки
        for j in range(self.tableWidget.rowCount()):
                self.tableWidget.hideRow(j)

        matchingItems = self.tableWidget.findItems(string, QtCore.Qt.MatchFlag.MatchContains)
        if matchingItems:
            for item in matchingItems:
                self.tableWidget.showRow(item.row())
        

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = App()
    window.show()
    app.exec()

if __name__ == '__main__':
    main()