# Form implementation generated from reading ui file 'infowindow.ui'
#
# Created by: PyQt6 UI code generator 6.3.0
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(600, 400)
        self.gridLayout = QtWidgets.QGridLayout(Form)
        self.gridLayout.setObjectName("gridLayout")
        self.textBrowser = QtWidgets.QTextBrowser(Form)
        self.textBrowser.setStyleSheet("background-color: transparent")
        self.textBrowser.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.textBrowser.setLineWidth(0)
        self.textBrowser.setObjectName("textBrowser")
        self.gridLayout.addWidget(self.textBrowser, 0, 0, 1, 1)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "О программе"))
        self.textBrowser.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body>\n"
"<p>Программа предназначена для заполнения документов по шаблону в формате .docx данными из таблиц .xlsx</p>\n"
"<p>Как пользоваться?</p>\n"
"<p>1) Выбрать шаблон</p>\n"
"<p>2) Выбрать таблицу</p>\n"
"<p>3) Выбрать лист таблицы</p>\n"
"<p>4) Выделить необходимые ячейки</p>\n"
"<p>5) Выбрать место сохранения и сохранить</p>\n"
"<p><br /></p>\n"
"<p>Как написать свой шаблон?</p>\n"
"<p>Создайте документ или выберите существующую форму. В тех местах, куда будут подставлены данные, необходимо написать следующую конструкцию (тег): {{ AAAA }}</p>\n"
"<p>(от {{ AAAA }} до {{ ZZZZ }} при необходимости, итого до 26 различных полей)</p>\n"
"<p><br /></p>\n"
"<p>Две открывающие фигурные скобки, пробел, 4 заглавные буквы латинского алфавита, пробел, две закрывающие фигурные скобки.</p>\n"
"<p>Шрифт, размер шрифта, курсив, полужирность текста и т.д. сохранятся, будут такими же, как у тега.</p>\n"
"<p><br /></p>\n"
"<p>Программа работает следующим образом</p>\n"
"<p>Выделенные столбцы именуются в алфавитном порядке 1-й - A, 2-й - B, 3-й - C и т.д.</p>\n"
"<p>Данные из столбца встают на место соответствующего тега в шаблоне. Полученный документ сохраняется в выбранном месте. Процесс повторяется столько раз, сколько строк выделено.</p>\n"
"<p>Убедитесь, что число уникальных тегов в шаблоне равно числу выделенных столбцов.</p>\n"
"<p>Два одинаковых тега? - Текст из ячейки появится в двух местах.</p>\n"
"<p>Тегов меньше, чем выделенных столбцов? - Часть данных не появится в итоговом документе.</p>\n"
"<p>Тегов больше, чем выделенных столбцов? - На месте тегов появится надпись None.</p>\n"
"<p><br /></p>\n"
"<p>Работа с таблицей в программе:</p>\n"
"<p>Выделить ячейку - Левая кнопка мыши</p>\n"
"<p>Выделить несмежные ячейки - Ctrl + Левая кнопка мыши</p>\n"
"<p>Выделить строку или столбец - Левая кнопка мыши по номеру</p>\n"
"<p>Выделить несмежные строки или столбцы - Ctrl + Левая кнопка мыши по номеру</p>\n"
"<p>Выделить всю таблицу - Клик левой кнопкой мыши в левый верхний угол таблицы</p></body></html>"))
