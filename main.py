import xlrd, xlwt, math as m, pandas as pd, openpyxl  # подключение библиотек нужных для работы программы
import sys  # sys нужен для передачи argv в QApplication
import design  # конвертированный в py файл дизайна Qt

from PyQt5 import QtCore, QtGui, QtWidgets

class Exceldata():
    def __init__(self):
        super().__init__()
    def Data_E(k):
        xls = pd.ExcelFile('/home/user/Документы/test1.xlsx')
        df = pd.read_excel(xls, 'Лист1')
        dfP = pd.DataFrame()
        dfQ = pd.DataFrame()
        dfS = pd.DataFrame()
        i = 2
        n = 33
        P = []
        Q = []
        S = []
        while i <= n:
            x = ((df.iloc[i - 2, 0]))
            P.append(int(x))
            dfP.loc[i, 'P, кВт'] = x
            i = i + 1
        i = 2
        while i <= n:
            x = ((df.iloc[i - 2, 1]))
            Q.append(int(x))
            dfQ.loc[i, 'P, кВт'] = x
            i = i + 1
        i = 2
        while i <= n:
            x = m.sqrt(P[i - 2] ** 2 + Q[i - 2] ** 2)
            x = float('{:.2f}'.format(x))
            S.append(x)
            dfS.loc[i, 'S, кВА'] = x
            i = i + 1
        if k == 1:
            return (dfP)
        if k == 2:
            return (dfQ)
        if k == 3:
            return (dfS)
    def Data_P(self):
        pd.self.Data = pd.self.Data_E.copy(k=1)
        return(self.Data)
    def Data_Q(self):
        self.Data = self.Data_E(k=2)
        return(self.Data)
    def Data_S(self):
        self.Data = self.Data_E(k=3)
        return(self.Data)


class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton.clicked.connect(self.P)  # Выполнить функцию P при нажатии кнопки
        self.pushButton_2.clicked.connect(self.Q)  # Выполнить функцию Q при нажатии кнопки
        self.pushButton_3.clicked.connect(self.S)  # Выполнить функцию S при нажатии кнопки

    def P(self):
        self.listWidget.clear()  # На случай, если в списке уже есть элементы
        #directory1 = QtWidgets.QFileDialog.getOpenFileUrl(self, "Выберите папку")
        #print(str(directory1))
        xls = pd.ExcelFile('/home/user/Документы/test1.xlsx')
        df = pd.read_excel(xls, 'Лист1')
        dfP = pd.DataFrame()
        dfQ = pd.DataFrame()
        dfS = pd.DataFrame()
        i = 2
        n = 33
        P = []
        Q = []
        S = []
        while i <= n:
            x = ((df.iloc[i - 2, 0]))
            P.append(int(x))
            dfP.loc[i, 'P, кВт'] = x
            i = i + 1
        self.listWidget.addItem(str(dfP))
        dfP.to_excel("test2.xlsx")

    def Q(self):
        self.listWidget.clear()  # На случай, если в списке уже есть элементы
        xls = pd.ExcelFile('/home/user/Документы/test1.xlsx')
        df = pd.read_excel(xls, 'Лист1')
        dfP = pd.DataFrame()
        dfQ = pd.DataFrame()
        dfS = pd.DataFrame()
        i = 2
        n = 33
        P = []
        Q = []
        S = []
        while i <= n:
            x = ((df.iloc[i - 2, 0]))
            P.append(int(x))
            dfP.loc[i, 'P, кВт'] = x
            i = i + 1
        i = 2
        while i <= n:
            x = ((df.iloc[i - 2, 1]))
            Q.append(int(x))
            dfQ.loc[i, 'Q, кВА'] = x
            i = i + 1
        self.listWidget.addItem(str(dfQ))
        dfQ.to_excel("test2.xlsx")

    def S(self):
        self.listWidget.clear()  # На случай, если в списке уже есть элементы
        xls = pd.ExcelFile('/home/user/Документы/test1.xlsx')
        df = pd.read_excel(xls, 'Лист1')
        dfP = pd.DataFrame()
        dfQ = pd.DataFrame()
        dfS = pd.DataFrame()
        i = 2
        n = 33
        P = []
        Q = []
        S = []
        while i <= n:
            x = ((df.iloc[i - 2, 0]))
            P.append(int(x))
            dfP.loc[i, 'P, кВт'] = x
            i = i + 1
        i = 2
        while i <= n:
            x = ((df.iloc[i - 2, 1]))
            Q.append(int(x))
            dfQ.loc[i, 'P, кВт'] = x
            i = i + 1
        i = 2
        while i <= n:
            x = m.sqrt(P[i - 2] ** 2 + Q[i - 2] ** 2)
            x = float('{:.2f}'.format(x))
            S.append(x)
            dfS.loc[i, 'S, кВА'] = x
            i = i + 1
        self.listWidget.addItem(str(dfS))
        dfS.to_excel("test2.xlsx")


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    D = Exceldata # Загрузка данных
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
