import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import pandas as pd
import matplotlib.pyplot as plt
import os

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("UI22.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        #pixmap = QPixmap('cutter.png')
        #self.lbl_img.setPixmap(pixmap)

        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        #self.tbl.setStyleSheet("QWidget { background-color: #aa8888; } QHeaderView::section { background-color: #88aa88; } QTableWidget QTableCornerButton::section {background-color: #8888aa; }")

        self.open_btn.clicked.connect(self.clickopenbtn)
        self.tbl.cellDoubleClicked.connect(self.call_chart)
        self.img_btn.clicked.connect(self.clickimgbtn)

    def clickimgbtn(self):
        file_name, ext = QFileDialog.getOpenFileName(self, '파일 열기', os.getcwd(), 'PNG (*.png)')
        if file_name:
            pixmap = QPixmap(file_name)
            self.lbl_img.setPixmap(pixmap)


    def clickopenbtn(self):
        file_path, ext = QFileDialog.getOpenFileName(self, '파일 열기', os.getcwd(), 'excel file (*.xls *.xlsx)')
        if file_path:
            self.df_list = self.loadData(file_path)
            self.initTableWidget(0)

    def loadData(self, file_name):
        df_list = []
        with pd.ExcelFile(file_name) as wb:
            for i, sn in enumerate(wb.sheet_names):
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                except Exception as e:
                    print('File read error:', e)
                else:
                    df_list.append(df)
        return df_list

    def initTableWidget(self, id):
        self.tbl.clear()
        df = self.df_list[id];
        col = len(df.keys())
        self.tbl.setColumnCount(col)
        self.tbl.setHorizontalHeaderLabels(df.keys())
        # dataframe의 행 개수 확인
        row = len(df.index)
        self.tbl.setRowCount(row)
        self.tbl.setVerticalHeaderLabels((
                                             "Cutter1;Cutter2;Cutter3;Cutter4;Cutter5;Cutter6;Cutter7;Cutter8;Cutter9;Cutter10;Cutter11;Cutter12").split(
            ';'))
        self.writeTableWidget(id, df, row, col)

    def writeTableWidget(self, id, df, row, col):
        for r in range(row):
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r][c]))
                item.setTextAlignment(Qt.AlignCenter)
                self.tbl.setItem(r, c, item)

    def call_chart(self, row, col):
        item = self.tbl.item(row, col)
        dtf = pd.DataFrame(self.df_list[row + 1])
        if col == 0:
            col = 'Wear'
        elif col == 1:
            col = 'Weight'
        else:
            col = 'RPM'
        dtf.set_index(dtf.columns[0], drop=True, inplace=True)
        name = 'Cutter' + str(row + 1) + ' - ' + col
        dtf = dtf[col]
        dtf.plot()
        plt.title(name)
        plt.show()

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()