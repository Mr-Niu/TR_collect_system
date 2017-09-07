from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import subprocess
import os
from TR_server import *
import time
import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("server")

text='ui'
dataPath = 'TR_data.xls'
path = os.getcwd()
filePath = path + '\\data\\' + dataPath


class window(QMainWindow, QWidget):
    def __init__(self, parent = None):
        super(window, self).__init__(parent)
        self.icon = QIcon('./pics/icon.png')
        self.icon_right = QPixmap('./pics/right.png')
        self.icon_error = QPixmap('./pics/error.png')
        self.isRun = False
        self.createUi()

        self.server = server()
        self.connect()
        # self.up();

    def createUi(self):
        # 窗口
        self.setWindowTitle('本地TR记录')
        self.setWindowIcon(self.icon)
        # 状态栏
        self.statusbar = QStatusBar()
        self.barLight = QLabel()
        self.barLight.setPixmap(self.icon_error)
        self.barLabel = QLabel()
        self.statusbar.addWidget(self.barLight)
        self.statusbar.addWidget(self.barLabel)
        self.setStatusBar(self.statusbar)

        # 启动、停止、打开； 服务器信息
        self.startButton = QPushButton()
        self.startButton.setText('启动服务器')
        self.stopButton = QPushButton()
        self.stopButton.setText('关闭服务器')
        self.reloadButton =QPushButton()
        self.reloadButton.setText('加载数据库')
        self.openDateButton = QPushButton('打开数据库')
        self.ipLabel = QLabel('本机IP及端口：')
        ip = self.getIp() + ' : 6667'
        self.ipLabel2 = QLabel(ip)
        font = QFont('Arial')
        font.setBold(True)
        self.ipLabel2.setFont(font)

        # 连接的客户端及事件
        self.listLabel = QLabel()

        self.listLabel.setAlignment(Qt.AlignTop)
        self.scrollArea = QScrollArea(self)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setFixedWidth(200)
        self.scrollArea.setWidget(self.listLabel)
        self.table = QTableWidget(10, 4)
        self.table.setHorizontalHeaderLabels(['序号', '提交IP', '内容', '附件'])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.eventText = QPlainTextEdit()
        self.eventText = QTextBrowser()
        # self.eventText.setText('this is events occured')
        self.eventText.setMaximumHeight(150)

        #布局
        self.hLayout1 = QHBoxLayout()
        self.hLayout1.addWidget(self.startButton)
        self.hLayout1.addWidget(self.stopButton)
        self.hLayout1.addWidget(self.reloadButton)
        self.hLayout1.addWidget(self.openDateButton)
        self.hLayout1.addStretch(1)

        self.hLayout2 = QHBoxLayout()
        self.hLayout2.addWidget(self.ipLabel)
        self.hLayout2.addWidget(self.ipLabel2)
        self.hLayout2.addStretch(1)


        self.hLayout3 = QHBoxLayout()
        self.vLayout_table_event = QVBoxLayout()
        self.hLayout3.addWidget(self.scrollArea)
        self.vLayout_table_event.addWidget(self.table)
        self.vLayout_table_event.addWidget(self.eventText)
        self.hLayout3.addLayout(self.vLayout_table_event)

        self.vLayout1 = QVBoxLayout()
        self.vLayout1.addLayout(self.hLayout1)
        self.vLayout1.addLayout(self.hLayout2)
        self.vLayout1.addLayout(self.hLayout3)

        self.widget = QWidget(self)
        self.widget.setLayout(self.vLayout1)
        self.setCentralWidget(self.widget)

    def connect(self):
        self.startButton.clicked.connect(self.startServer)
        self.stopButton.clicked.connect(self.stopServer)
        self.openDateButton.clicked.connect(self.openDate)
        self.reloadButton.clicked.connect(self.loadData)

    def getIp(self):
        s = socket.socket()
        name = socket.gethostname()
        ip = socket.gethostbyname(name)
        return ip

    # 更新组件
    def update_widgets(self, text):
        global connectList
        s=''
        for i in range(len(connectList)):
            s += connectList[i] +'\n'
        self.listLabel.setText(s)

        # s = self.eventText.toPlainText()
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        s = t  + ' from '+ text
        # self.eventText.setText(s)
        self.eventText.append(s)
        self.loadData()

    def startServer(self):
        try:
            self.barLight.setPixmap(self.icon_right)
            self.barLabel.setText('Start server')
            self.server = server()
            self.server.ud[str].connect(self.update_widgets)
            # self.server.ud[list, str].connect(self.update_widgets)
            self.server.start()

        except Exception as e:
            print(e)
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('Start server error')


    def stopServer(self):
        try:
            self.server.stop()
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('Server stopped')
        except:
            self.barLabel.setText('Stop server failed')


    def openDate(self):
        try:
            # os.system('excel.exe -readOnly '+filePath)
            subprocess.Popen('excel.exe -readOnly '+filePath)
        except:
            self.barLabel.setText('打开失败，请确认是否已将Excel添加至系统环境变量')

    def loadData(self):
        try:
            excel = xlrd.open_workbook(filePath)
            sheet = excel.sheet_by_index(0)
            row = sheet.nrows
            for i in range(self.table.rowCount()):
                for j in range(self.table.columnCount()):
                    item = QTableWidgetItem('')
                    self.table.setItem(i, j, item)
            for i in range(row):
                if(self.table.rowCount()==i):
                    self.table.insertRow(i)
                for j in range(self.table.columnCount()):
                    try:
                        d = sheet.cell(i,j).value
                    except:
                        d=''
                    item = QTableWidgetItem(d)
                    self.table.setItem(i, j, item)
        except:
            self.barLabel.setText('Load data error')
        pass

    def closeEvent(self, event):
        self.server.stop()
        # self.centralWidget().closeEvent(event)
        print('exitted')
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = window()
    w.show()
    sys.exit(app.exec_())


