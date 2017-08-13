from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import subprocess
import xlrd
from xlutils.copy import copy
import os
import sys
import socket
import time
import threading
import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("client")

dataPath = 'Tr_client.xls'
path = os.getcwd()
filePath = path + '\\data\\' + dataPath
serverIP = ''
port = 6666
class window(QMainWindow, QWidget):
    def __init__(self, parent = None):
        super(window, self).__init__(parent)
        self.icon = QIcon('./pics/client2.png')
        self.icon_right = QPixmap('./pics/right.png')
        self.icon_error = QPixmap('./pics/error.png')
        self.isRun = False
        self.createUi()
        self.connect()

    def createUi(self):
        # 窗口
        self.setWindowTitle('TR记录客户端')
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
        self.scanButton = QPushButton()
        self.scanButton.setText('扫描服务器')
        self.reloadButton =QPushButton()
        self.reloadButton.setText('加载数据库')
        self.openDateButton = QPushButton('打开数据库')
        self.ipLabel = QLabel('服务器IP及端口：')
        self.ipLabel2 = QLabel('请先扫描服务器IP！')
        font = QFont('Arial')
        font.setBold(True)
        self.ipLabel2.setFont(font)

        # 连接的客户端及事件
        self.table = QTableWidget(10, 3)
        self.table.setHorizontalHeaderLabels(['序号', '提交IP', '内容'])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.trText = QTextEdit()
        self.trText.setMaximumHeight(150)
        font = QFont()
        font.setPixelSize(24)
        self.trText.setFont(font)
        self.sendButton = QPushButton()
        self.sendButton.setText('提交至服务器')

        #布局
        self.hLayout1 = QHBoxLayout()
        self.hLayout1.addWidget(self.scanButton)
        self.hLayout1.addWidget(self.reloadButton)
        self.hLayout1.addWidget(self.openDateButton)
        self.hLayout1.addStretch(1)

        self.hLayout2 = QHBoxLayout()
        self.hLayout2.addWidget(self.ipLabel)
        self.hLayout2.addWidget(self.ipLabel2)
        self.hLayout2.addStretch(1)


        self.vLayout_table_event = QVBoxLayout()
        self.hLayout3 = QHBoxLayout()
        self.vLayout_table_event.addWidget(self.table)
        self.vLayout_table_event.addWidget(self.trText)
        self.hLayout3.addStretch(1)
        self.hLayout3.addWidget(self.sendButton)
        self.vLayout_table_event.addLayout(self.hLayout3)

        self.vLayout1 = QVBoxLayout()
        self.vLayout1.addLayout(self.hLayout1)
        self.vLayout1.addLayout(self.hLayout2)
        self.vLayout1.addLayout(self.vLayout_table_event)

        self.widget = QWidget(self)
        self.widget.setLayout(self.vLayout1)
        self.setCentralWidget(self.widget)

    def connect(self):
        self.scanButton.clicked.connect(self.scanIp)
        self.reloadButton.clicked.connect(self.loadData)
        self.openDateButton.clicked.connect(self.openDate)
        self.sendButton.clicked.connect(self.sendTr)
        self.trText.textChanged.connect(self.state)

    def scanIp(self):
        global serverIP
        self.barLight.setPixmap(self.icon_error)
        self.barLabel.setText('正在扫描服务器IP...')
        for i in range(8):
            threading.Thread(target=self.getIp, args=(i*32,)).start()
        for i in range(50):
            if '' == serverIP:
                self.barLabel.setText('未能扫描到服务器，网络不通或服务器尚未启动！')
            else:
                self.ipLabel2.setText(serverIP + ': ' + str(port))
                self.barLight.setPixmap(self.icon_right)
                self.barLabel.setText('扫描服务器IP成功，服务器IP为：' + serverIP)
                break
            time.sleep(0.1)

    def getIp(self, start):
        global serverIP
        for i in range(start, start + 32):
            s = socket.socket()
            s.settimeout(0.1)
            iphead = socket.gethostbyname(socket.gethostname())
            iplist = iphead.split('.')
            iphead = '.'.join(iplist[:-1]) + '.'
            # time.sleep(0.1)
            # s.setblocking(False)
            try:
                IP = iphead+ str(i)
                print('try:'+IP)
                s.connect((IP, port))
                s.send(' '.encode())
                serverIP = IP
                break
            except:
                continue

    # 更新组件
    def update_widgets(self, connectList, text):
        s=''
        for i in range(len(connectList)):
            s += connectList[i] +'\n'
        self.listLabel.setText(s)

        # s = self.trText.toPlainText()
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        s = t  + ' from '+ text
        # self.trText.setText(s)
        self.trText.append(s)
        self.loadData()
        
    def sendTr(self):
        if serverIP == '':
            self.barLabel.setText('尚未确定主机IP，请扫描！')
            return
        text = self.trText.toPlainText()
        if text.strip() == '':
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('无内容，拒绝提交！')
            return
        s = socket.socket()
        s.settimeout(1)
        c=''
        try:
            s.connect((serverIP, port))
            s.send(text.encode())
            c = s.recv(1024).decode()
            if 'ok' == c:
                self.write(text)
                self.barLabel.setText('恭喜，提交成功！')
                self.barLight.setPixmap(self.icon_right)
                self.trText.clear()
                self.loadData()
                return
        except:
            pass
        self.barLight.setPixmap(self.icon_error)
        self.barLabel.setText('提交失败，请重试！')


    def state(self):
        text = self.trText.toPlainText()
        if '' != text.strip():
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('有改动，尚未提交')

    def openDate(self):
        try:
            # os.system('excel.exe -readOnly '+filePath)
            subprocess.Popen('excel.exe -readOnly '+filePath)
        except:
            self.barLabel.setText('Open data failed' + filePath)

    def loadData(self):
        print('loading data...')
        try:
            excel = xlrd.open_workbook(filePath)
            sheet = excel.sheet_by_index(0)
            row = sheet.nrows
            for i in range(row):
                if(self.table.rowCount()==i):
                    self.table.insertRow(i)
                for j in range(3):
                    item = QTableWidgetItem(sheet.cell(i, j).value)
                    self.table.setItem(i, j, item)
        except:
            self.barLabel.setText('Load data error')
        pass

    # 将接收的TR记录至Excel文件中
    def write(self,text):
        readFile = xlrd.open_workbook(filePath)
        read_sheet = readFile.sheet_by_index(0)
        row = read_sheet.nrows
        writeFile = copy(readFile)
        write_sheet = writeFile.get_sheet(0)
        number = 'ARVS-' + str(row)
        write_sheet.write(row, 0, number)
        write_sheet.write(row, 1, serverIP + '(' + socket.gethostname() + ')')
        write_sheet.write(row, 2, text)
        writeFile.save(filePath)
        readFile.release_resources()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = window()
    w.show()
    sys.exit(app.exec_())
