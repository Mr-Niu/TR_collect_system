from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import subprocess
import xlrd
from xlutils.copy import copy
import os
import sys
import shutil
import socket
import time
import threading
import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("client")

dataPath = 'Tr_client.xls'
path = os.getcwd()
filePath = path + '\\data\\' + dataPath
serverIP = '10.20.220.87'
port = 6667
sendPort = 8888

class window(QMainWindow, QWidget):
    def __init__(self, parent = None):
        super(window, self).__init__(parent)
        self.icon = QIcon('./pics/client2.png')
        self.icon_right = QPixmap('./pics/right.png')
        self.icon_error = QPixmap('./pics/error.png')
        self.attachList = []
        self.fileList = []
        self.oldPath = '.'
        self.num = ''
        self.isRun = False
        self.createUi()
        self.connect()
        # self.process = QProgressDialog(self)
        # self.process.setModal(True)
        # self.process.setRange(0,100)
        # self.process.setWindowFlags(Qt.WindowTitleHint|Qt.WindowCloseButtonHint)

    def createUi(self):
        # 窗口
        self.setWindowTitle('TR记录客户端')
        self.setWindowIcon(self.icon)
        # 状态栏
        self.statusbar = QStatusBar()
        self.barLight = QLabel()
        self.barLight.setPixmap(self.icon_error)
        self.barLabel = QLabel()
        self.processBar = QProgressDialog()
        self.processBar.setRange(0,100)
        self.processBar.setModal(True)
        self.processBar.setCancelButton(None)
        self.statusbar.addWidget(self.barLight)
        self.statusbar.addWidget(self.barLabel)
        self.statusbar.addWidget(self.processBar)
        self.setStatusBar(self.statusbar)

        # 启动、停止、打开； 服务器信息
        self.scanButton = QPushButton()
        self.scanButton.setText('扫描服务器(&S)')
        self.reloadButton =QPushButton()
        self.reloadButton.setText('加载数据库(&L)')
        self.openDateButton = QPushButton('打开数据库(&O)')
        self.ipLabel = QLabel('服务器IP及端口：')
        self.ipLabel2 = QLabel('10.20.220.87:6667')
        font = QFont('Arial')
        font.setBold(True)
        self.ipLabel2.setFont(font)

        # 连接的客户端及事件
        self.table = QTableWidget(10, 4)
        self.table.setHorizontalHeaderLabels(['序号', '提交IP', '内容', '附件'])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.textArea = QLabel('TR 内容：')
        #self.textArea.setFixedHeight(150)
        self.trText = QTextEdit() # TR文本框
        self.trText.setMaximumHeight(150)
        font = QFont()
        font.setPixelSize(24)
        self.trText.setFont(font)

        self.attachArea = QLabel('上传附件：')
        self.attachmentText = QTextEdit()
        self.attachmentText.setFixedHeight(60)
        self.attachButton = QPushButton('附\n件')
        self.attachButton.setFixedWidth(30)
        self.attachButton.setFixedHeight(60)
        self.sendButton = QPushButton()
        self.sendButton.setText('提交至服务器(&P)')
        # self.sendButton.setShortcut(QKeySequence(Qt.ControlModifier+ Qt.Key_Enter))

        #布局
        self.hLayout1 = QHBoxLayout()
        # self.hLayout1.addWidget(self.scanButton)
        self.hLayout1.addWidget(self.reloadButton)
        self.hLayout1.addWidget(self.openDateButton)
        self.hLayout1.addStretch(1)

        self.hLayout2 = QHBoxLayout()
        self.hLayout2.addWidget(self.ipLabel)
        self.hLayout2.addWidget(self.ipLabel2)
        self.hLayout2.addStretch(1)


        self.vLayout_table_event = QVBoxLayout()
        self.hLayoutAttach = QHBoxLayout()
        self.hLayout3 = QHBoxLayout()
        self.vLayout_table_event.addWidget(self.table)
        self.vLayout_table_event.addWidget(self.textArea)
        self.vLayout_table_event.addWidget(self.trText)
        self.vLayout_table_event.addWidget(self.attachArea)
        self.hLayoutAttach.addWidget(self.attachmentText)
        self.hLayoutAttach.addWidget(self.attachButton)
        self.vLayout_table_event.addLayout(self.hLayoutAttach)
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
        self.attachButton.clicked.connect(self.selectFiles)
        self.sendButton.clicked.connect(self.sendTr)
        self.trText.textChanged.connect(self.state)

    def scanIp(self):
        global serverIP
        self.barLight.setPixmap(self.icon_error)
        self.barLabel.setText('正在扫描服务器IP...')
        for i in range(32):
            threading.Thread(target=self.getIp, args=(i*8,)).start()
        for i in range(30):
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
        for i in range(start, start + 8):
            s = socket.socket()
            s.settimeout(0.1)
            iphead = socket.gethostbyname(socket.gethostname())
            iplist = iphead.split('.')
            iphead = '.'.join(iplist[:-1]) + '.'
            # time.sleep(0.1)
            # s.setblocking(False)
            # iphead = '10.20.220.'
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

    # 设定目录文件以确定上传文件
    def selectFiles(self):
        self.attachList = QFileDialog.getOpenFileNames( self, '选择附件', self.oldPath)[0]
        self.oldPath = os.path.split(self.attachList[0])[0]
        for i in self.attachList:
            self.attachmentText.append(i +';')
        # self.attachmentText.setText(files)
        self.barLabel.setText('成功添加附件')


    def sendTr(self):
        self.getFile()
        if serverIP == '':
            self.barLabel.setText('尚未确定主机IP，请扫描！')
            return
        text = self.trText.toPlainText()
        if text.strip() == '':
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('无内容，拒绝提交！')
            return
        self.write(text)
        s = socket.socket()
        # s.settimeout(1)
        #s.setblocking(False)
        try:
            s.connect((serverIP, port))
            s.send(text.encode())
            c = s.recv(1024).decode()
            print('received c is: ' + c)
            # self.barLabel.setText('恭喜，提交成功！')
            self.barLabel.setText('开始发送附件')
            ok = self.sendFiles()
            if 'ok' == c and ok:
                self.barLabel.setText('恭喜，提交成功！')
                self.barLight.setPixmap(self.icon_right)
                self.trText.clear()
                self.loadData()
                return
        except:
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('提交失败，请重试！')
        s.close()
        self.loadData()



    def state(self):
        text = self.trText.toPlainText()
        if '' != text.strip():
            self.barLight.setPixmap(self.icon_error)
            self.barLabel.setText('有改动，尚未提交')

    def openDate(self):
        try:
            # os.system('excel.exe -readOnly '+filePath)
            subprocess.Popen('excel.exe -readOnly ' + filePath)
        except:
            self.barLabel.setText('打开失败，请确认是否已将Excel添加至系统环境变量')

    def loadData(self):
        print('loading data...')
        try:
            excel = xlrd.open_workbook(filePath)
            sheet = excel.sheet_by_index(0)
            row = sheet.nrows
            for i in range(self.table.rowCount()):
                for j in range(4):
                    item = QTableWidgetItem('')
                    self.table.setItem(i, j, item)
            for i in range(row):
                if(self.table.rowCount()==i):
                    self.table.insertRow(i)
                for j in range(4):
                    try:
                        d = sheet.cell(i,j).value
                    except:
                        d=''
                    item = QTableWidgetItem(d)
                    self.table.setItem(i, j, item)
        except:
            self.barLabel.setText('Load data error')
        

    # 将接收的TR记录至Excel文件中
    def write(self,text):
        readFile = xlrd.open_workbook(filePath)
        read_sheet = readFile.sheet_by_index(0)
        row = read_sheet.nrows
        writeFile = copy(readFile)
        write_sheet = writeFile.get_sheet(0)
        number = 'ARVS-' + str(row)
        self.num = number
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        write_sheet.write(row, 0, number)
        write_sheet.write(row, 1, serverIP + '(' + socket.gethostname() + ') at ' + t)
        # 写入附件内容
        write_sheet.write(row, 2, text)
        attach=''
        for i in self.fileList:
            attach += os.path.split(i)[1] +'; '
        write_sheet.write(row, 3, attach)
        writeFile.save(filePath)
        readFile.release_resources()

    def sendFiles(self):
        try:
            client = socket.socket()
            # client.settimeout(3)
            client.connect((serverIP, sendPort))
        except:
            self.barLabel.setText('附件发送失败')
            return 0

        print('开始发送')

        if self.fileList == []:
            client.send('finished:-2'.encode())
            client.recv(1024)
            self.barLabel.setText('发送完成.')
            client.close()
            return 1
        try:
            os.mkdir('./data/' + self.num)
        except:
            pass
        for file in self.fileList:
            file = file.strip('\n')
            filename = os.path.split(file)[1]
            self.barLabel.setText('准备发送文件：'+filename)
            f = open(file, 'rb')
            self.barLabel.setText('正在发送：' + filename)
            size = os.path.getsize(file)

            dest = os.getcwd() + '\\data\\'+ self.num + '\\' + filename
            # commond = 'copy ' + file + ' ./data/'+ self.num + '/' + filename
            # os.system(commond.replace('/', '\\'))
            shutil.copyfile(file, dest)
            title = filename + ':'+str(size)
            print(title)
            client.send(title.encode())
            client.recv(1024)
            print('=============')
            cursize=0
            while True:
                c = f.read(1024)
                if not c:
                    break
                client.send(c)
                client.recv(1024)
                cursize += len(c)
                if cursize % 1024*1024 == 0:
                    self.processBar.setValue(int(cursize/size*100)+1)
            self.processBar.setValue(100)
            f.close()

        client.send('finished:-2'.encode())
        client.recv(1024)
        client.close()
        self.barLabel.setText('发送完成.')

        self.attachmentText.clear()
        return 1
    def getFile(self):
        s = self.attachmentText.toPlainText()
        file = ''
        fileList = []
        for c in s:
            if ';' != c:
                file += c
                continue
            fileList.append(file)
            file = ''
        self.fileList = fileList


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = window()
    w.show()
    sys.exit(app.exec_())
