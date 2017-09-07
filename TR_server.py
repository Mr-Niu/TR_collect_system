import xlrd, xlwt
from xlutils.copy import copy
import time
import threading
import socket
from PyQt5.QtCore import QObject, pyqtSignal
import os


# 记录TR的XLS文档
fileRecord = './data/TR_data.xls'

# 本地服务器及端口
host = socket.gethostbyname(socket.gethostname())
port = 6666+1
sendPort = 8888 #附件传送端口
connectList=['提交过TR的IP:']

# 附件传输通道


# 定义server
class server(threading.Thread, QObject):
    ud = pyqtSignal(str)
    def __init__(self, parent = None):
        # super(server, self).__init__(parent)
        threading.Thread.__init__(self)
        super(QObject,self).__init__()
        self.mutex = threading.Lock()
        self.s = socket.socket()
        self.attachSer = socket.socket()#附件传送服务器


    def run(self):
        self.s = socket.socket()
        self.s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.s.bind((host, port))
        self.s.listen(10)

        self.attachSer = socket.socket()
        self.attachSer.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.attachSer.bind((host, sendPort))
        self.attachSer.listen(10)
        # self.s.setblocking(False)

        self.isStop = False
        self.listen()
    # 监听联接
    def listen(self):
        while not self.isStop:
            try:
                client, address = self.s.accept()
            except:
                time.sleep(0.05)
                continue
            t = threading.Thread(target=self.receive, args=(client, address))
            t.start()
            time.sleep(0.1)
        # self.s.shutdown(True)
        self.s.close()
        # self.attachSer.shutdown(True)
        self.attachSer.close()

    # 接收TR
    def receive(self, client, address):
        global connectList
        clientInfo = socket.gethostbyaddr(address[0])
        ipString = address[0] + ' (' + clientInfo[0] + ')'
        text = client.recv(20480).decode().strip()
        client.send('ok'.encode())
        if '' == text.strip():
            return
        self.mutex.acquire()
        self.write(ipString, text)
        self.mutex.release()
        if ipString not in connectList:
            connectList.append(ipString)



    # 将接收的TR记录至Excel文件中
    def write(self, ipString, text):
        readFile = xlrd.open_workbook(fileRecord)
        read_sheet = readFile.sheet_by_index(0)
        row = read_sheet.nrows
        writeFile = copy(readFile)
        write_sheet = writeFile.get_sheet(0)
        number = 'ARVS-' + str(row)
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

        print('waitting for connection')
        # self.attachSer.settimeout(3)
        attClient, attAddress = self.attachSer.accept()

        # print('开始接收')
        filelist=''
        while True:
            # 接收报头
            # print('等待报头')
            try:
                title = attClient.recv(1024).decode()
                attClient.send('received'.encode())
            except:

                break
            file =''
            # 解析文件大小
            size = int(title.split(':')[1])
            if size == -2:
                # print('接收完成')
                break
            file += title.split(':')[0] + '; '
            filelist += file
           # 创建目录
            try:
                os.mkdir('./data/'+number)
            except:
                pass

            # 接收文件
            name ='./data/' + number + '/'+ title.split(':')[0]
            length = 0
            f = open(name, 'wb')
            try:
                while True:
                    # print('length: ', length, 'size: ', size)
                    if(size - length > 1024):
                        data = attClient.recv(1024)
                    else:
                        data = attClient.recv(size - length)
                    attClient.send('received'.encode())
                    f.write(data)
                    length += len(data)
                    # print(length)
                    if length >= size:
                        print('break')
                        break
            except:
                print('接收错误，重启服务器')
                self.stop()
                self.run()
            f.close()

        attClient.close()

        write_sheet.write(row, 0, number)
        write_sheet.write(row, 1, ipString + ' at ' + t)
        write_sheet.write(row, 2, text)
        write_sheet.write(row, 3, filelist)
        writeFile.save(fileRecord)
        readFile.release_resources()
        event = ipString + ': '+text +'(附件：' + filelist + ')'
        self.ud[ str].emit(event)

    # 停止server
    def stop(self):
        self.isStop = True
        # self.s.shutdown(True)
        self.s.close()
        # self.attachSer.shutdown(True)
        self.attachSer.close()
        print('stopping...')





