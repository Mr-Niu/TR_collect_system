import xlrd, xlwt
from xlutils.copy import copy
import time
import threading
import socket
from PyQt5.QtCore import QObject, pyqtSignal


# 记录TR的XLS文档
file = './data/TR_data.xls'

# 本地服务器及端口
host = socket.gethostbyname(socket.gethostname())
port = 6666

# 定义server
class server(threading.Thread, QObject):
    ud = pyqtSignal(list, str)
    def __init__(self, parent = None):
        # super(server, self).__init__(parent)
        threading.Thread.__init__(self)
        super(QObject,self).__init__()
        self.mutex = threading.Lock()
        # self.s = socket.socket()


    def run(self):
        self.s = socket.socket()
        self.s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.s.bind((host, port))
        self.s.listen(10)
        self.s.setblocking(False)
        self.isStop = False
        self.listen()
    # 监听联接
    def listen(self):
        while not self.isStop:
            try:
                client, address = self.s.accept()
            except:
                time.sleep(0.5)
                continue
            t = threading.Thread(target=self.receive, args=(client, address))
            t.start()
            time.sleep(0.1)

    # 接收TR
    def receive(self, client, address):
        connectList=[]
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
        self.ud[list, str].emit(connectList, ipString + 'at ' + text)


    # 将接收的TR记录至Excel文件中
    def write(self, ipString, text):
        readFile = xlrd.open_workbook(file)
        read_sheet = readFile.sheet_by_index(0)
        row = read_sheet.nrows
        writeFile = copy(readFile)
        write_sheet = writeFile.get_sheet(0)
        number = 'ARVS-' + str(row)
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        write_sheet.write(row, 0, number)
        write_sheet.write(row, 1, ipString + ' at ' + t)
        write_sheet.write(row, 2, text)
        writeFile.save(file)
        readFile.release_resources()

    # 停止server
    def stop(self):
        self.isStop = True
        self.s.close()
        print('stopping...')





