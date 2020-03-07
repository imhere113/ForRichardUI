#import xlrd
import json
import os
from multiprocessing import Process
from openpyxl import Workbook
import sip
from PyQt5.QtWidgets import (QWidget, QApplication, QLabel, QLineEdit, QGridLayout, QPushButton,
                             QVBoxLayout, QMessageBox, QProgressBar, QLayout)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import *
import multiprocessing
import time
import configparser





class MyWidget(QWidget):
    key = ''
    num = ''
    site = ''
    def __init__(self, parent=None):
        super(MyWidget, self).__init__(parent)

        # self.resize(800, 500)
        self.setWindowTitle('ForRichard 1.0')
        self.KeyEdit = QLineEdit()
        self.NumEdit = QLineEdit()
        self.SiteEdit = QLineEdit()
        self.pbar = QProgressBar()
        self.pbar.setGeometry(30, 40, 200, 25)
        #self.emationRes = Qlabele()
        Key = QLabel('key')
        Num = QLabel('Num')
        Site = QLabel('site')
        tips = QLabel('tips:这个版本还有些bug,待优化')


        leftLayout = QGridLayout()
        leftLayout.addWidget(Key, 0, 0)
        leftLayout.addWidget(self.KeyEdit, 0, 1)
        leftLayout.addWidget(Num, 1, 0)
        leftLayout.addWidget(self.NumEdit, 1, 1)
        leftLayout.addWidget(Site, 2, 0)
        leftLayout.addWidget(self.SiteEdit, 2, 1)

        #leftLayout.addWidget(self.pbar, 3, 0)
        leftLayout.setColumnStretch(0, 1)
        leftLayout.setColumnStretch(1, 3)

        self.ok_button = QPushButton("Start", self)
        # self.closePushButton = QPushButton("Close", self)

        rightLayout = QVBoxLayout()
        #rightLayout.set setMargin(10)
        rightLayout.addStretch(7)
        rightLayout.addWidget(self.ok_button)
        # rightLayout.addWidget(self.closePushButton)

        mainLayout = QGridLayout(self)
        #mainLayout.setMargin(15)
        mainLayout.setSpacing(15)
        mainLayout.addLayout(leftLayout, 0, 0)
        mainLayout.addLayout(rightLayout, 0, 1)
        mainLayout.setSizeConstraint(QLayout.SetFixedSize)
        mainLayout.addWidget(self.pbar, 1, 0)
        mainLayout.addWidget(tips, 2, 0)
        self.ok_button.clicked.connect(lambda :self.on_ok_button_clicked())
        #self.connect(self.closePushButton, QtCore.SIGNAL("clicked()"), self, QtCore.SLOT("close()"))


    def json2excel(self, jsfile, excfile):
        # 读取json数据
        wb = Workbook()
        ws = wb.active
        cols = []
        if os.path.exists(jsfile):
            # 先用key值写表头
            with open(jsfile, 'r', encoding='UTF-8') as fp:
                # 先用key值写表头
                while len(cols) == 0:
                    line = fp.readline()

                    if not line:
                        print("没有内容")
                    elif line == '[\n' or line == ']':
                        print("抬头or结尾")
                    else:
                        # 每一行转换成字典类型
                        if line[-2] == ',':
                            line = line[:-2]
                        else:
                            line = line[:-1]
                        jsdata = json.loads(line)
                        # 用key值做标题
                        for k in jsdata.keys():
                            if k not in cols:
                                cols.append(k)
                        ws.append(cols)  # 标题
            # 写值
            with open(jsfile, 'r', encoding='utf8') as fp:
                # 循环写值
                a = 10
                while True:
                    # print('正在写入的行数%s：' % a)
                    line = fp.readline()
                    if not line:
                        break
                    # 转换为python对象
                    elif line == '[\n' or line == ']':
                        print('useless')
                    else:
                        if line[-2] == ',':
                            line = line[:-2]
                        else:
                            line = line[:-1]

                        jsdata = json.loads(line)
                        rowdata = []
                        for col in cols:
                            # 获取每一行key值对应的value值
                            rowdata.append(jsdata.get(col))
                        ws.append(rowdata)  # 写行
                    a += 1
                    self.pbar.setValue(a)
                    # ws.append(cols) # 标题
        print('保存中')
        wb.save(excfile)  # 保存

    def ResSlot(self, res):
        jsfile = "item.json"
        excfile = self.key + ".xlsx"
        if os.path.exists(excfile):
            os.remove(excfile)
        self.json2excel(jsfile, excfile)
        QMessageBox.question(self, "提示", "本次爬虫已经结束",
                             QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Ok)
        self.pbar.setValue(0)
        #self.emationRes.setText(res)
    # 自定义实现爬虫的槽函数
    def on_ok_button_clicked(self):
        self.key = self.KeyEdit.text()
        self.num = self.NumEdit.text()
        self.site = self.SiteEdit.text()
        self.pbar.setMinimum(0)
        self.pbar.setMaximum(int(self.num) + 10)
        self.pbar.setValue(10)
        # 执行爬虫
        self.my_thread = MyThread(key=self.key, num=self.num, site=self.site)  # 实例化线程对象
        self.my_thread.resSignal.connect(self.ResSlot)
        self.my_thread.start()



# 后台程序
class MyThread(QThread):  # 继承QThread
    resSignal = pyqtSignal(str)  # 注册一个信号

    def __init__(self, parent=None, key="", num="", site=""):  # 从前端界面中传递参数到这个任务后台
        super(MyThread, self).__init__(parent)
        self.key = key
        self.num = num
        self.site = site

    def run(self):  # 重写run  比较耗时的后台任务可以在这里运行

        if os.path.exists('ForRichard.cfg'):
            fp = configparser.ConfigParser()
            fp.read('ForRichard.cfg')
            fp.set("Richard", "key", self.key)
            fp.set("Richard", "num", self.num)
            fp.set("Richard", "site", self.site)
            fp.write(open("ForRichard.cfg", "w"))

        #print('⽗进程 %d.' % os.getpid())
        p = Process(target=run_proc)
        print('child go')
        p.start()
        p.join()
        print('child done')
        self.Resematin = 'ok'
        self.resSignal.emit(self.Resematin)  # 任务完成后，发送信号

def run_proc():
    print('child in')
    if os.path.exists('items.json'):
        os.remove('items.json')
    ret = os.system('gogogo.exe')
    #execute(["scrapy", "crawl", "ForRichard", "-a", str1, "-a", str2, "-o", "items.json"])

if __name__ == "__main__":
    import sys

    multiprocessing.freeze_support()
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    w = MyWidget()
    w.show()
    app.exec_()






