from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QPushButton, QInputDialog, QLabel, QGridLayout, QTextEdit, QFileDialog
from PyQt5 import QtCore
import sys, word;


class WDPWindos(QWidget):
    def __init__(self):
        super().__init__()
        self.ChuangJianChuangKou()

    def ChuangJianChuangKou(self):
        self.setGeometry(50, 50, 1280, 720)
        # pathbutton1 = QPushButton('输入文档1的文件路径', self)
        # pathbutton2 = QPushButton('输入文档2的文件路径', self)
        # filenamebutton1 = QPushButton('输入文档1的文件名', self)
        # filenamebutton2 = QPushButton('输入文档2的文件名', self)
        # pathbutton1.clicked.connect(self.path1)
        # pathbutton2.clicked.connect(self.path2)
        # filenamebutton1.clicked.connect(self.filename1)
        # filenamebutton2.clicked.connect(self.filename2)
        # self.d1p = QLabel('未输入文档1文件路径')
        # self.d2p = QLabel('未输入文档2文件路径')
        # self.d1n = QLabel('未输入文档1的文件名')
        # self.d2n = QLabel('未输入文档2的文件名')

        v2button1 = QPushButton('选择第一个文档', self)
        v2button2 = QPushButton('选择第二个文档', self)
        v2button1.clicked.connect(self.v2file1)
        v2button2.clicked.connect(self.v2file2)
        self.lb1 = QLabel('未选择文档，请选择第一个文档')
        self.lb2 = QLabel('未选择文档，请选择第二个文档')


        compbutton = QPushButton('开始对比', self)
        compbutton.clicked.connect(self.startComp)

        self.messg = QTextEdit()
        self.messg.setText('本程序用于比较两个word文档的重复内容，只支持.docx格式。'+'\n'+'按左边的按钮选择两个word文档，输入完毕后点击”开始对比“，对比结果将显示在此文本框。'
                           +'\n'+'现版本十分粗糙，没有做任何异常处理。'+'\n'+'文档对比核心部分代码参考CSDN用户’赫伟博士‘的文章：https://haolaoshi.blog.csdn.net/article/details/103798581，此程序在文章基础上做了图形界面并完成打包')
        self.messg.setFocusPolicy(QtCore.Qt.NoFocus)

        grid = QGridLayout()
        grid.setSpacing(10)
        # grid.addWidget(pathbutton1, 1, 0)
        # grid.addWidget(self.d1p, 1, 1)
        # grid.addWidget(filenamebutton1, 2, 0)
        # grid.addWidget(self.d1n, 2, 1)
        # grid.addWidget(pathbutton2, 3, 0)
        # grid.addWidget(self.d2p, 3, 1)
        # grid.addWidget(filenamebutton2, 4, 0)
        # grid.addWidget(self.d2n, 4, 1)
        grid.addWidget(compbutton, 5, 0)
        grid.addWidget(self.messg, 1, 1, 4, 9)
        grid.addWidget(v2button1, 1, 0)
        grid.addWidget(self.lb1, 2, 0)
        grid.addWidget(v2button2, 3, 0)
        grid.addWidget(self.lb2, 4, 0)

        self.setLayout(grid)
        self.show()

    # def path1(self):
    #     do1path, ok = QInputDialog.getText(self, '文档1路径', '请输入文档1的文件路径，如 D:\Document/Work/'+'\n'+'可找到文件位置，然后在文件资源管理器地址栏复制')
    #     if ok:
    #         self.d1p.setText(str(do1path))
    #
    # def path2(self):
    #     do2path, ok = QInputDialog.getText(self, '文档2路径', '请输入文档2的文件路径，如 D:\Document/Work/'+'\n'+'可找到文件位置，然后在文件资源管理器地址栏复制')
    #     if ok:
    #         self.d2p.setText(str(do2path))
    #
    # def filename1(self):
    #     do1filename, ok = QInputDialog.getText(self, '文档1文件名', '请输入文档1的文件名（不包括后缀".docx"）')
    #     if ok:
    #         self.d1n.setText(str(do1filename) + '.docx')
    #
    # def filename2(self):
    #     do2filename, ok = QInputDialog.getText(self, '文档2文件名', '请输入文档2的文件名（不包括后缀".docx"）')
    #     if ok:
    #         self.d2n.setText(str(do2filename) + '.docx')

    def v2file1(self):
        self.lb1.setText(str(QFileDialog.getOpenFileName(self, '选择第一个文档', '/', '*.docx')[0]))

    def v2file2(self):
        self.lb2.setText(str(QFileDialog.getOpenFileName(self, '选择第二个文档', '/', '*.docx')[0]))

    def startComp(self):
        t1 = str(self.lb1.text())
        t2 = str(self.lb2.text())
        word.wordrun(t1, t2, self.messg)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WDPWindos()
    sys.exit(app.exec_())