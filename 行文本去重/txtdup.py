from PySide2.QtCore import QFile
from PySide2.QtWidgets import QApplication
from PySide2.QtUiTools import QUiLoader
import datetime
import os.path
import tkinter
from tkinter import filedialog

# 导入ui
class Txtdup:
    def __init__(self):
        super(Txtdup, self).__init__()

        # 设置UI文件只读
        qfile = QFile("ui.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()

        # 加载UI
        self.ui = QUiLoader().load(qfile)

        # 初始化要用的遍历
        self.path = os.getcwd() + "/去重.txt"  # 取得当前路径
        self.filename = "去重.txt"

        # 链接信号和槽函数
        self.ui.button_path_get.clicked.connect(self.ask_path)
        self.ui.button_path_open.clicked.connect(self.open_path)
        self.ui.button_start.clicked.connect(self.start_check)
        self.ui.button_quit.clicked.connect(self.do_quit)

    # 设置槽函数
    def ask_path(self):
        """弹窗选择文件夹路径"""
        root = tkinter.Tk()
        root.withdraw()  # 用来隐藏窗口
        self.path = filedialog.askopenfilename()  # 选择路径弹窗  #filetypes=("Text Files", "*.txt")
        self.filename = os.path.split(self.path)[1]
        self.ui.line_edit_path.setText(self.path)

    def open_path(self):
        """打开路径"""
        os.startfile(os.path.split(self.path)[0])

    def start_check(self):
        if self.ui.line_edit_path.text() == '输入文件路径，默认为当前文件夹的"去重.txt"':
            tkinter.messagebox.showwarning(title='注意', message='请选择文件路径')
        else:
            self.start_dup()

    def start_dup(self):
        """执行去重操作"""
        read_file = open(self.path, 'r', encoding='utf-8')  # 读取文件
        add_line = {''}  # 创建集合
        for line in read_file:  # 向集合中导入文本
            add_line.add(line)
        # filename_new = f'去重{datetime.datetime.now()}.txt'.replace(':', '.')
        new_filename = os.path.splitext(self.filename)[0] + str(' ') \
                       + str(datetime.datetime.now())[:-7].replace(':', '.') + '.txt'  # 设置文件名，格式为原文件名+时间
        new_path = self.path.replace(self.filename, new_filename)
        txt_dup = open(new_path, 'w', encoding='utf-8')
        for line in add_line:  # 去重后的密码写入新文件
            if line.strip():
                txt_dup.write(line)
        txt_dup.close()  # 关闭文件
        read_file.close()
        tkinter.messagebox.showinfo(title='提示', message='已完成去重操作')

    def do_quit(self):
        quit()


def main():
    app = QApplication()
    txtdup = Txtdup()
    txtdup.ui.show()
    app.exec_()

if __name__ == "__main__":
    main()
