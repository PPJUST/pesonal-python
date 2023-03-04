import os
import tkinter
from tkinter import filedialog
from PySide2.QtCore import QFile
from PySide2.QtWidgets import QApplication
from PySide2.QtUiTools import QUiLoader
import datetime
import sys
import shutil
from natsort import natsorted
import random


class Quickmove:
    def __init__(self):
        super(Quickmove, self).__init__()

        # 设置UI文件只读
        qfile = QFile("ui.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()

        # 加载UI
        self.ui = QUiLoader().load(qfile)

        # 初始化要用的变量
        self.paths = {}     # 建立一个空字典
        self.config_load()
        self.ui.line_edit_path_old.setText(self.paths["path_old"])
        self.ui.line_edit_path_move_1.setText(self.paths["path_move_1"])
        self.ui.line_edit_path_move_2.setText(self.paths["path_move_2"])
        self.ui.line_edit_path_move_3.setText(self.paths["path_move_3"])
        self.ui.line_edit_path_move_4.setText(self.paths["path_move_4"])
        self.ui.line_edit_path_move_5.setText(self.paths["path_move_5"])
        self.select_model = "文件模式"      # 默认遍历模式
        self.move_number = 0        # 默认第1个文件

        # 连接信号与槽函数
        self.ui.button_ask_path_old.clicked.connect(self.ask_path_old)
        self.ui.button_ask_path_move_1.clicked.connect(self.ask_path_move_1)
        self.ui.button_ask_path_move_2.clicked.connect(self.ask_path_move_2)
        self.ui.button_ask_path_move_3.clicked.connect(self.ask_path_move_3)
        self.ui.button_ask_path_move_4.clicked.connect(self.ask_path_move_4)
        self.ui.button_ask_path_move_5.clicked.connect(self.ask_path_move_5)
        self.ui.button_makesure.clicked.connect(self.makesure)
        self.ui.radio_button_file.clicked.connect(self.check_model)
        self.ui.radio_button_folder.clicked.connect(self.check_model)
        self.ui.button_move_1.clicked.connect(self.move_1)
        self.ui.button_move_2.clicked.connect(self.move_2)
        self.ui.button_move_3.clicked.connect(self.move_3)
        self.ui.button_move_4.clicked.connect(self.move_4)
        self.ui.button_move_5.clicked.connect(self.move_5)
        self.ui.button_quit.clicked.connect(self.quit_button)
        self.ui.button_open_old.clicked.connect(self.open_old)
        self.ui.text_info.textChanged.connect(self.scroll)
        # 每次文本行有变动则更新配置文件
        self.ui.line_edit_path_old.textChanged.connect(self.config_save)
        self.ui.line_edit_path_move_1.textChanged.connect(self.config_save)
        self.ui.line_edit_path_move_2.textChanged.connect(self.config_save)
        self.ui.line_edit_path_move_3.textChanged.connect(self.config_save)
        self.ui.line_edit_path_move_4.textChanged.connect(self.config_save)
        self.ui.line_edit_path_move_5.textChanged.connect(self.config_save)

    # 编写槽函数
    def config_load(self):
        """读取配置文件"""
        with open("config.json", "r", encoding="utf-8") as cr:
            line = cr.readlines()       # 按行读取，分配给字典
            self.paths["path_old"] = line[0].strip()
            self.paths["path_move_1"] = line[1].strip()
            self.paths["path_move_2"] = line[2].strip()
            self.paths["path_move_3"] = line[3].strip()
            self.paths["path_move_4"] = line[4].strip()
            self.paths["path_move_5"] = line[5].strip()
        self.ui.text_info.insertPlainText("\n" + self.get_time() + '已读取配置文件')

    def config_save(self):
        """保存配置文件"""
        with open("config.json", "w", encoding="utf-8") as cw:
            for key in self.paths:
                cw.write(self.paths[key].strip() + "\n")        # 逐行写入配置文件
        self.ui.text_info.insertPlainText("\n" + self.get_time() + '已保存配置文件')

    def ask_path(self):
        """选取文件夹路径"""
        root = tkinter.Tk()
        root.withdraw()  # 用来隐藏窗口
        path_check = filedialog.askdirectory()
        if path_check == "":
            self.path = "选择路径"
        else:
            self.path = path_check
        self.ui.text_info.insertPlainText("\n" + self.get_time() + '已选择文件夹')
        return self.path

    def ask_path_old(self):
        """选择源文件夹按钮"""
        self.ask_path()
        self.paths["path_old"] = self.path
        self.ui.line_edit_path_old.setText(self.paths["path_old"])      # 赋值给字典

    def ask_path_move_1(self):
        """选择移动文件夹按钮"""
        self.ask_path()
        self.paths["path_move_1"] = self.path
        self.ui.line_edit_path_move_1.setText(self.paths["path_move_1"])

    def ask_path_move_2(self):
        """选择移动文件夹按钮"""
        self.ask_path()
        self.paths["path_move_2"] = self.path
        self.ui.line_edit_path_move_2.setText(self.paths["path_move_2"])

    def ask_path_move_3(self):
        """选择移动文件夹按钮"""
        self.ask_path()
        self.paths["path_move_3"] = self.path
        self.ui.line_edit_path_move_3.setText(self.paths["path_move_3"])

    def ask_path_move_4(self):
        """选择移动文件夹按钮"""
        self.ask_path()
        self.paths["path_move_4"] = self.path
        self.ui.line_edit_path_move_4.setText(self.paths["path_move_4"])

    def ask_path_move_5(self):
        """选择移动文件夹按钮"""
        self.ask_path()
        self.paths["path_move_5"] = self.path
        self.ui.line_edit_path_move_5.setText(self.paths["path_move_5"])

    def makesure(self):
        """确认路径，遍历文件"""
        self.move_number = 0
        self.files = []
        self.folders = []
        path_travel = os.listdir(self.paths["path_old"])  # 遍历后的文件路径
        for i in path_travel:  # for循环遍历
            total_path = self.paths["path_old"] + "/" + i   # 组合完整文件名
            if os.path.isdir(total_path):       # 判断是否是文件夹，放入对于列表
                self.folders.append(total_path)
            else:
                self.files.append(total_path)

        self.files = natsorted(self.files, key=lambda x: x.lower())     # 在natsorted中加入参数使之忽略大小写影响
        self.folders = natsorted(self.folders, key=lambda x: x.lower())  # 在natsorted中加入参数使之忽略大小写影响

        # 确认要移动的文件类型
        if self.select_model == "文件模式":
            need_moves = self.files
        else:
            need_moves = self.folders
        # 保存要移动的文件到配置文件
        with open("configf.json", "w", encoding="utf-8") as fw:
            for i in need_moves:
                fw.write(i.strip() + "\n")        # 逐行写入配置文件
        # 先显示第一个文件
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
            try:
                self.ui.label_show_file.setText(line[0])
                self.ui.text_info.insertPlainText("\n" + self.get_time() + '确认文件路径，已将文件目录写入配置文件')
            except IndexError:
                tkinter.messagebox.showwarning(title='注意', message='当前路径没有文件/文件夹')

    def check_model(self):
        if self.ui.radio_button_file.isChecked():
            self.select_model = "文件模式"
        else:
            self.select_model = "文件夹模式"

    def move_1(self):
        """移动到文件夹"""
        self.strat_move("path_move_1")

    def move_2(self):
        """移动到文件夹"""
        self.strat_move("path_move_2")

    def move_3(self):
        """移动到文件夹"""
        self.strat_move("path_move_3")

    def move_4(self):
        """移动到文件夹"""
        self.strat_move("path_move_4")

    def move_5(self):
        """移动到文件夹"""
        self.strat_move("path_move_5")

    def strat_move(self, path_move_number):
        """移动文件夹操作，需要一个目标文件夹路径的变量"""
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)  # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:  # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            move_files = os.listdir(self.paths[path_move_number])  # 检查目标文件夹下的文件，是否和要移动的文件重复
            if os.path.split(line[self.move_number].strip())[1] in move_files:
                new_name = os.path.splitext(line[self.move_number].strip())[0] + '【重复' + str(random.randint(1, 100)) + '】' + \
                           os.path.splitext(line[self.move_number].strip())[1]
                os.renames(line[self.move_number].strip(), new_name)
                shutil.move(new_name, self.paths[path_move_number])
            else:
                shutil.move(line[self.move_number].strip(), self.paths[path_move_number])
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText(
            "\n" * 2 + self.get_time() + "完成文件移动： " + line[self.move_number - 1].strip() + " >>> " + self.paths[
                path_move_number])
        # 是否自动打开下一个文件
        if self.ui.check_box_open_next.isChecked() == True: # 检查勾选框状态
            if self.select_model == "文件模式":
                os.startfile(line[self.move_number].strip())
            elif self.select_model == "文件夹模式":
                os.startfile(line[self.move_number].strip() + "/" + natsorted(os.listdir(line[self.move_number].strip()))[0])


    def get_time(self):
        """获取当前时间"""
        tm = str(datetime.datetime.now())[:-7].replace(':', '.') + ":"
        return tm

    def quit_button(self):
        sys.exit(1)

    def open_old(self):
        """打开源文件夹"""
        os.startfile(self.paths["path_old"])

    def scroll(self):
        """文本框下拉到底"""
        self.ui.text_info.verticalScrollBar().setValue(self.ui.text_info.verticalScrollBar().maximum())


def main():
    app = QApplication()
    quickmove = Quickmove()
    quickmove.ui.show()
    app.exec_()

if __name__ == "__main__":
    main()
