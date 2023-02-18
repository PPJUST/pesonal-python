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
from pypinyin import lazy_pinyin


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
        # 下面代码创建字典，文件名与其大写文件名对应，排序后再变回去，去除大小写影响
        files_dict = {}
        files_upper = []
        files_reup = []
        for i in self.files:
            files_dict["".join(lazy_pinyin(i.upper()))] = i     # 先全转大写，在利用pypinyin将汉字转拼音并连接起来
            files_upper.append("".join(lazy_pinyin(i.upper()))) # 将转换后的结果放到一个单独的列表
        files_upper = natsorted(files_upper)   # 使用natsorted对新列表排序，这和Windows默认排序相同，sort排序则会不一样
        for y in files_upper:
            files_reup.append(files_dict[y])    # 最后将新排序列表的key值和value值对应起来，得到Windows默认排序的列表
        self.files = files_reup

        folders_dict = {}
        folders_upper = []
        folders_reup = []
        for i in self.folders:
            folders_dict["".join(lazy_pinyin(i.upper()))] = i
            folders_upper.append("".join(lazy_pinyin(i.upper())))
        folders_upper = natsorted(folders_upper)  # 使用natsorted排序，这和Windows默认排序相同，sort排序则会不一样
        for y in folders_upper:
            folders_reup.append(folders_dict[y])
        self.folders = folders_reup

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
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)     # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:      # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            shutil.move(line[self.move_number].strip(), self.paths["path_move_1"])      # 利用shutil移动文件
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成文件移动：" + line[self.move_number].strip() + " >>> " + self.paths["path_move_1"])

    def move_2(self):
        """移动到文件夹"""
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)     # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:      # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            shutil.move(line[self.move_number].strip(), self.paths["path_move_2"])
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成文件移动：" + line[self.move_number].strip() + " >>> " + self.paths["path_move_2"])

    def move_3(self):
        """移动到文件夹"""
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)     # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:      # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            shutil.move(line[self.move_number].strip(), self.paths["path_move_3"])
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成文件移动：" + line[self.move_number].strip() + " >>> " + self.paths["path_move_3"])

    def move_4(self):
        """移动到文件夹"""
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)     # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:      # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            shutil.move(line[self.move_number].strip(), self.paths["path_move_4"])
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成文件移动：" + line[self.move_number].strip() + " >>> " + self.paths["path_move_4"])

    def move_5(self):
        """移动到文件夹"""
        with open("configf.json", "r", encoding="utf-8") as fr:
            line = fr.readlines()
        move_number_max = len(line)     # 确认需要移动的总文件数量
        if self.move_number + 1 > move_number_max:      # 确认移动到第几个文件了，是否超限了
            tkinter.messagebox.showwarning(title='注意', message='已完成全部文件的移动，点击确认将重新遍历文件')
        else:
            shutil.move(line[self.move_number].strip(), self.paths["path_move_5"])
            self.move_number += 1
            self.ui.label_show_file.setText(line[self.move_number])
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成文件移动：" + line[self.move_number].strip() + " >>> " + self.paths["path_move_5"])

    def get_time(self):
        """获取当前时间"""
        tm = str(datetime.datetime.now())[:-7].replace(':', '.') + ":"
        return tm

    def quit_button(self):
        sys.exit(1)

def main():
    app = QApplication()
    quickmove = Quickmove()
    quickmove.ui.show()
    app.exec_()

if __name__ == "__main__":
    main()
