import os
from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter
from tkinter import filedialog
from PySide2.QtCore import QFile
from PySide2.QtWidgets import QApplication
from PySide2.QtUiTools import QUiLoader
import datetime
import re
import sys


# 导入ui
class DoujinshiRename:
    def __init__(self):
        super(DoujinshiRename, self).__init__()

        # 设置UI文件只读
        qfile = QFile("ui.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()

        # 加载UI
        self.ui = QUiLoader().load(qfile)

        # 初始化要用的变量
        self.read_config()  # 初始显示配置文件
        self.path = ""

        # 配置信号连接
        self.ui.button_path_get.clicked.connect(self.path_get_button)
        self.ui.button_path_open.clicked.connect(self.open_path_button)
        self.ui.button_open_help.clicked.connect(self.open_help_button)
        self.ui.button_check.clicked.connect(self.run_rename_check)
        self.ui.button_open_check.clicked.connect(self.open_check)
        self.ui.button_quit.clicked.connect(self.quit_botton)
        self.ui.button_start.clicked.connect(self.start_standard)
        self.ui.button_update_config.clicked.connect(self.update_config)


    def read_config(self):
        """读取配置文件，分配类"""
        # 首先对配置文件进行去空行
        with open("config.json", encoding="utf-8") as cr:
            crr = cr.readlines()
        with open("config.json", "w", encoding="utf-8") as cw:
            for line in crr:
                if line.split():
                    cw.write(line)

        config_all = []  # 设置个空列表
        with open("config.json", encoding="utf-8") as cj:
            config_all = cj.read().splitlines()
            line_row = 0    # 设置行，确定匹配到的关键词在第几行
            key_model, key_market, key_original, key_localization, key_other = 0, 0, 0, 0, 0
            for line in config_all:
                line_row += 1   # 每次循环+1行
                if line == "#model\n":
                    key_model = line_row
                if line == "#market":
                    key_market = line_row
                if line == "#original":
                    key_original = line_row
                if line == "#localization":
                    key_localization = line_row
                if line == "#other":
                    key_other = line_row
            self.config_model = config_all[key_model+1:key_market-1]
            self.config_market = list(tuple(config_all[key_market:key_original - 1]))   # 先转元组再转列表，完成去重
            self.config_original = list(tuple(config_all[key_original:key_localization - 1]))
            self.config_localization = list(tuple(config_all[key_localization:key_other - 1]))
            self.config_other = list(tuple(config_all[key_other:]))
            self.show_config()  # 执行方法

    def show_config(self):
        """显示配置文件中的关键词"""
        self.ui.line_edit_model.setText("".join(self.config_model))
        self.ui.text_edit_market.setText("\n".join(self.config_market))     # setText + \n.join组合列表完成多行显示列表中的元素
        self.ui.text_edit_original.setText("\n".join(self.config_original))
        self.ui.text_edit_localization.setText("\n".join(self.config_localization))
        self.ui.text_edit_other.setText("\n".join(self.config_other))
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "程序初始化完成，已读取配置文件")

    def update_config(self):
        """更新配置文件"""
        import shutil
        newname = "config_backup " + str(datetime.datetime.now())[:-7].replace(':', '.') + ".json"
        p = os.getcwd() # 获取当前路径
        shutil.copy2("config.json", p + "\\config_backup\\" + newname)     # 调用shutil模块的方法复制文件
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "已备份配置文件")

        # 设置临时变量，用于临时存放新的关键词（显示在文本框里的文本）
        new_config_model = self.ui.line_edit_model.text().split()
        new_config_market = self.ui.text_edit_market.toPlainText().split("\n")      # 先读取文本，然后split+换行符组成正常的列表
        new_config_original = self.ui.text_edit_original.toPlainText().split("\n")
        new_config_localization = self.ui.text_edit_localization.toPlainText().split("\n")
        new_config_other = self.ui.text_edit_other.toPlainText().split("\n")
        with open("config.json", "w", encoding="utf-8") as cw:      # 打开文件，重新写入5组关键词
            new_config = ["#model"] + new_config_model + \
                         ["#market"] + list(set(new_config_market)) + \
                         ["#original"] + list(set(new_config_original)) + \
                         ["#localization"] + list(set(new_config_localization)) + \
                         ["#other"] + list(set(new_config_other))     # 转集合再转列表完成去重，然后重组全部的关键词
            for line in new_config:
                cw.write(line + "\n")
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "已更新配置文件")
        # 更新配置文件后重新读取
        self.read_config()

    def path_get_button(self):
        """选取路径文件夹按钮弹窗"""
        root = tkinter.Tk()
        root.withdraw()  # 用来隐藏窗口
        self.path = filedialog.askdirectory()
        self.ui.line_edit_path.setText(self.path)  # 结果显示在文本框

    def open_path_button(self):
        """打开文件夹按钮"""
        os.startfile(self.path)
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "打开文件夹")

    def open_help_button(self):
        """打开帮助文件"""
        os.startfile("说明文档.png")
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "打开说明文档")

    def get_time(self):
        """获取当前时间"""
        tm = str(datetime.datetime.now())[:-7].replace(':', '.') + ":"
        return tm

    """改名主逻辑代码"""
    """模块1
    新建xlsx >>> 遍历文件名 >>> 保存到xlsx"""
    def xlsx_create(self):
        """新建工作簿并且写入原文件名"""
        wb = Workbook()
        try:
            wb.save("改名测试.xlsx")  # 创建xlsx
        except PermissionError:
            tkinter.messagebox.showerror(title='出错了！', message='请关闭改名测试文件后重试')
        wb = load_workbook("改名测试.xlsx")  # 加载xlsx
        ws = wb.active
        ws["A1"] = "原文件名"
        paths = os.listdir(self.path)
        for i in range(2, len(paths) + 2):  # 将遍历得到的文件名写入xlsx
            ai = "A" + str(i)
            ws[ai] = paths[i - 2]
        wb.save("改名测试.xlsx")

    """模块2
    区分括号内外元素 >>> 
    括号内元素 >>> 匹配txt中的关键词进行分类
    括号外元素 >>> 判断是压缩包还是文件夹 >>> 去除文件两端多余空格"""
    def assort(self, input_filename):
        """文件名内元素分类"""
        # 先在函数内初始化要用的变量
        self.filename_standard = ""    # 标准化后的文件名
        self.filename_inbrac = []  # 用于存放括号里的内容
        self.filename_outbrac = ""  # 用于存放括号外的内容
        self.item_market = []  # 分类：即卖会名
        self.item_original = []  # 分类：原作名
        self.item_localization = []  # 分类：汉化组
        self.item_other = []  # 分类：其他信息
        self.standard_model = self.ui.line_edit_model.text().split("@")    # 切割关键词组成模型
        filename_usetomod = input_filename

        # 尝试区分压缩包
        zip_check = False
        if os.path.splitext(filename_usetomod)[1].upper() in ['.ZIP', '.RAR', '.7Z']:    # 分割文件名后转大写，然后找压缩包后缀
            zip_check = True

        # 区分括号内外元素
        self.filename_inbrac.extend(re.findall(r"\[.*?\]|\(.*?\)|（.*?）|【.*?】", filename_usetomod))  # 正则提取括号内容，extend拼接列表
        self.filename_outbrac = re.sub(r"\[.*?\]|\(.*?\)|（.*?）|【.*?】", "", filename_usetomod)  # 正则删除括号内容

        # 分类括号内容
        # 判断括号元素是否包含关键词：即卖会名
        for line in self.config_market:
            for ele in self.filename_inbrac:
                if ele.find(line) != -1:  # find()方法不为-1说明匹配成功
                    self.item_market.append(ele)  # 将匹配成功的元素转移到对应变量
                    self.filename_inbrac.remove(ele)  # 删除已转移的括号内容
        # 判断括号元素是否包含关键词：原作名
        for line in self.config_original:  # 判断括号元素是否包含关键词：即卖会名
            for ele in self.filename_inbrac:
                if ele.find(line) != -1:  # find()方法不为-1说明匹配成功
                    self.item_original.append(ele)  # 将匹配成功的元素转移到对应变量
                    self.filename_inbrac.remove(ele)  # 删除已转移的括号内容
        # 判断括号元素是否包含关键词：汉化组
        for line in self.config_localization:  # 判断括号元素是否包含关键词：即卖会名
            for ele in self.filename_inbrac:
                if ele.find(line) != -1:  # find()方法不为-1说明匹配成功
                    self.item_localization.append(ele)  # 将匹配成功的元素转移到对应变量
                    self.filename_inbrac.remove(ele)  # 删除已转移的括号内容
        # 判断括号元素是否包含关键词：其他信息
        for line in self.config_other:  # 判断括号元素是否包含关键词：即卖会名
            for ele in self.filename_inbrac:
                if ele.find(line) != -1:  # find()方法不为-1说明匹配成功
                    self.item_other.append(ele)  # 将匹配成功的元素转移到对应变量
                    self.filename_inbrac.remove(ele)  # 删除已转移的括号内容

        # 重组文件名，设置分类
        filename_usetorecomb = input_filename
        remove_brac = self.item_market + self.item_original + self.item_localization + self.item_other    # 整合需要删除的分类元素
        for i in remove_brac:
            filename_usetorecomb = filename_usetorecomb.replace(i, "")  # 替换相关括号元素
            filename_usetorecomb = filename_usetorecomb.strip()  # 清除两端多余空格
            filename_usetorecomb = filename_usetorecomb.replace("  ", " ")  # 2个空格变1个空额清除多余空格
        if zip_check:  # 上述方法无法清除带后缀的文件 后缀与文件名之间的空格
            if filename_usetorecomb.find(" .", -5) != -1:
                filename_usetorecomb = filename_usetorecomb[:-5] + filename_usetorecomb[-5:].replace(" .", ".")  # 切片替换 .去除多余空格

        # 对照标准化文件名格式，对文件名进行元素的删改
        if zip_check == True:
            filename_usetorecomb_delzip = os.path.splitext(filename_usetorecomb)[0]    # splitext分离后缀
            filename_usetorecomb_suffix = os.path.splitext(filename_usetorecomb)[1]    # splitext提取后缀
            for i in self.standard_model:
                if i == "[社团名(作者名)]标题":
                    self.filename_standard = self.filename_standard + " " + filename_usetorecomb_delzip
                elif i == "(原作名)":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_original)
                elif i == "(即卖会名)":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_market)
                elif i == "[汉化]":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_localization)
                elif i == "[其他信息]":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_other)
            self.filename_standard = self.filename_standard + filename_usetorecomb_suffix
        else:
            for i in self.standard_model:
                if i == "[社团名(作者名)]标题":
                    self.filename_standard = self.filename_standard + " " + filename_usetorecomb
                elif i == "(原作名)":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_original)
                elif i == "(即卖会名)":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_market)
                elif i == "[汉化]":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_localization)
                elif i == "[其他信息]":
                    self.filename_standard = self.filename_standard + " " + " ".join(self.item_other)

        # 上述执行完成后可能有多余空格，继续用上面方法去除空格
        self.filename_standard = self.filename_standard.strip()  # 清除两端多余空格
        self.filename_standard = self.filename_standard.replace("  ", " ")  # 2个空格变1个空额清除多余空格
        if self.filename_standard.find(" .", -5) != -1:
            self.filename_standard = self.filename_standard[:-5] + self.filename_standard[-5:].replace(" .", ".")  # 切片替换 .去除多余空格

    """模块3
    将已经处理过的元素、文件名写入xlsx"""

    def run_rename_check(self):
        """重组文件名，设置分类"""
        # 数据写入xlsx
        self.xlsx_create()      # 首先执行之前编写的创建xlsx的函数
        wb = load_workbook("改名测试.xlsx")
        ws = wb.active
        max_row = ws.max_row
        for i in range(2, max_row + 2 - 1):
            ai = "A" + str(i)
            filename_oringinal = ws[ai].value  # 赋值于变量，用于后续操作
            self.assort(filename_oringinal)  # 执行模块2进行分类
            ws["B1"] = "标准化文件名"  # 添加标题
            ws["C1"] = "即卖会名"
            ws["D1"] = "原作名"
            ws["E1"] = "汉化"
            ws["F1"] = "其他信息"
            bi = "B" + str(i)  # 建立单元格编号
            ci = "C" + str(i)
            di = "D" + str(i)
            ei = "E" + str(i)
            fi = "F" + str(i)
            if filename_oringinal == self.filename_standard:  # 如果新文件名和原文件名相同，清空数据
                ws[bi] = ""
                ws[ci] = ""
                ws[di] = ""
                ws[ei] = ""
                ws[fi] = ""
            else:  # 如果不同，则写入数据
                ws[bi] = self.filename_standard
                ws[ci] = ",".join(self.item_market)
                ws[di] = ",".join(self.item_original)
                ws[ei] = ",".join(self.item_localization)
                ws[fi] = ",".join(self.item_other)

        # 执行查重操作，如果重复则添加前缀【】
        dup_check = []
        for i in range(2, max_row + 2 - 1):
            ai = "A" + str(i)
            dup_check.append(ws[ai].value)  # 先将所有原文件名添加到查重列表
        for i in range(2, max_row + 2 - 1):
            bi = "B" + str(i)
            dup_check.append(ws[bi].value)  # 逐行添加新文件名到查重列表
            if ws[bi].value == "":
                pass
            elif dup_check.count(ws[bi].value) > 1:  # 如果新文件名在查重列表里，添加前缀
                ws[bi] = "【重复" + str(dup_check.count(ws[bi].value) - 1) + "】" + ws[bi].value
            else:
                dup_check.append(ws[bi].value)
        wb.save("改名测试.xlsx")
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "完成改名测试，已生成xlsx文件")

    def start_standard(self):
        """正式执行改名操作"""
        n = 0  # 用于统计改名操作的次数
        wb = load_workbook("改名测试.xlsx")
        ws = wb.active
        max_row = ws.max_row  # 获取最大行数
        for i in range(2, max_row + 2):
            ai = "A" + str(i)
            bi = "B" + str(i)
            if ws[bi].value:  # 判断新文件名是否为空，不为空则执行改名
                os.rename(self.path + "/" + ws[ai].value, self.path + "/" + ws[bi].value)
                self.ui.text_info.insertPlainText("\n"*2 + self.get_time() + "执行改名：" + ws[ai].value + " >>> " + ws[bi].value)
                n += 1  # 统计操作次数
        self.ui.text_info.insertPlainText("\n"*2 + self.get_time() + "已完成全部改名操作，共执行" + str(n) + "次")

    def open_check(self):
        """打开测试文件"""
        os.startfile("改名测试.xlsx")
        self.ui.text_info.insertPlainText("\n" + self.get_time() + "打开测试文件")

    def quit_botton(self):
        sys.exit(1)


def main():
    app = QApplication()
    doujinshirename = DoujinshiRename()
    doujinshirename.ui.show()
    app.exec_()

if __name__ == "__main__":
    main()