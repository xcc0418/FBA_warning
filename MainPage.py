from view import *
import tkinter.font as tkFont


class MainPage(object):
    def __init__(self, master=None):
        self.root = master  # 定义内部变量root
        # self.root.geometry('%dx%d' % (600, 500))  # 设置窗口大小
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.ft = tkFont.Font(family='microsoft yahei', size=11)
        self.createPage()

    def createPage(self):
        self.checking = FBA_order(self.root)
        self.exchange = FBA_warning(self.root)
        # self.pair = Paie_fnsku(self.root)
        # Calendar(self.root)
        self.fba_order()  # 默认显示数据录入界面
        menubar = Menu(self.root)
        menubar.add_command(label='FBA库存处理', command=self.fba_order, font=self.ft)
        menubar.add_command(label='更新FBA库存', command=self.exchange_fnsku, font=self.ft)
        # menubar.add_command(label='配对', command=self.pair_fnsku, font=self.ft)
        self.root['menu'] = menubar  # 设置菜单栏

    def fba_order(self):
        self.root.geometry('%dx%d' % (1200, 800))
        self.checking.pack()
        self.exchange.pack_forget()
        # self.pair.pack_forget()

    def exchange_fnsku(self):
        self.root.geometry('%dx%d' % (350, 150))
        self.exchange.pack()
        # self.pair.pack_forget()
        self.checking.pack_forget()
    #
    # def pair_fnsku(self):
    #     self.root.geometry('%dx%d' % (550, 200))
    #     self.pair.pack()
    #     self.exchange.pack_forget()
    #     self.checking.pack_forget()