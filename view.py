import json
import os
from tkinter import *
from tkinter import StringVar, Label
import tkinter.font as tkFont
import datetime
import openpyxl
import pymysql
from tkinter import ttk
import FBA_inventory
from tkinter import messagebox
import global_var


class FBA_order(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master  # 定义内部变量root
        self.xscroll = Scrollbar(self, orient=HORIZONTAL)
        self.yscroll = Scrollbar(self, orient=VERTICAL)
        self.ft = tkFont.Font(family='microsoft yahei', size=10)
        self.msku = StringVar()
        self.num_msku = StringVar()
        self.sku = StringVar()
        self.country = StringVar()
        self.supplier = StringVar()
        self.itemname = StringVar()
        self.country = StringVar()
        self.fnsku = StringVar()
        self.Radiolist = IntVar()
        self.create()

    def sql(self):
        self.connection = pymysql.connect(host='3354n8l084.goho.co',  # 数据库地址
                                          port=24824,
                                          user='test_user',  # 数据库用户名
                                          password='a123456',  # 数据库密码
                                          db='storage',  # 数据库名称
                                          charset='utf8',
                                          cursorclass=pymysql.cursors.DictCursor)
        # 使用 cursor() 方法创建一个游标对象 cursor
        self.cursor = self.connection.cursor()

    def sql_close(self):
        self.cursor.close()
        self.connection.close()

    def create(self):
        Label(self).grid(row=0, stick=W, pady=10)
        self.tree = ttk.Treeview(self, show='headings', xscrollcommand=self.xscroll.set, yscrollcommand=self.yscroll.set, height=30)
        self.tree['columns'] = ('SKU', 'FNSKU', '品名', '国家', '状态')
        self.tree.column("SKU", width=150)  # #设置列
        self.tree.column("FNSKU", width=150)
        self.tree.column("品名", width=400)
        self.tree.column("国家", width=150)
        self.tree.column("状态", width=150)
        self.tree.heading("SKU", text="SKU")  # #设置显示的表头名
        self.tree.heading("FNSKU", text="FNSKU")
        self.tree.heading("品名", text="品名")
        self.tree.heading("国家", text="国家")
        self.tree.heading("状态", text="状态")
        # self.xscroll.config(command=self.tree.xview)
        # self.xscroll.grid(side=BOTTOM, fill=X)
        vbar = ttk.Scrollbar(self, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        vbar.grid(row=2, column=12, sticky=NS)
        self.tree.grid(row=2, column=0, columnspan=10)
        self.tree.bind("<Double-1>", self.onDoubleClick)
        Button(self, text='获取未处理库存', font=self.ft, command=self.get_sku, width=15).grid(row=1, column=1, stick=W,padx=10, pady=10)
        Label(self, text='国家：', font=self.ft).grid(row=1, column=3, stick=E, padx=0, pady=10)
        self.cmb = ttk.Combobox(self, width=10, textvariable=self.country, state='readonly', takefocus=False, font=self.ft)
        self.cmb.grid(row=1, column=4, pady=5, stick=W)
        self.cmb['values'] = ['美国', '加拿大', '日本', '墨西哥', '荷兰', '瑞典', '波兰', '印度', '澳洲', '阿联酋', '新加坡',
                              '沙特阿拉伯', '土耳其','英国', '德国', '巴西', '意大利', '法国', '西班牙']
        self.lable = Label(self, text='相同SKU的FNSKU：', font=self.ft, bg='lightgray')
        self.lable.grid(row=1, column=5, stick=E, padx=0, pady=15)
        self.cmb2 = ttk.Combobox(self, width=15, textvariable=self.fnsku, takefocus=False, font=self.ft)
        self.cmb2.grid(row=1, column=6, pady=5, stick=W)
        self.cmb2['values'] = []
        Button(self, text='提交', font=self.ft, command=self.write_sql, width=15).grid(row=1, column=7, stick=W,padx=10, pady=10)
        Checkbutton(self, text="筛选换标FNSKU", onvalue=1, variable=self.Radiolist, font=self.ft).grid(row=1, column=2, padx=10, pady=10)
        Button(self, text='导出表格', font=self.ft, command=self.get_excl, width=10).grid(row=1, column=8, stick=W,padx=10, pady=10)

    def get_sku(self):
        items = self.tree.get_children()
        for i in items:
            self.tree.delete(i)
        country = self.country.get()
        self.sql()
        if country:
            sql = "select * from `process`.`FBA_warning` where `状态` = '未处理' and `国家` = '%s'" % country
        else:
            sql = "select * from `process`.`FBA_warning` where `状态` = '未处理'"
        self.cursor.execute(sql)
        result = self.cursor.fetchall()
        if result:
            list_fnsku = []
            information = []
            print(len(result))
            for i in result:
                if [i['FNSKU'], i['国家'], i['SKU']] not in list_fnsku:
                    information.append(i)
                    list_fnsku.append([i['FNSKU'], i['国家'], i['SKU']])
            print(len(information))
            self.show_information(information)
        else:
            messagebox.showinfo(message='当前没有未处理的库存')

    def write_sql(self):
        try:
            list_sku = self.get_tree()
            if list_sku:
                self.sql()
                print(list_sku)
                for i in list_sku:
                    if i[4] == '已处理':
                        sql = "update `process`.`FBA_warning` set `状态` = '已处理' where `SKU` = '%s' and `FNSKU` = '%s'" % (i[0], i[1])
                        self.cursor.execute(sql)
                self.connection.commit()
                messagebox.showinfo(message='提交成功')
            else:
                messagebox.showinfo(message='当前没有要提交的库存')
            self.get_sku()
        except Exception as e:
            messagebox.showinfo(message=e)

    def show_information(self, infromation):
        index = self.Radiolist.get()
        if index:
            for i in range(0, len(infromation)):
                sql = "select * from `amazon_form`.`pre_msku` where `SKU` = '%s' and `国家` = '%s' and `FNSKU` = " \
                      "'%s'" % (infromation[i]['SKU'], infromation[i]['国家'], infromation[i]['FNSKU'])
                self.cursor.execute(sql)
                res = self.cursor.fetchall()
                if res:
                    data = [infromation[i]['SKU'], infromation[i]['FNSKU'], infromation[i]['品名'], infromation[i]['国家'],'未处理']
                    self.tree.insert('', i, text=f'line{i + 1}', values=data)
        else:
            for i in range(0, len(infromation)):
                data = [infromation[i]['SKU'], infromation[i]['FNSKU'], infromation[i]['品名'], infromation[i]['国家'], '未处理']
                self.tree.insert('', i, text=f'line{i+1}', values=data)

    def get_tree(self):
        items = self.tree.get_children()
        # print(msg)
        list_sku = []
        for i in items:
            dict_item = self.tree.item(i)
            change_num = dict_item['values'][4]
            # print(dict_item['values'])
            if change_num == '已处理':
                list_sku.append(dict_item['values'])
        return list_sku

    def onDoubleClick(self, event):
        rowid = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        region = self.tree.identify_element(event.x, event.y)
        num = column[1:3]
        parent = self.tree.parent(rowid)
        if parent == '':
            pass
        x, y, width, height = self.tree.bbox(rowid, column)
        # self.xscroll.config(command=self.tree.xview)
        # self.xscroll.pack(side=BOTTOM, fill=X)
        # self.yscroll.config(command=self.tree.yview)
        # self.yscroll.pack(side=RIGHT, fill=Y)
        pady = height // 2
        text = self.tree.item(rowid, 'values')
        # print(text, rowid)
        a = int(num) - 1
        # print(a)
        if a >= 0:
            if a == 4:
                if text[4] == '未处理':
                    self.tree.set(rowid, column=4, value='已处理')
                elif text[4] == '已处理':
                    self.tree.set(rowid, column=4, value='未处理')
            else:
                self.entryPopup = EntryPopup(self.tree, rowid, text[a], a, width=12)
                self.entryPopup.place(x=x, y=y + pady, anchor=W)
                if a == 0:
                    self.get_fnsku(text[0], text[3])

    def get_fnsku(self, sku, country):
        self.cmb2['values'] = []
        get_url = f"https://erp.lingxing.com/api/storage/fbaLists?cid=&bid=&attribute=&asin_principal=&" \
                  f"search_field=sku&search_value={sku}&is_cost_page=0&status=&offset=0&length=20&" \
                  f"fulfillment_channel_type=&is_hide_zero_stock=0&is_parant_asin_merge=0&is_inclide_header=0&" \
                  f"is_contain_del_ls=0&req_time_sequence=/api/storage/fbaLists$$"
        get_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:95.0) Gecko/20100101 Firefox/95.0',
                       'Referer': 'https://erp.lingxing.com/erp/msupply/fbaInventory',}
        get_msg = global_var.s.get(get_url, headers=get_headers)
        get_msg = json.loads(get_msg.text)
        list_fnsku = []
        if get_msg['code'] == 1 and get_msg['msg'] == '操作成功':
            for i in get_msg['list']:
                if i['afn_fulfillable_quantity'] > 0:
                    self.sql()
                    sql = "select * from `process`.`FBA_inventory` where `FNSKU` = '%s' and `SKU` = '%s'" % (i['fnsku'], sku)
                    self.cursor.execute(sql)
                    result = self.cursor.fetchall()
                    if result and country == result[0]['国家']:
                        list_fnsku.append(i['fnsku'])
        print(list_fnsku)
        if list_fnsku:
            self.lable.configure(bg='green')
        else:
            self.lable.configure(bg='lightgray')
        self.cmb2['values'] = list_fnsku

    def get_excl(self):
        ask = messagebox.askokcancel(message='是否获取当前页面信息')
        if ask:
            try:
                items = self.tree.get_children()
                # print(msg)
                list_sku = []
                for i in items:
                    dict_item = self.tree.item(i)
                    list_sku.append(dict_item['values'])
                if list_sku:
                    self.makdir()
                    wb = openpyxl.Workbook()
                    wb_sheet = wb.active
                    wb_sheet.append(['SKU', 'FNSKU', '品名', '国家', '状态'])
                    for i in list_sku:
                        wb_sheet.append(i)
                    time_now = datetime.datetime.now().strftime("%Y%m%d%H%M")
                    wb.save(f'D:/FBA可售预警/FBA-{time_now}.xlsx')
                    messagebox.showinfo(message='生成成功')
                    os.startfile("D:/FBA可售预警")
                else:
                    messagebox.showinfo(message='当前页面无信息')
            except Exception as e:
                messagebox.showinfo(message=e)

    def makdir(self):
        folder = os.path.exists("D:/FBA可售预警")
        if not folder:
            os.makedirs("D:/FBA可售预警")


class FBA_warning (Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master  # 定义内部变量root
        self.xscroll = Scrollbar(self, orient=HORIZONTAL)
        self.yscroll = Scrollbar(self, orient=VERTICAL)
        self.ft = tkFont.Font(family='microsoft yahei', size=10)
        self.create()

    def create(self):
         Button(self, text='更新FBA库存情况', font=self.ft, command=self.get_warning, width=15).grid(row=1, column=4, stick=W, padx=10,pady=10)

    def get_warning(self):
        order = FBA_inventory.Find_order()
        if order.check():
            try:
                order.check_1()
                order.get_FBA()
                # print("succeed1")
                order.read_excl()
                # print("succeed2")
                order.get_fnsku()
                # print("succeed3")
                # order.check_0()
                messagebox.showinfo(message='生成成功')
                os.startfile("D:/FBA库存提醒")
            except Exception as e:
                # return False
                messagebox.showinfo(message=e)
            finally:
                order.check_0()
        else:
            messagebox.showinfo(message='当前正在生成表格，请稍后再试')
            return


class EntryPopup(Entry):
    def __init__(self, parent, iid, text, a, **kw):
        super().__init__(parent, **kw)
        self.tv = parent
        self.iid = iid
        self.count = a

        self.insert(0, text)
        self['readonlybackground'] = 'white'
        self['selectbackground'] = '#1BA1E2'
        self['exportselection'] = False

        self.focus_force()
        self.bind("<Return>", self.on_return)
        self.bind("<Control-a>", self.select_all)
        self.bind("<Escape>", lambda *ignore: self.destroy())
        self.bind("<FocusOut>", lambda *ignore: self.destroy())

    def on_return(self, event):
        new_value = self.get()
        old_values = self.tv.item(self.iid, 'values')
        new_values = list(old_values)
        new_values[self.count] = new_value
        values = tuple(new_values)
        self.tv.item(self.iid, values=values)
        self.destroy()

    def select_all(self, *ignore):
        self.selection_range(0, 'end')
        return 'break'