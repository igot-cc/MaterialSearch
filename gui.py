from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import main, search_in_sh as sc
import os


class My_gui(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        # self.master = master
        # self.pack()

        # # 菜单栏
        # self.menubar = Menu(master)
        # self.filemenu = Menu(self.menubar, tearoff=0)
        # self.menubar.add_cascade(label='文件', menu=self.filemenu)
        # self.filemenu.add_command(label='打开')
        # self.filemenu.add_command(label='保存')
        # self.filemenu.add_command(label='退出')
        # master.config(menu=self.menubar)

        # PanedWindow上窗口
        self.panedwin = PanedWindow(self.master, orient='vertical', sashrelief='sunken')
        self.panedwin.pack(fill='both', expand=1)

        # 状态栏
        self.separator = ttk.Separator(self.master).pack(padx=2, fill='x')
        self.status_frame = ttk.Frame(self.separator, relief='raised', height='20').pack(fill='x')
        # self.status_label = ttk.Label(self.status_frame, text='状态栏').pack(side=LEFT, fill='x')
        # # 进度条
        self.pb = ttk.Progressbar(self.status_frame, length=150, value=0, mode='indeterminate')
        # label
        self.label_status = ttk.Label(self.status_frame, text='')
        self.label_status.pack(side=LEFT)

        # 锚点
        # self.sizegrip = ttk.Sizegrip(self.status_frame).pack(anchor=NE)


        # 窗口
        self.panedwin_top = PanedWindow(self.panedwin, orient=VERTICAL, sashrelief=SUNKEN)
        self.panedwin_bottom = PanedWindow(self.panedwin, orient=HORIZONTAL, sashrelief=SUNKEN)

        self.top_frame = ttk.Frame(self.panedwin_top, height='500', relief=FLAT)
        self.left_frame = ttk.Frame(self.panedwin_bottom, width='200', relief=FLAT)
        self.right_frame = ttk.Frame(self.panedwin_bottom, width='200', relief=FLAT)

        self.panedwin.add(self.panedwin_top)
        self.panedwin.add(self.panedwin_bottom)
        self.panedwin_top.add(self.top_frame)
        self.panedwin_bottom.add(self.left_frame)
        self.panedwin_bottom.add(self.right_frame)

        # label、Entry、Button
        self.label1 = ttk.Label(self.top_frame, text='A002物料表地址：')
        self.entry1 = ttk.Entry(self.top_frame, width=30)
        self.entry1.insert(0, 'D:\\')
        self.label2 = ttk.Label(self.top_frame, text='A002物料表名称：')
        self.entry2 = ttk.Entry(self.top_frame, width=20)
        self.entry2.insert(0, 'A002采购数据')

        self.label3 = ttk.Label(self.top_frame, text='C016物料表地址：')
        self.entry3 = ttk.Entry(self.top_frame, width=30)
        self.entry3.insert(0, 'D:\\')
        self.label4 = ttk.Label(self.top_frame, text='C016物料表名称：')
        self.entry4 = ttk.Entry(self.top_frame, width=20)
        self.entry4.insert(0, 'C016温州仓库物料')

        self.label5 = ttk.Label(self.top_frame, text='BOM表地址：')
        self.entry5 = ttk.Entry(self.top_frame, width=30)
        self.entry5.insert(0, 'D:\\')
        self.label6 = ttk.Label(self.top_frame, text='BOM表名称：')
        self.entry6 = ttk.Entry(self.top_frame, width=20)
        self.entry6.insert(0, 'WG_NILMV02')

        self.label1.grid(row=0, sticky=W, pady=2)
        self.entry1.grid(row=0, column=1, sticky=E+W, pady=2)#, columnspa=1
        self.label2.grid(row=0, column=2, sticky=W, pady=2)
        self.entry2.grid(row=0, column=3, sticky=E+W, pady=2)#, columnspa=1

        self.label3.grid(row=1, sticky=W, pady=2)
        self.entry3.grid(row=1, column=1, sticky=E + W, pady=2)  # , columnspa=1
        self.label4.grid(row=1, column=2, sticky=W, pady=2)
        self.entry4.grid(row=1, column=3, sticky=E + W, pady=2)  # , columnspa=1

        self.label5.grid(row=2, sticky=W, pady=2)
        self.entry5.grid(row=2, column=1, sticky=E+W, pady=2)#, columnspa=1
        self.label6.grid(row=2, column=2, sticky=W, pady=2)
        self.entry6.grid(row=2, column=3, sticky=E + W, pady=2)#, columnspa=1

        self.button1 = ttk.Button(self.top_frame, text='确定', command=self.ok)
        self.button2 = ttk.Button(self.top_frame, text='移除', command=self.cancel)
        self.button3 = ttk.Button(self.top_frame, text='查找A002', command=self.start_A002)
        self.button4 = ttk.Button(self.top_frame, text='查找C016', command=self.start_C016)
        self.button5 = ttk.Button(self.top_frame, text='完成', command=self.finish)
        self.button1.grid(row=4, column=2, sticky=W, pady=2, padx=2)
        self.button2.grid(row=4, column=3, sticky=W, pady=2, padx=28)
        # self.button3.grid(row=0, column=4, sticky=E, pady=3, padx=2, rowspan=2)
        self.button3.grid(row=0, column=4, sticky=E, pady=2, padx=8)
        # self.button4.grid(row=1, column=4, sticky=E, pady=3, padx=2, rowspan=2)
        self.button4.grid(row=1, column=4, sticky=E, pady=2, padx=8)
        self.button5.grid(row=4, column=4, sticky=E, pady=2, padx=8)

        # 左边listbox
        self.labelframe1 = ttk.LabelFrame(self.left_frame, text='仓库物料')
        self.scrollbar1 = ttk.Scrollbar(self.labelframe1)
        self.listbox_left = Listbox(self.labelframe1, width=50, height=14, yscrollcommand=self.scrollbar1.set, exportselection=False)
        self.scrollbar1.config(command=self.listbox_left.yview)
        self.labelframe1.pack()
        self.scrollbar1.pack(side=RIGHT, fill=Y)
        self.listbox_left.pack(side=LEFT)
        # self.listbox_left.bind("<Button-1>", lambda e: self.get_value(e.y))

        # 右边listbox
        self.labelframe2 = ttk.LabelFrame(self.right_frame, text='BOM表物料')
        self.scrollbar2 = ttk.Scrollbar(self.labelframe2)
        self.listbox_right = Listbox(self.labelframe2, width=50, height=14, yscrollcommand=self.scrollbar2.set)
        self.scrollbar2.config(command=self.listbox_right.yview)
        self.labelframe2.pack()
        self.scrollbar2.pack(side=RIGHT, fill=Y)
        self.listbox_right.pack(side=RIGHT)
        self.listbox_right.bind("<Button-1>", lambda e: self.insert_wuliao(e.y))

        # 定义一个字典将右边listbox中的值对应到bom表中的索引
        self.value_to_index = {}
        # 定义一个列表存储匹配到的每一个物料
        self.list = []
        # 定义一个flag代表确定按钮选取的事那个仓库的物料，1：A002仓库，2：C016仓库
        self.flag = 0

    # 完成按钮：选择查找到的正确的物料
    def ok(self):
        dict_temp = {}
        value = self.get_left_value()   # 拿到左边窗口点击的值
        excel_index = self.index + 2
        dict_temp[excel_index] = list(value)
        if self.flag == 1:
            main.dict_A002.update(dict_temp)
        if self.flag == 2:
            main.dict_C016.update(dict_temp)
        main.has_searched_index.append(self.index)
        # print(excel_index)
        # print(value)
        # print(main.dict)
        self.clear()

    # 删除右边选中的值，同时清除左边
    def clear(self):
        self.listbox_right.delete(self.listbox_right.curselection())
        self.listbox_left.delete(0, END)

    def clear_all(self):
        self.listbox_right.delete(0, 'end')
        self.listbox_left.delete(0, 'end')

    def cancel(self):
        self.clear()

    def progressbar_show(self):
        self.pb.pack(self.status_frame, side=LEFT, padx=3, pady=2)
        self.pb.start()
    #
    # def progressbar_kill(self):
    #     self.pb.stop()
    #     self.pb.grid_forget()

    def start_A002(self):
        self.label_status.configure(text='正在查找...')
        self.master.update()
        self.flag = 1
        self.labelframe1.config(text='A002物料')
        try:
            # 获取地址
            A002_address = self.entry1.get()
            A002_name = self.entry2.get() + '.xlsx'
            A002 = os.path.join(A002_address, A002_name)
            bom_address = self.entry5.get()
            bom_name = self.entry6.get() + '.xlsx'
            main.BOM = os.path.join(bom_address, bom_name)
            main.BOM_sheet = self.entry6.get()
            export_name = self.entry6.get() + '-new.xlsx'
            main.export_address = os.path.join(bom_address, export_name)

            if not main.comment:
                self.labelframe1.text = 'A002物料'
                main.read_bom()
                main.read_warehouse(A002)
                # print(len(main.wuliao_script))
                main.search_null()
                dict_res = main.search_res()
                dict_cap = main.search_cap()
                main.dict_A002.update(dict_res)
                main.dict_A002.update(dict_cap)
                self.clear_all()
                main.get_other()
                self.insert_bom(main.unsearched_index)
                self.label_status.config(text='')

            else:
                # main.dict_C016.update(main.dict)
                # main.dict.clear()
                main.read_warehouse(A002)
                dict_res = main.search_res()
                dict_cap = main.search_cap()
                main.dict_A002.update(dict_res)
                main.dict_A002.update(dict_cap)
                self.clear_all()
                main.get_other()
                self.labelframe1.text = 'A002物料'
                self.insert_bom(main.unsearched_index)
                self.label_status.configure(text='')
        except FileNotFoundError:
            self.label_status.configure(text='')
            self.warning(message='文件名或路径错误！')

    def start_C016(self):
        self.label_status.configure(text='正在查找...')
        self.master.update()
        self.flag = 2
        self.labelframe1.config(text='C016物料')
        try:
            # 获取地址
            C016_address = self.entry3.get()
            C016_name = self.entry4.get() + '.xlsx'
            C016 = os.path.join(C016_address, C016_name)
            bom_address = self.entry5.get()
            bom_name = self.entry6.get() + '.xlsx'
            main.BOM = os.path.join(bom_address, bom_name)
            main.BOM_sheet = self.entry6.get()
            export_name = self.entry6.get() + '-new.xlsx'
            main.export_address = os.path.join(bom_address, export_name)

            if not main.comment:
                main.read_bom()
                main.read_warehouse(C016)
                main.search_null()
                dict_res = main.search_res()
                dict_cap = main.search_cap()
                main.dict_C016.update(dict_res)
                main.dict_C016.update(dict_cap)
                self.clear_all()
                main.get_other()
                self.insert_bom(main.unsearched_index)
                self.label_status.config(text='')
            else:
                # main.first_dict.update(main.dict)
                # print(main.first_dict)
                # main.dict.clear()
                main.read_warehouse(C016)
                # print(len(main.wuliao_script))
                dict_res = main.search_res()
                dict_cap = main.search_cap()
                main.dict_C016.update(dict_res)
                main.dict_C016.update(dict_cap)
                self.clear_all()
                main.get_other()
                self.labelframe1.text = 'C016物料'
                self.insert_bom(main.unsearched_index)
                self.label_status.configure(text='')
        except FileNotFoundError:
            self.label_status.configure(text='')
            self.warning(message='文件名或路径错误！')
        pass

    def finish(self):
        if (len(main.dict_A002)) | (len(main.dict_C016)):
            # for i in range(main.comment_len):
            #     if (i not in main.has_searched_index) & (i not in main.unsearched_index):
            #         print(main.comment[i])
            main.write_to_excel()
            self.master.destroy()
        else:
            self.warning('未完成匹配或无匹配项')

    # 插入bom表检索出的数据
    def insert_bom(self, unsearched_index):
        if not unsearched_index:
            self.listbox_left.insert(END, '无需要匹配项')
        for i in unsearched_index:
            comment_other = main.comment[i]
            footprint_other = main.footprint[i]
            str = comment_other + ' / ' + footprint_other
            self.value_to_index[str] = i
            self.listbox_right.insert(END, str)

    # 获取右边窗口点击的条目并搜索后插入左边listbox
    def insert_wuliao(self, y):
        # if self.listbox_right.curselection():     # curselection()方法返回的是上一次点击的序号
        curvalue = self.listbox_right.nearest(y)
        if curvalue != -1:
            value = self.listbox_right.get(curvalue)    # value值为listbox_right显示的值
            self.index = self.value_to_index[value]       # 将value转为对应的excel序号
            key = main.unsearched_key[self.index]     # 将序号转为关键字key，然后去物料表匹配
            print(key)
            self.list = sc.search_other_in_a002(key)    # list为匹配到的物料列表
            if not self.list:
                self.list.append('无匹配项')
            # print(curvalue)
            # print(value)
            # print(index)
            # print(key)
            # print(list)

            # 将查找到的物料插入到左边lisbox
            self.listbox_left.delete(0, END)
            for i in self.list:
                self.listbox_left.insert(END, i)

    # 获取左边窗口点击的值
    def get_left_value(self):
        index = self.listbox_left.curselection()
        value = self.listbox_left.get(index)
        return value

    def warning(self, message):
        messagebox.showwarning(message=message)

# root = Tk()
# root.geometry('750x450+200+100')
# # root.minsize(100,50)
# root.resizable(0,0)
# root.title('search')
# app = My_gui(master=root)
#
# root.mainloop()
