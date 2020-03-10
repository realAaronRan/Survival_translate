'''
@Description: make a GUI of translate.py
@Author: Aaron Ran
@Github: https://github.com/realAaronRan
@FilePath: \Survival_translate\translate_GUI.py
@Date: 2020-03-05 22:29:02
@LastEditTime: 2020-03-09 00:27:33
'''

import pandas as pd
import xlwings as xw
import multiprocessing as mlp
import tkinter as tk
from tkinter import ttk,filedialog
from pandastable import Table
from copy import deepcopy
from translate import translate

class BaseLangError(Exception):
    pass

class Root(tk.Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("My Love")
        self.height = 500
        self.width = 350
        # self.geometry('250x500')
        self.minsize(self.height, self.width)
        self.resizable(0, 0)

        # 步骤一：文档导入区
        self.file_zone = ttk.LabelFrame(self, borderwidth=3, relief='sunken', text='步骤一：导入文档')
        self.file_zone.grid(column=0, row=0)

        # 导入原始文档，并预览
        self.label_raw = tk.Label(self.file_zone, text="原始文档：")
        self.label_raw.grid(column=0, row=1)

        self.raw_file = tk.StringVar()
        self.textbox_raw = ttk.Entry(self.file_zone, width=50, textvariable=self.raw_file)
        self.textbox_raw.grid(column=1, row=1)

        self.button_raw = ttk.Button(self.file_zone, text = "浏览文件", command=self.exploreRawFile)
        self.button_raw.grid(column=2, row=1)

        self.button_preview_raw = ttk.Button(self.file_zone, text="预览", command=self.previewRawFile)
        self.button_preview_raw.grid(column=3, row=1)
        

        # 导入翻译文档
        self.label_trans = ttk.Label(self.file_zone, text="翻译文档：")
        self.label_trans.grid(column=0, row=2)

        self.trans_file = tk.StringVar()
        self.textbox_trans = ttk.Entry(self.file_zone, width=50, textvariable=self.trans_file)
        self.textbox_trans.grid(column=1, row=2)

        self.button_trans = ttk.Button(self.file_zone, text="浏览文件", command=self.exploreTransFile)
        self.button_trans.grid(column=2, row=2)

        self.button_preview_trans = ttk.Button(self.file_zone, text="预览", command=self.previewTransFile)
        self.button_preview_trans.grid(column=3, row=2)

        # 步骤二：语言选择区
        self.lang_select = ttk.LabelFrame(self, borderwidth=3, relief='sunken', text="步骤二：选择语言")
        self.lang_select.grid(column=0, row=4, sticky='W')

        # 确定基本语言
        self.label_base_lang = ttk.Label(self.lang_select, text='基本语言：')
        self.label_base_lang.grid(column=0, row=5)

        self.base_lang = tk.StringVar()
        self.combo_base_lang = ttk.Combobox(self.lang_select, width=48, textvariable=self.base_lang)
        self.combo_base_lang.grid(column=1, row=5, sticky='W')

        self.button_acquire_lang = ttk.Button(self.lang_select, text='获取语言', command=self.findAllLanguages)
        self.button_acquire_lang.grid(column=2, row=5, sticky='W')

        # 分析可翻译的语言
        self.label_valid_lang = ttk.Label(self.lang_select, text="翻译语言：")
        self.label_valid_lang.grid(column=0, row=6)

        self.check_frame = ttk.Frame(self.lang_select, borderwidth=3)
        self.check_frame.grid(column=1, row=6, rowspan=2, sticky='W')

        self.button_matched_lang = ttk.Button(self.lang_select, text='可行语言', command=self.findMatchedLanguages)
        self.button_matched_lang.grid(column=2, row=6)

        # 步骤三：游戏角色说明区
        self.role_frame = ttk.LabelFrame(self, borderwidth=3, relief='sunken', text='步骤三：角色说明')
        self.role_frame.grid(column=0, row=8, sticky='W')

        self.male_label = ttk.Label(self.role_frame, text='男性角色：')
        self.male_label.grid(column=0, row=0, sticky='W')

        self.male_text = tk.scrolledtext.ScrolledText(self.role_frame, width=48, height=1, wrap=tk.WORD)
        self.male_text.insert(tk.INSERT, '(角色名间请用一个空格分隔，boss名可以写入任意一个框格)')
        self.male_text.grid(column=1, row=0)

        self.female_label = ttk.Label(self.role_frame, text='女性角色：')
        self.female_label.grid(column=0, row=1, sticky='W')

        self.female_text = tk.scrolledtext.ScrolledText(self.role_frame, width=48, height=1, wrap=tk.WORD)
        self.female_text.grid(column=1, row=1)

        self.female_button = ttk.Button(self.role_frame, text='确定', command=self.getRoles)
        self.female_button.grid(column=2, row=0, rowspan=2)

        self.button_start_translate = ttk.Button(self, text='开始翻译', command=self.startTranslate)
        self.button_start_translate.grid(column=0, row=9, pady=15)


    def exploreRawFile(self):
        self.filename = filedialog.askopenfilename(initialdir="E:\\个人事宜\\Survival_translate", title="选择文件", filetype=(("All Files", "*.*"), ("xls", "*.xls"), ("xlsx", "*.xlsx")))
        self.textbox_raw.delete(0, tk.END)
        self.textbox_raw.insert(0, self.filename)
        self.raw_file.set(self.filename)

    def previewRawFile(self):
        try:
            excel_df = pd.read_excel(self.raw_file.get(), sheet_name='Sheet1').head(20)

            new_window = tk.Toplevel(self)
            new_window.title('原始文档预览')
            frame = tk.Frame(new_window)
            frame.pack(expand=True, fill='both')
            self.table_raw = Table(frame, showstatusbar=True)
            self.table_raw.model.df = excel_df
            self.table_raw.show()

        except FileNotFoundError:
            tk.messagebox.showerror(title='Warning', message="无法打开指定文件，请检查是否指定了文件路径，及文件路径是否正确。")

    def previewTransFile(self):
        try:
            excel_df = pd.read_excel(self.trans_file.get(), sheet_name='Sheet1').head(20)

            new_window = tk.Toplevel(self)
            new_window.title('翻译文档预览')

            frame = tk.Frame(new_window)
            frame.pack(expand=True, fill='both')
            
            self.table_trans = Table(frame, showstatusbar=True)
            self.table_trans.model.df = excel_df
            self.table_trans.show()

        except FileNotFoundError:
            tk.messagebox.showerror(title='Warning', message="无法打开指定文件，请检查是否指定了文件路径，及文件路径是否正确。")
        
    def exploreTransFile(self):
        self.filename = filedialog.askopenfilename(initialdir="E:\\个人事宜\\Survival_translate", title="选择文件", filetype=(("All Files", "*.*"), ("xls", "*.xls"), ("xlsx", "*.xlsx")))
        self.textbox_trans.delete(0, tk.END)
        self.textbox_trans.insert(0, self.filename)
        self.trans_file.set(self.filename)

    def fileSubmit(self):
        pass
    
    def findAllLanguages(self):
        # 读入excel文件
        try:
            excel_raw = pd.read_excel(self.raw_file.get(), sheet_name='Sheet1', header=None).head()
            lang_list = excel_raw.iloc[:2, 2:].fillna('0')
            self.lang_list = []
            for i in range(lang_list.shape[1]):
                c_lang = lang_list.iloc[0, i]
                e_lang = lang_list.iloc[1, i]
                if c_lang != '0' and e_lang != '0':
                    self.lang_list.append(c_lang)
            
            self.combo_base_lang['values'] = self.lang_list
        except FileNotFoundError:
            tk.messagebox.showerror(title='Error', message='读入原始文档失败，请检查是否已指定原始文档路径，及路径是否正确。')

    def findMatchedLanguages(self):
        try:
            excel_trans = pd.read_excel(self.trans_file.get(), sheet_name='Sheet1', header=None).head(1)
            all_lang = deepcopy(self.lang_list)
            all_lang.remove(self.base_lang.get())
            trans_lang = [excel_trans.iloc[0, i] for i in range(1, excel_trans.shape[1])]
            self.valid_lang = [lang for lang in trans_lang if lang in all_lang]
            try:
                for widget in self.check_frame.winfo_children():
                    widget.destroy()
            except:
                pass
            
            self.lang_select_states = []
            for i, lang in enumerate(self.valid_lang):
                # TODO: 记录选择的变量
                chk_value = tk.BooleanVar()
                check_button = ttk.Checkbutton(self.check_frame, text=lang, var=chk_value)
                check_button.grid(column=i%4, row=i//4, sticky='W')
                self.lang_select_states.append(chk_value)
        except FileNotFoundError:
            tk.messagebox.showerror(title='Error', message='读入翻译文档失败，请检查是否已指定翻译文档路径，及路径是否正确。')

    def getRoles(self):
        male_roles = self.male_text.get('0.0', tk.END).replace('\n', ' ').split(' ')
        female_roles = self.female_text.get('0.0', tk.END).replace('\n', ' ').split(' ')
        self.roles = {}
        self.roles['male'] = male_roles[:-1]
        self.roles['female'] = female_roles[:-1]

    def startTranslate(self):
        try:
            excel_raw = pd.read_excel(self.raw_file.get(), sheet_name='Sheet1')
            excel_trans = pd.read_excel(self.trans_file.get(), sheet_name='Sheet1')
            # 检查文档格式是否正确：翻译文档的第一列为基本语言
            first_lang = excel_trans.columns.values[0]
            if first_lang != self.base_lang.get():
                raise BaseLangError

            self.column_index_raw = {}
            for index, lang in enumerate(excel_raw.columns.values.tolist()):
                self.column_index_raw[lang] = index+1
            lang_select_list = [self.valid_lang[i] for i in range(len(self.valid_lang)) if self.lang_select_states[i].get()]

            self.return_dict = mlp.Manager().dict()
            cpu_num = int(0.8 * mlp.cpu_count())
            pool = mlp.Pool(processes=cpu_num)
            for lang in lang_select_list:
                sht_raw = excel_raw.loc[:, ['类型', self.base_lang.get()]]
                sht_trans = excel_trans.loc[:, [self.base_lang.get(), lang]]
                pool.apply_async(translate, (sht_raw, sht_trans, self.roles, self.return_dict, ))
            pool.close()
            pool.join()
            self.destroy()

        except BaseLangError:
            tk.messagebox.showerror('Error', '翻译文档的第一列不是基本语言。')
if __name__ == "__main__":
    root = Root()
    root.mainloop()

    wb_raw = xw.Book(root.raw_file.get())
    sht_raw = wb_raw.sheets('Sheet1')
    for lang, result in root.return_dict.items():
        index = root.column_index_raw[lang]
        sht_raw.range(3, index).options(transpose=True).value = result
    # 翻译结束后弹窗告知
    tk.messagebox.showinfo('Info', "已经成功完成翻译，并写入原始文档中，请查看。未匹配单元格已标注为\'$ERROR$\'")