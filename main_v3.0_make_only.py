#英语词汇默写生成
#西位Nemo
import xlwings as xw
import time
import random
import tkinter as tk
from tkinter import filedialog
import webbrowser
import sys
import os
from tkinter import *

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False


class GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name

    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("英语词汇默写")   #窗口名称
        self.init_window_name.geometry('960x680+100+100')    #分辨率
        self.init_window_name["bg"] = "white"   #背景色

        #标签
        self.unlock_button = Button(self.init_window_name, bg='silver', relief='sunken', text='解除Excel锁定',
                                    command=self.clear_excel)
        self.unlock_button.grid(row=1, column=1)
        self.exit_button = Button(self.init_window_name, bg='silver', relief='sunken', text='退出',
                                  command=self.exit_all)
        self.exit_button.grid(row=1, column=4)


        self.database_label = Label(self.init_window_name,bg='white',height = 2, text="词库文件路径:")
        self.database_label.grid(row=3,column=0)
        self.database_text = Text(self.init_window_name,bg='gainsboro',relief='sunken',width=50,height=5)
        self.database_text.grid(row=3, column=1)
        self.database_button = Button(self.init_window_name,bg='silver',relief='sunken',text='选择路径',
                                      command=self.choose_datapath)
        self.database_button.grid(row=3,column=4)

        self.out_label = Label(self.init_window_name, bg='white', height=2, text="保存路径:")
        self.out_label.grid(row=10, column=0)
        self.out_text = Text(self.init_window_name, bg='gainsboro', relief='sunken', width=50, height=5)
        self.out_text.grid(row=10, column=1)
        self.out_button = Button(self.init_window_name, bg='silver', relief='sunken', text='选择路径',
                                      command=self.choose_savepath)
        self.out_button.grid(row=10, column=4)


        self.init_sheet_label = Label(self.init_window_name, bg='white', height=2, text="工作表名称:")
        self.init_sheet_label.grid(row=4, column=0)
        self.sheet_text = Text(self.init_window_name, bg='gainsboro', relief='sunken', width=50, height=1)
        self.sheet_text.grid(row=4, column=1)

        self.init_start_label = Label(self.init_window_name, bg='white', height=2, text="起始行数:")
        self.init_start_label.grid(row=5, column=0)
        self.init_start_text = Text(self.init_window_name, bg='gainsboro', relief='sunken', width=50, height=1)
        self.init_start_text.grid(row=5, column=1)

        self.init_end_label = Label(self.init_window_name, bg='white', height=2, text="截止行数:")
        self.init_end_label.grid(row=6, column=0)
        self.init_end_text = Text(self.init_window_name, bg='gainsboro', relief='sunken', width=50, height=1)
        self.init_end_text.grid(row=6, column=1)

        self.init_num_label = Label(self.init_window_name, bg='white', height=2, text="默写个数:")
        self.init_num_label.grid(row=7, column=0)
        self.init_num_text = Text(self.init_window_name, bg='gainsboro', relief='sunken', width=50, height=1)
        self.init_num_text.grid(row=7, column=1)
        self.make_out_button = Button(self.init_window_name, bg='silver', relief='sunken', text='出卷',
                                      command=self.make_out)
        self.make_out_button.grid(row=15, column=2)

    def exit_all(self):
        sys.exit(0)

    def clear_excel(self):
        os.system('taskkill /f /t /im EXCEL.exe')

    def choose_datapath(self):
        global Filepath
        src = self.database_text.get(1.0, END).strip().replace("\n", "").encode()
        root = tk.Tk()
        root.withdraw()
        Filepath = filedialog.askopenfilename()
        self.database_text.delete(1.0, END)
        self.database_text.insert(1.0,Filepath)

    def choose_savepath(self):
        global Filepath
        src = self.out_text.get(1.0, END).strip().replace("\n", "").encode()
        root = tk.Tk()
        root.withdraw()
        folderpath = filedialog.askdirectory()
        self.out_text.delete(1.0, END)
        self.out_text.insert(1.0, folderpath)

    def make_out(self):
        filepath = self.database_text.get(1.0, END).strip().replace("\n", "")
        sheetname = self.sheet_text.get(1.0, END).strip().replace("\n", "")
        start = self.init_start_text.get(1.0, END).strip().replace("\n", "")
        end = self.init_end_text.get(1.0, END).strip().replace("\n", "")
        num = self.init_num_text.get(1.0, END).strip().replace("\n", "")
        folderpath = self.out_text.get(1.0, END).strip().replace("\n", "")

        wb1 = app.books.open(Filepath)
        wbnew = app.books.add()
        wbnew2 = app.books.add()

        sht1 = wb1.sheets[sheetname]

        if end == '':
            endnum = 1
            while True:

                if sht1.range('c' + str(endnum)).value == None and sht1.range('b' + str(endnum)).value == None:
                    break

                endnum = endnum + 1

            end = str(endnum - 1)

        num_range = 'a' + start + ':a' + end
        words_range = 'b' + start + ':b' + end
        meanings_range = 'c' + start + ':c' + end

        num_list = sht1.range(num_range).value
        words_list = sht1.range(words_range).value
        meanings_list = sht1.range(meanings_range).value

        ticket = time.time()

        shtnew = wbnew.sheets['Sheet1']
        shtnew.range('b1').value = '词汇'
        shtnew.range('c1').value = '释义'
        shtnew.range('z1').value = ticket
        shtnew.range('z2').value = '答案'

        shtnew2 = wbnew2.sheets['Sheet1']
        shtnew2.range('b1').value = '词汇'
        shtnew2.range('c1').value = '释义'
        shtnew2.range('d1').value = '批改'
        shtnew2.range('z1').value = ticket
        shtnew2.range('z2').value = '试卷'
        shtnew2.range('e1').value = '禁止在答题区域单元格外篡改数据，避免影响阅卷。'

        num = int(num)
        sum = 0

        while sum < num:
            sum = sum + 1
            word_num = random.randint(0, len(words_list) - 1)

            shtnew.range('b' + str(sum + 1)).value = words_list[word_num]
            shtnew.range('c' + str(sum + 1)).value = meanings_list[word_num]

            shtnew2.range('c' + str(sum + 1)).value = meanings_list[word_num]

            del words_list[word_num]
            del meanings_list[word_num]

        shtnew2.range('d' + str(sum + 2)).value = 'end'
        wbnew.save(folderpath + '/' + '答案.xlsx')
        wbnew2.save(folderpath + '/' + '试卷.xlsx')
        wb1.close()
        wbnew.close()
        wbnew2.close()
        app.quit()

def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    Vocabulary_writing = GUI(init_window)
    # 设置根窗口默认属性
    Vocabulary_writing.set_init_window()

    init_window.mainloop()          #父窗口进入事件循环，可以理解为保持窗口运行，否则界面不展示


gui_start()
