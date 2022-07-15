#英语词汇默写生成
#西位Nemo
import xlwings as xw
import time
import random
import tkinter as tk
from tkinter import filedialog
import webbrowser

print('英语词汇默写生成器')
print()
print('by 西位Nemo')
print()
print('已完美支持 xls/xlsx 文件')
print()
print('------------------------------------------------')
print('请选择文件')
print()
time.sleep(1)

root = tk.Tk()
root.withdraw()
Filepath = filedialog.askopenfilename()
print('文件路径：' + Filepath)
print()

app = xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False
wb1 = app.books.open(Filepath)
wbnew = app.books.add()


start = input('请输入词汇起始行数：\n')
print()
end = input('请输入词汇结束行数：\n')
print()
num = input('请输入默写个数：\n')
print()

num_range = 'a' + start + ':a' + end
words_range = 'b' + start + ':b' + end
meanings_range = 'c' + start + ':c' + end

sht1 = wb1.sheets['Sheet1']
num_list = sht1.range(num_range).value
words_list = sht1.range(words_range).value
meanings_list = sht1.range(meanings_range).value

sht2 = wbnew.sheets['Sheet1']
sht2.range('a1').value = '原表格序号'
sht2.range('b1').value = '词汇'
sht2.range('c1').value = '释义'

num = int(num)
sum = 0
while sum < num:
    sum = sum + 1
    word_num = random.randint(0, len(num_list)-1)
    sht2.range('a' + str(sum + 1)).value = num_list[word_num]
    sht2.range('b' + str(sum + 1)).value = words_list[word_num]
    sht2.range('c' + str(sum + 1)).value = meanings_list[word_num]
    del num_list[word_num]
    del words_list[word_num]
    del meanings_list[word_num]


name = input('请输入想要的文件名：\n')
name = name + '.xlsx'
print()

print('请选择保存路径')
print()
time.sleep(1)

root = tk.Tk()
root.withdraw()
Folderpath = filedialog.askdirectory()
print('保存路径：' + Folderpath + '/' + name)
print()

wb1.save()
wbnew.save(Folderpath + '/' + name)
wb1.close()
wbnew.close()
app.quit()

print('------------------------------------------------')
last_word = input('输入 e 退出；输入 v 访问源码：\n')

if last_word == 'v':
    webbrowser.open('https://github.com/nemoshistudio/vocabulary_translate_exam_maker_4_excel')
    
