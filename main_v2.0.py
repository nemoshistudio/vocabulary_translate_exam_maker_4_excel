#英语词汇默写生成
#西位Nemo
import xlwings as xw
import time
import random
import tkinter as tk
from tkinter import filedialog
import webbrowser
import sys

print('英语词汇默写生成器')
print()
print('by 西位Nemo')
print()
print('已完美支持 xls/xlsx 文件')
print()
print('------------------------------------------------')
print()
mode = input('出卷请输入 1 ，阅卷请输入 2 ：\n')
print()
print('------------------------------------------------')

app = xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False

if mode == '1':
    print('请选择文件')
    print()
    time.sleep(1)

    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename()
    print('文件路径：' + Filepath)
    print()


    wb1 = app.books.open(Filepath)
    wbnew = app.books.add()
    wbnew2 = app.books.add()


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

    ticket = time.time()

    shtnew = wbnew.sheets['Sheet1']
    shtnew.range('a1').value = '原表格序号'
    shtnew.range('b1').value = '词汇'
    shtnew.range('c1').value = '释义'
    shtnew.range('z1').value = ticket
    shtnew.range('z2').value = '答案'

    shtnew2 = wbnew2.sheets['Sheet1']
    shtnew2.range('a1').value = '原表格序号'
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
        word_num = random.randint(0, len(num_list)-1)
        shtnew.range('a' + str(sum + 1)).value = num_list[word_num]
        shtnew.range('b' + str(sum + 1)).value = words_list[word_num]
        shtnew.range('c' + str(sum + 1)).value = meanings_list[word_num]

        shtnew2.range('a' + str(sum + 1)).value = num_list[word_num]
        shtnew2.range('c' + str(sum + 1)).value = meanings_list[word_num]


        del num_list[word_num]
        del words_list[word_num]
        del meanings_list[word_num]

    shtnew2.range('d' + str(sum + 2 )).value = 'end'

    print()

    print('请选择保存路径')
    print()
    time.sleep(1)

    root = tk.Tk()
    root.withdraw()
    Folderpath = filedialog.askdirectory()
    print('保存路径：' + Folderpath + '/' )
    print()


    wbnew.save(Folderpath + '/' + '答案.xlsx')
    wbnew2.save(Folderpath + '/' + '试卷.xlsx')
    wb1.close()
    wbnew.close()
    wbnew2.close()
    app.quit()



if mode == '2':

    while True :
        print('请选择答案文件')
        print()
        time.sleep(1)

        root = tk.Tk()
        root.withdraw()
        Filepath2 = filedialog.askopenfilename()
        print('答案文件路径：' + Filepath2)
        print()

        print('请选择试卷文件')
        print()
        time.sleep(1)

        root = tk.Tk()
        root.withdraw()
        Filepath3 = filedialog.askopenfilename()
        print('试卷文件路径：' + Filepath3)
        print()

        wb2 = app.books.open(Filepath2)
        wb3 = app.books.open(Filepath3)

        sht2 = wb2.sheets['Sheet1']
        sht3 = wb3.sheets['Sheet1']

        if sht2.range('z1').value == sht3.range('z1').value and sht2.range('z2').value == '答案' and sht3.range('z2').value == '试卷':
            break

        print('----------------------------------------------------')
        print('答案 或 试卷不符')
        print()
        print('文件错误，请重启程序')
        time.sleep(10)
        sys.exit(1)


    print('----------------------------------------------------')
    print('开始阅卷')
    t0 = time.time()

    num = 1
    rightnum = 0
    wrongnum = 0
    while True:
        num = num + 1

        if sht3.range('d' + str(num)).value == 'end':
            break

        key = sht2.range('b' + str(num)).value
        answer = sht3.range('b' + str(num)).value

        if answer == key:
            sht3.range('d' + str(num)).value = '正确'
            rightnum =rightnum + 1

        if not answer == key:
            sht3.range('d' + str(num)).value = key
            wrongnum = wrongnum + 1

        wb3.save()

    print('----------------------------------------------------')
    print('阅卷完毕')
    print('正确数：' + str(rightnum))
    print('错误数：' + str(wrongnum))
    print('正确率：' + str(rightnum / (rightnum + wrongnum)))
    print('阅卷用时：' + str(time.time() - t0) + 's')


    wb3.save('已批改试卷.xlsx')
    wb2.close()
    wb3.close()
    app.quit()

print('------------------------------------------------')
last_word = input('输入 e 退出；输入 v 访问源码：\n')

if last_word == 'v':
    webbrowser.open('https://github.com/nemoshistudio/vocabulary_translate_exam_maker_4_excel')
