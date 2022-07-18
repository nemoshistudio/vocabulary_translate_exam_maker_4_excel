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


#选择是否需要新手教程
help_mode = input('如果需要帮助，输入 help 查看帮助，或输入任意字符跳过（目前仅出卷功能）：\n')

if help_mode == 'help':
    print('已进入帮助模式，每一步都将有详细的提示')


#初始化表格处理函数
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False


#选择模式
print('------------------------------------------------')
print()
if help_mode == 'help':
    print('这一步将选择你要使用的功能，如果选错需要重启程序\n'
          '注意：使用阅卷功能需要在电脑上规范答题')
mode = input('出卷请输入 1 ，阅卷请输入 2 ：\n')
print()
print('------------------------------------------------')


#出卷模式
if mode == '1':
    if help_mode == 'help':
        print('你已经选择了 出卷 模式\n'
              '接下来要选择词库文件\n'
              '词库文件需要符合一定格式要求才可使用\n')
        more_option = input('如需了解规格详情请输入 more 下载示例文件\n')
        if more_option == 'more':
            webbrowser.open('https://github.com/nemoshistudio/vocabulary_translate_exam_maker_4_excel')


    #选择词库
    print('请选择词库文件')
    print()
    time.sleep(1)


    #调出选择窗口
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename()
    print('文件路径：' + Filepath)
    print()


    #打开文件，创建新文件
    wb1 = app.books.open(Filepath)
    wbnew = app.books.add()
    wbnew2 = app.books.add()



    #获取默写范围、数量等选项
    print('----------------------------------------------------')
    if help_mode == 'help':
        print('这一步要输入工作表名称，一般位于Excel窗口左下方，默认为Sheet1\n')

    sheetname = input('请输入工作表名称（默认Sheet1）：\n')
    if sheetname == '':
        sheetname = 'Sheet1'
    print()

    if help_mode == 'help':
        print('这一步要输入默写范围的起始词行数，如示例文件中 able 位于 B5 ，则输入 5\n')
    start = input('请输入词汇起始行数：\n')
    print()

    if help_mode == 'help':
        print('这里选择了默写范围的最后一行\n')
    end = input('请输入词汇结束行数：\n'
                'Tips:如一直到文件最后，可以在 C 列（第三列）末尾单元格输入 end，并在此处输入 end\n')
    print()

    if help_mode == 'help':
        print('这一步选择了你想要默写的数量\n')
    num = input('请输入默写个数：\n')
    print()

    if help_mode == 'help':
        print('这一步选择了是否保留序号，避免以序号为参照猜出首字母或更多\n')
    wantnum = input('请选择是否保留原表格序号，y 代表保留 ，n 代表 不保留(原表格第一列无序号也输入 n)：\n')
    print('----------------------------------------------------')


    #打开工作表
    print('正在初始化默写词库与试卷')
    sht1 = wb1.sheets[sheetname]


    #如输入为end，确定默写结束位置
    if end == 'end':
        endnum = 1
        while True:

            if sht1.range('c' + str(endnum)).value == 'end':
                break

            endnum = endnum + 1

        end = str(endnum - 1)


    #计算序号、单词、解释的单元格区域
    num_range = 'a' + start + ':a' + end
    words_range = 'b' + start + ':b' + end
    meanings_range = 'c' + start + ':c' + end


    #将单元格数据输入到列表

    num_list = sht1.range(num_range).value
    words_list = sht1.range(words_range).value
    meanings_list = sht1.range(meanings_range).value


    #通过时间生成唯一标识符，配对试卷与答案
    ticket = time.time()


    #初始化 答案.xlsx
    shtnew = wbnew.sheets['Sheet1']
    if wantnum == 'y':
        shtnew.range('a1').value = '原表格序号'
    shtnew.range('b1').value = '词汇'
    shtnew.range('c1').value = '释义'
    shtnew.range('z1').value = ticket
    shtnew.range('z2').value = '答案'


    #初始化 试卷.xlsx
    shtnew2 = wbnew2.sheets['Sheet1']
    if wantnum == 'y':
        shtnew2.range('a1').value = '原表格序号'
    shtnew2.range('b1').value = '词汇'
    shtnew2.range('c1').value = '释义'
    shtnew2.range('d1').value = '批改'
    shtnew2.range('z1').value = ticket
    shtnew2.range('z2').value = '试卷'
    shtnew2.range('e1').value = '禁止在答题区域单元格外篡改数据，避免影响阅卷。'



    #抽取题目并写入到位于内存的表格
    num = int(num)
    sum = 0
    print('正在随机抽取词汇出题')
    while sum < num:
        sum = sum + 1
        word_num = random.randint(0, len(words_list)-1)
        if wantnum == 'y':
            shtnew.range('a' + str(sum + 1)).value = num_list[word_num]
        shtnew.range('b' + str(sum + 1)).value = words_list[word_num]
        shtnew.range('c' + str(sum + 1)).value = meanings_list[word_num]

        if wantnum == 'y':
            shtnew2.range('a' + str(sum + 1)).value = num_list[word_num]
        shtnew2.range('c' + str(sum + 1)).value = meanings_list[word_num]

        if wantnum == 'y':
            del num_list[word_num]
        del words_list[word_num]
        del meanings_list[word_num]

    shtnew2.range('d' + str(sum + 2 )).value = 'end'



    #保存文件
    print('出题完毕')
    print()
    if help_mode == 'help':
        print('这一步将选择试卷与答案将保存于哪一个文件夹内\n')
    print('请选择保存路径')
    print()
    time.sleep(1)


    #调出窗口获取保存路径
    root = tk.Tk()
    root.withdraw()
    Folderpath = filedialog.askdirectory()
    print('保存路径：' + Folderpath + '/' )
    print()


    #保存
    wbnew.save(Folderpath + '/' + '答案.xlsx')
    wbnew2.save(Folderpath + '/' + '试卷.xlsx')
    print('保存完毕')
    wb1.close()
    wbnew.close()
    wbnew2.close()
    app.quit()
    #退出



#阅卷模式
if mode == '2':

    #选择答案与试卷
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

        #确认试卷与答案是否对应
        if sht2.range('z1').value == sht3.range('z1').value and sht2.range('z2').value == '答案' and sht3.range('z2').value == '试卷':
            break

        print('----------------------------------------------------')
        print()
        print('答案 或 试卷不符')
        print()
        print('文件错误，请重启程序')
        time.sleep(10)
        sys.exit(1)


    print('----------------------------------------------------')
    print()
    print('开始阅卷')
    t0 = time.time()


    #初始化
    num = 1
    rightnum = 0
    wrongnum = 0


    #阅卷
    while True:
        num = num + 1

        if sht3.range('d' + str(num)).value == 'end':#判断是否阅卷完毕
            break


        #获取答案
        key = sht2.range('b' + str(num)).value
        answer = sht3.range('b' + str(num)).value

        if answer == key:#正确
            sht3.range('d' + str(num)).value = '正确'
            rightnum =rightnum + 1

        if not answer == key:#错误
            sht3.range('d' + str(num)).value = key
            wrongnum = wrongnum + 1


    #计算正确率
    accuracy = rightnum / (rightnum + wrongnum)


    #写入正常率
    num = num + 1
    sht3.range('d' + str(num)).value = '正确数：' + rightnum
    num = num + 1
    sht3.range('d' + str(num)).value = '错误数：' + wrongnum
    num = num + 1
    sht3.range('d' + str(num)).value = '正确率：' + accuracy



    #输出正确率
    print('----------------------------------------------------')
    print()
    print('阅卷完毕')
    print()
    print('正确数：' + str(rightnum))
    print()
    print('错误数：' + str(wrongnum))
    print()
    print('正确率：' + str())
    print()
    print('阅卷用时：' + str(time.time() - t0) + 's')
    print()
    print('----------------------------------------------------')
    print()


    #保存批阅后答卷
    print('请选择改后试卷保存路径')
    root = tk.Tk()
    root.withdraw()
    Folderpath = filedialog.askdirectory()
    print()
    print('保存路径：' + Folderpath + '/' + '已批阅试卷.xlsx')


    wb3.save(Folderpath + '/' + '已批阅试卷.xlsx')
    wb2.close()
    wb3.close()
    app.quit()

print('------------------------------------------------')
last_word = input('输入 e 退出；输入 v 访问源码：\n')

if last_word == 'v':
    webbrowser.open('https://github.com/nemoshistudio/vocabulary_translate_exam_maker_4_excel')