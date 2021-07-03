import xlwt, threading
import xlrd, time
from xlutils.copy import copy
import pandas as pd
import os

def instr(a1,a2,):
    return a2 in a1
#列索引序号
def get_n(data,name_1):
    a=0
    for i, v in data.iteritems():
        a=a+1
        if i==name_1:
            break
    return a

def column_classification(start, end):
    # 读excel
    for i in range(start, end):
        print(i)
        mutex.acquire()
        data_list = []
        data_list.extend(table.row_values(i))
        item = data_list[index]
        print('name:'+item)
        temp = False
        #  字符串比较
        for j in range(len(type_list)):
            if type_list[j] == item:
                temp = True
        if temp:
            write_excel(item, data_list)
        else:
            # 上锁
            type_list.append(item)
            new_excel(item, title)
            write_excel(item, data_list)
            # 释放锁
        mutex.release()


# 新建excel
def new_excel(name, new_list):
    try:
        # xlwt 支持256行  todo 需要改
        new_data = xlwt.Workbook()
        new_table = new_data.add_sheet(name)
        for n in range(len(new_list)):
            new_table.write(0, n, new_list[n])
        new_data.save(name + '.xls')
    except Exception as e:
        print(e)


# 往已存在的excel写数据
def write_excel(name, write_list):
    try:
        old_wb = xlrd.open_workbook(name + '.xls')
        new_wb = copy(old_wb)
        new_ws = new_wb.get_sheet(0)
        high = old_wb.sheets()[0].nrows
        for w in range(len(write_list)):
            new_ws.write(high, w, write_list[w])
        new_wb.save(name + '.xls')
    except Exception as e:
        print(e)


# 创建一个互斥锁，默认是没有上锁的
mutex = threading.Lock()


if __name__ == '__main__':
    print("="*40)
    print(" ")
    print("将表格与exe放入同一目录")
    print(" "*30+"by 徐阳")
    print(" ")
    print("="*40)




    print("输入需处理表格")
    path1=input()
    print("输入匹配关键词表格")
    path2=input()
    data=pd.read_excel(os.getcwd()+"/"+path1+'.xlsx')
    key=pd.read_excel(os.getcwd()+"/"+path2+'.xlsx')
    print("输入数据结果替换行")
    name_1=input()
    print("输入需检索数据行")
    name_2=input()
    print("需要线程数")
    thread_num = int(input())


    #填未处理
    for index, row in data.iterrows():
        data.loc[index,name_1]="未处理"



    #遍历匹配
    for key_w in key.itertuples():
        word_2=key_w[get_n(key,"关键词")]
        word_3=key_w[get_n(key,"结果")]
        print(word_2+"="*8+word_3)
        for index, row in data.iterrows():
            word_1=str(data.loc[index,name_2])
            if instr(word_1,word_2)==True:
                data.loc[index,name_1]=word_3
            else:
                pass

    

    
    data.to_excel('=====结果=====.xlsx')

    print("="*40)
    print("拆分表格 y/n")
    print("完成后回车退出")
    print("="*40)
    c=input()
    if c=="y":
        pass
    else:
        exit(0)


    file_name = os.getcwd()+"/"+'=====结果=====.xlsx'
    index = get_n(data,name_1)

    data = xlrd.open_workbook(file_name)
    table = data.sheets()[0]
    # 记录当前的行数据
    # 记录已经生成的excel
    type_list = []
    # 记录第一行
    title = []
    title.extend(table.row_values(0))

    rows = table.nrows


    # 多线程读写excel

    # 余数
    remainder = rows % thread_num
    pos = int((rows-remainder)/thread_num)


    t = []

    start_time = time.time()

    # 加入线程组
    for x in range(thread_num-1):
        t.append(threading.Thread(target=column_classification, args=(x*pos+1, (x+1)*pos)))
    t.append(threading.Thread(target=column_classification, args=((thread_num-1)*pos, (thread_num*pos)+remainder)))

    # 启动线程
    for tt in t:
        # 守护线程
        # tt.setDaemon(True)
        tt.start()

    end_time = time.time()
    print('花费时间是:', round(end_time-start_time, 4))



    input()