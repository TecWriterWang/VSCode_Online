from ast import Continue
from concurrent.futures import thread
import threading
import urllib.request
import urllib.error
import time
import os
import sys
import threading
import openpyxl

# https://www.cnblogs.com/luyuze95/p/11289143.html
# https://blog.csdn.net/Key_book/article/details/80258022
# http://baostock.com/baostock/index.php/%E9%A6%96%E9%A1%B5


stock_number = ["sz300059","sz300604","sh600760","sh601899","sh600559","sz000807",
"sz000422","sh601988","sh600096","sh601877","sz002714","sh601919","sh600048","sh601669",
"sz160119","sz160706","sh600460","sz002192","sh600502","sh600388","sh000001"]

# fomat time data
#time.strftime('%Y%m%d %H:%M:%S', time.localtime())

#currTime = time.strftime('%Y%m%d', time.localtime())
currTime = time.strftime('%Y%m', time.localtime())

str_html = []


def updata2xls():
    stock_number = ["sz300059","sz300604","sh600760","sh601899","sh600559"
    ,"sz000807","sz000422","sh601988","sz002714","sh601919"]

    stock_number_last_Price=('B9','E9','H9','K9','B20','E20','H20','K20','B31','E31')

    #=HYPERLINK("#"&CELL("address",Q586),Q586)

    Price_list = list()

    # new workbook
    #wb = openpyxl.Workbook()
    #sh = book['TaoLi']

    # Load exist workbook
    wb = openpyxl.load_workbook(r'C:\Users\wangned\Desktop\TaoLi_Road_Release.xlsm')
    sheet = wb['Summary']


    for code in stock_number:
        with urllib.request.urlopen('https://qt.gtimg.cn/q=' + str(code)) as response:
            html = response.read()
            html = html.decode('gbk')
        
        TodayPrice = str(html)      
        #print(TodayPrice)

        # Code_1 = "~" + str(code[2:-1])
        # Stock_Name = TodayPrice[TodayPrice.find("~")+1:TodayPrice.find(Code_1)]
        
        Code_Pattern = "~" + str(code[2:-1]) + str(code[-1]) +  "~"
        #print(Code_Pattern)

        Price = TodayPrice[TodayPrice.find(Code_Pattern) + 8:
                        TodayPrice.find("~",TodayPrice.find(Code_Pattern) + 8)]

        Price_list.append(Price)


    for i in range(0,len(stock_number_last_Price)):
        sheet[stock_number_last_Price[i]] = float(Price_list[i])

    wb.save(r'C:\Users\wangned\Desktop\TaoLi_Road_Release.xlsm')


def get_html(index=0):
    global str_html

    str_html.clear()
    index = 0

    try:
        for code in stock_number:
            html = ""
            
            # 超时重新读取,直到获取到新内容
            while 1:
                try:
                    with urllib.request.urlopen('https://qt.gtimg.cn/q=' + str(code), timeout=5) as response:
                        html = response.read()
                        # print(html.decode('utf8'))
                        str_html.append(str(html.decode('gbk')))
                    print("**" * index,end="\r")
                    index = index + 1
                    break
                except Exception as e:
                    os.system("cls")
                    continue

    # except urllib.error.URLError:
    # except urllib.error.HTTPError as e:
    #     # if e.code == 404:
    #     print("Error 404")
    #     return 0
    except Exception as e:
         print("\nError "+str(e))

def print_log():
    global str_html
    i = 0

    if len(str_html) < len(stock_number):
        print("="*20)
        return 0
    else:
        os.system("cls")
                
    for code in stock_number:
        TodayPrice = str_html[i]
        i = i+1

        Code_1 = "~" + str(code[2:-1])
        
        Stock_Name = TodayPrice[TodayPrice.find("~")+1:TodayPrice.find(Code_1)]
        
        # code_Number
        Code_Number = "~" + str(code[2:-1]) + str(code[-1]) +  "~"
        
        Price_info = TodayPrice[TodayPrice.find(Code_Number) + 8:TodayPrice.find(Code_Number) + 48]
        Price_list = list(Price_info.split("~",5))
        
        New_Price = float(Price_list[0])
        Yesterday_End_Price = float(Price_list[1])
        Today_Begin_Price = float(Price_list[2])
        
        #不考虑ST、科创
        if str(code[2:5]) != "300":     
            limit_Raising_Price = Today_Begin_Price + Today_Begin_Price*0.1
            Limit_Down_Price = Today_Begin_Price - Today_Begin_Price*0.1
        else:
            limit_Raising_Price = Today_Begin_Price + Today_Begin_Price*0.2
            Limit_Down_Price = Today_Begin_Price - Today_Begin_Price*0.2
        
        #New_Price = TodayPrice[TodayPrice.find(Code_Number) + 8:TodayPrice.find("~",TodayPrice.find(Code_Number) + 8)]
        
        Code_Number = "~~" + currTime
        Startindex = TodayPrice.find(Code_Number)
        Startindex = TodayPrice.find("~",Startindex+2)
        
        Price_info = TodayPrice[Startindex+1:Startindex+50]
        
        Price_list = list(Price_info.split("~",4))
        
        #涨跌幅
        diff_Value = float(Price_list[0])
        #涨跌幅百分比
        diff_Value_100 = float(Price_list[1])
        #最高价
        Highest_Price = float(Price_list[2])
        #最低价
        Lowest_Price = float(Price_list[3])
        #成交量        
        temp_vale = list(str(Price_list[4]).split("/",2))

        Trans_quantity = int(temp_vale[1])/10000
        #成交金额
        Trans_crash = int(temp_vale[2].split("~",1)[0])/100000000

        #print("Code:{0}\tName:{1}\t Price:{2}".format(code,Stock_Name,New_Price))
        if 1: 
            #print("{0} {1}\t".format(Stock_Name,New_Price))
            print(Stock_Name,New_Price,end="\t",flush=True)
        else:
            #print("Code  :{0}\nName  :{1}\n最新价:{2}\n最高价:{3}\n最低价:{4}\n涨幅% :{5}%\n涨幅  :{6}\n涨停价:{7}\n跌停价:{8}\n成交量:{9}万\n成交额:{10}亿\n".format(code, Stock_Name, New_Price, Highest_Price,
            print("Code  :{0}\tName  :{1}\n最新价:{2}\t最高价:{3}\t最低价:{4}\n涨幅% :{5}%\t涨幅  :{6}\n涨停价:{7}\t跌停价:{8}\n成交量:{9}万\t成交额:{10}亿\n\n".format(code, Stock_Name, New_Price, Highest_Price,
            Lowest_Price, diff_Value_100, diff_Value, format(limit_Raising_Price, '.2f'), format(Limit_Down_Price, '.2f'), format(Trans_quantity, '.2f'), format(Trans_crash, '.2f')),end='')
            time.sleep(2)
            os.system("cls")


if __name__ == '__main__':

    while 1:
        currTime_H = int(time.strftime('%H', time.localtime()))
        currTime_M = int(time.strftime('%M', time.localtime()))/60

        try:
            if currTime_H + currTime_M < 9.4 or currTime_H + currTime_M >= 15.5:
                if currTime_H + currTime_M >= 15.5:
                    updata2xls()
                break

            if currTime_H + currTime_M > 11.5 and currTime_H < 13:
                continue
            
            if 0:
                th_html = threading.Thread(target=get_html)
                th_html.setName("get_html")

                th_log = threading.Thread(target=print_log)
                th_log.setName("print_log")

                th_html.start()
                th_html.join()

                # th_log.setDaemon(True)   #把子进程设置为守护线程，必须在start()之前设置
                th_log.start()
                th_log.join() # 主线程会等待子线程结束
            else:
                get_html()
                print_log()
                print()

        except Exception as e:
            # print("\n\rMain Error "+str(e))
            continue
        
