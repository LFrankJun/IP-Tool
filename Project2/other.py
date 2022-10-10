"""
Time:    2022/08/08
Author:  Li YuanJun
Version: V0.1
File:    main.py
Describe: other issue
"""


from tkinter import *
import tkinter.messagebox
from win32com import client as wc
import logging
from tkinter import ttk
import json
import re




logging.getLogger().setLevel(logging.INFO)

# 读取Json文件中的你内容
def readJsonFile():

    info = ''
    try:
        with open("name.json", encoding='UTF8') as file:
                info = json.load(file)
    except Exception as e:
        logging.info("Read JSON File Error: %s", str(e))

    logging.info("Json Info: %s", info) 
    return info

# 写入数据到Json文件中
def writeJsonFile(mc):
    info = ""
    try:
        with open("name.json", 'w',encoding='UTF8') as file:
                json.dump(mc,file,ensure_ascii=False)
    except Exception as e:
        logging.info("Write JSON File Error: %s", str(e))




# 具体实现寻找错误
def main_other(allString, wL, hEW):

    isOK1 = True   # “发明”和“实用新型”同时出现
    isOk2 = True   # 重复词语
    isOk3 = True   # 错别字修改
 

    inventList = [substr.start() for substr in re.finditer("发明", allString)]
    shiyong = [substr.start() for substr in re.finditer("实用新型", allString)]

    if len(inventList) != 0 and len(shiyong) != 0:
        isOK1 = False

    jsonInfo = readJsonFile()

    dumWordList = []
    errorWordList = []
    if jsonInfo != "":
        DUMPList = jsonInfo['DUMP']
        ErrorList = jsonInfo['ERROR']
        RIGHTList = jsonInfo['RIGHT']
        logging.info("searchError json data DUMPList: %s", DUMPList)
        logging.info("searchError json data ErrorList: %s", ErrorList)
        logging.info("searchError json data RIGHTList: %s", RIGHTList)
        # 判断是否存在重复字
        for key in DUMPList:
            dumWordIndexList = [substr.start() for substr in re.finditer(key, allString)]
            if len(dumWordIndexList) != 0:
                isOk2 = False  
                dWL = [substr.group() for substr in re.finditer(key, allString)]
                dumWordList.append(dWL)
                logging.info("dumWordList: %s", dumWordList)

        # 判断错别字
        if len(ErrorList) == len(RIGHTList):
            for j in range(len(ErrorList)):
                errorWordIndexList = [substr.start() for substr in re.finditer(ErrorList[j], allString)]   # 判断是否存在错别字
                if len(errorWordIndexList) != 0:
                    isOk3 = False
                    eWL = [substr.group() for substr in re.finditer(ErrorList[j], allString)]
                    errorWordList.append(eWL)
                    logging.info("errorWordList: %s", errorWordList)


    r1 = ""
    r2 = ""
    r3 = ""
    r4 = ""
    if not isOK1:
        r1 = '“发明”和“实用新型”同时存在的错误,'
    if not isOk2:
        r2 = '“重复词语”' + str(dumWordList) + ','  
    if not isOk3:
        r3 = '“错别字”' + str(errorWordList) + ','
    
    if not wL:
        r4 = '说明书摘要超过300字，'
        if hEW:
            r4 = r4 + '因含有英文字符，无法确保计数完全正确,'

    result = 'Word文件中出现了' + r1 + r2 + r3 + r4+ '请做进一步检查!'
    logging.info("result: %s",result)
    if isOK1 and isOk2 and isOk3 and wL:
        pass
    else:
        tkinter.messagebox.showinfo('提示',result)


# if __name__ == '__main__':
#     main_other()
