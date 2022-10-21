"""
Time:    2022/10/11
Author:  Li YuanJun
Version: V3.0
"""

import json
import re
import copy
import logging
import threading
import qlyq2
import sms
import windnd
import ctypes
import other
from tkinter import *
from tkinter import ttk
import tkinter.font as tf
import tkinter.messagebox
from win32com import client as wc


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



# @function: 获取“段落编号”和“其特征在于（逗号）”之间的数字
# @param string: “段落编号”和“其特征在于（逗号）”之间的字符串
# return 提取到的数字，存放至列表中
def getNum(string):
    intList = []
    numList = []

    index1 = string.find("权利要求")
    index2 = string.find("所述")
    if index1 != -1 and index2 != -1:
        string = string[index1+4:index2]  # “权利要求”和“所述”之间的字符串
        for s in string:
            if s.isdigit():
                numList.append(s)
            else:
                num = ''.join(numList)
                if num != "":
                    intList.append(int(num))
                numList.clear()
        if len(numList) != 0:
            num = ''.join(numList)
            intList.append(int(num))
        
        # 考虑此类情况 权利要求1至3或者权利要求1~3，["至","-","到","~"]
        andWord = ["至","-","到","~"]
        if len(intList) >=2:
            for i in range(len(intList) - 1):
                index1 = string.find(str(intList[i]))
                index2 = string.find(str(intList[i+1]))
                key = string[index1+1:index2]
                if key in andWord:
                    if intList[i+1] - intList[i] >=2:
                        n = intList[i]
                        while n < intList[i+1] - 1:
                            n = n + 1
                            intList.append(n)
    return intList


# @function: 将段落编号格式化key-value形式
# @param pClist nlist: 段落编号列表， 引用数字列表
# return key-value
def getFormat(pCList, nlist):
    strValue = ""
    if len(pCList) == len(nlist):
        strValue = '{'
        for i in range(len(pCList)):
            strValue = strValue + '"' + \
                str(pCList[i]) + '"' + ':' + str(nlist[i])
            if i != len(pCList) - 1:
                strValue = strValue + ','
        strValue = strValue + '}'
    jsonValue = json.loads(strValue)
    logging.info("jsonValue: %s", jsonValue)

    # 剔除前面的权利引用后面权利的情况（eg: 权利3引用了权利4这种情况）
    for key, value in jsonValue.items():
        tempList = copy.deepcopy(value)  # 深拷贝
        for v in value:
            if int(key) <= v:
                tempList.remove(v)  # 删除掉超前引用权利数值
                tempList.append(-1) # 添加标记作为后面判断权利关系混乱的标记,不能直接添加到value列表中
        jsonValue[str(key)] = tempList  # 替换keyvalue列表

    logging.info("after-json: %s",jsonValue)

    newStr = "{"
    for key, value in jsonValue.items():
        for v in value:
            # 应对 key = -1的情况
            if v > 0:
                l = jsonValue[str(v)]
                value.extend(l)
        value = list(set(value))
        newStr = newStr + '"' + key + '"' + ':' + str(value) + ','
    newStr = newStr[:-1]
    newStr = newStr + '}'

    return newStr





# 主函数入口
def main_qc():

    qlyqContent = ""

    def getPartContent(doc, word, indexList, beforeMC, afterMC):
        Content = ""
        if len(indexList) != 0:
                for index in indexList:
                    for i in range(len(word.ActiveDocument.Sections[index].Range.Paragraphs)):
                        if str(word.ActiveDocument.Sections[index].Range.Paragraphs[i]).strip() != "":
                            parag = word.ActiveDocument.Sections[index].Range.Paragraphs[i]
                            paraNum = parag.Range.ListFormat.ListValue   # 段落编号
                            paramNumStr = "" if paraNum == 0 else str(paraNum) + "."
                            Content = Content + paramNumStr + str(word.ActiveDocument.Sections[index].Range.Paragraphs[i]).strip() + '\n'
                    

        else:
            isFlag11 = False
            isFlag12 = True
            titleText = ''
            for i in range(len(doc.Paragraphs)):
                paragraphString = ''.join([char for char in str(doc.Paragraphs[i]) if u'\u4e00' <= char <= u'\u9fa5'])  # “权 利 要 求 书”或者“，权 利 要 求 书”的格式，需要将逗号和空格去掉
                logging.info("doc.Paragraphs: %s",paragraphString.strip())
                if paragraphString.strip() == beforeMC:
                    isFlag11 = True   
                    titleText = str(doc.Paragraphs[i]).strip()  # 未替换逗号或者空格之前的“说明书摘要”或者“权利要求书”或者“说明书”
                if paragraphString.strip() == afterMC:
                    isFlag12 = False
                if isFlag11 and isFlag12:
                    if str(doc.Paragraphs[i]).strip() != "":
                        parag = doc.Paragraphs[i]
                        paraNum = parag.Range.ListFormat.ListValue   # 段落编号
                        paramNumStr = "" if paraNum == 0 else str(paraNum) + "."
                        Content = Content + paramNumStr + str(doc.Paragraphs[i]).strip() + '\n'
            l = len(titleText) + 1
            if len(Content) > l:
                Content = Content[l:]

            # 若还未定位到想要的段落，则通过其他方式获取到相应的段落
            if Content == "":
                # 获取整篇文章的内容
                AllText = ""
                for i in range(len(doc.Paragraphs)): 
                    AllText = AllText + str(doc.Paragraphs[i]).strip() + '\n'
                logging.info("AllText : %s",AllText)
                if len(AllText) != 0:
                    oneList = ['1.','1．','1、','1 ']    # 段落编号权利要求1（数字+.或者数字+．或者数字+、或者数字+空格）
                    jsly = "技术领域"
                    oneIndex = -1
                    jslyIndex = -1
                    for one in oneList:
                        index = AllText.find(one)
                        if index != -1:
                            oneIndex = index
                    # 获取说明书关键字“技术领域”
                    index = AllText.find(jsly)
                    if index != -1:
                        jslyIndex = index
                    
                    if beforeMC == "说明书摘要":
                        # 没找到权利要求1，也没找到说明书中的关键字“技术领域”，则全文都是说明书摘要
                        if oneIndex == -1 and jslyIndex == -1:
                            index = AllText.find("说明书摘要")
                            if index != -1:
                                Content = AllText[index+5:]
                            else:
                                Content = AllText
                        if oneIndex != -1:
                            index = AllText.find("说明书摘要")
                            if index != -1:
                                if index + 5 <= oneIndex:
                                    Content = AllText[index+5:oneIndex]
                            else:
                                Content = AllText[:oneIndex]
            
                    if beforeMC == "权利要求书":
                        if oneIndex != -1:
                            if jslyIndex == -1:
                                Content = AllText[oneIndex:]
                            else:
                                step = 0
                                num = 0
                                while jslyIndex -step >=0:
                                    targetStr = AllText[jslyIndex - step]
                                    step = step + 1
                                    if targetStr == '\n':
                                        num = num + 1
                                    if num == 2:
                                        endqlyqIndex = jslyIndex - step
                                        break
                                if oneIndex < endqlyqIndex:
                                    Content = AllText[oneIndex:endqlyqIndex]

                    if beforeMC == "说明书":
                        if jslyIndex != -1:
                            Content = AllText[jslyIndex:]
                            step = 0
                            num = 0
                            smsIndex = jslyIndex
                            while jslyIndex - step >=0:
                                targetStr = AllText[jslyIndex - step]
                                step = step + 1
                                if targetStr == '\n':
                                    num = num + 1
                                if num == 2:
                                    smsIndex = jslyIndex - step
                                    break
                            Content = AllText[smsIndex:]

        logging.info("Content: %s",Content)

        return Content

    


    ####################点击“开始检查”按钮跳转到此处###############

    # 多线程同时执行两个功能(1、检查权利要求说明书，2、检查说明书)
    def QLYQ(qC,wL,ff):
        # global resJson
        qlyq2.checkQLYQ(qC,wL,ff)
    
    def SMS(sC, fN, zN,ff):
        sms.main_sms(sC, fN, zN,ff)

    def OTHER(allS, wl, hEW):
        other.main_other(allS, wl, hEW)



    def checkQuality(wordhandle, file_path,funcNum):
        
        global qlyqContent
        try:
            doc = wordhandle.Documents.Open(file_path)
        except Exception as e:
            logging.info("Open Document has error!!")
        else:
            headNameList = []
            sections = wordhandle.ActiveDocument.Sections  # 所有页眉
            for i in range(len(sections)):
                name = wordhandle.ActiveDocument.Sections[i].Headers[0]
                spName = ''.join([char for char in str(name) if u'\u4e00' <= char <= u'\u9fa5'])  # 提取页眉,有些页眉带有空格或者其他标点符号
                if spName != "":
                    headNameList.append(spName)
            print("所有的页眉:%s",headNameList)


            smszyIndex = [] 
            qlyqsIndex = []
            smsIndex = []

            for i in range(len(headNameList)):
                string= headNameList[i]
                if "权利要求书" == string:
                    qlyqsIndex.append(i)
                if "说明书摘要" == string:
                    smszyIndex.append(i)
                if "说明书" == string:
                    smsIndex.append(i)

            ###  funcNum == 2的时候，全部检查
            if funcNum == 2:
                # 1.获取说明书摘要内容
                smszyContent = getPartContent(doc,wordhandle,smszyIndex,"说明书摘要","权利要求书")
                logging.info("XXXXXXXsmszyContent: %s",smszyContent)

                # 2.获取说明书内容
                smsContent = getPartContent(doc,wordhandle,smsIndex,"说明书","XXXX")
                logging.info("XXXXXXXsmsContent: %s",smsContent)

            # 3.获取权利要求部分内容
            qlyqContent = getPartContent(doc,wordhandle,qlyqsIndex,"权利要求书","说明书")
            logging.info("XXXXXXXqlyqContent: %s",qlyqContent)
           
            
   
            if funcNum == 2:
                # Function1: 判断说明书摘要文字个数
                wordLenth = True
                smszyContent = smszyContent.replace('\n','')  # 说明书摘要含有多行文字时，去除换行符
                logging.info("smszyContent:%s",smszyContent)
                logging.info("smszy length:%s",len(smszyContent.strip()))
                hasEnglishWord = False
                if len(smszyContent.strip()) > 300:
                    wordLenth = False
                    hasEnglishWord = bool(re.search('[a-zA-Z]', smszyContent))
                    logging.info("hasEnglishWord:%s",hasEnglishWord)
                    # tkinter.messagebox.showinfo("提示","说明书摘要超过300字,请做进一步检查！")



            # Function2: 判断权利要求部分
            wordLen = entry1.get()  # 最短截词长度
            # 获取附图标记说明，并进行格式化
            targetText1 = "附图标记说明"
            targetText12 = "附图标记"
            targetText13 = "标号说明"
            targetText2 = "具体实施方式"

            ftbjList = []  # 附图标记说明
            formatftbj = []  # 对不同格式的附图标记说明进行格式化

            try:
                table = doc.Tables(1)   # 1 代表文中的第一个表格，如果存在多个表格，默认第一个表格就是附图标记
            except Exception as e:
                # 当文中没有表格时，则采用下面方式获取附图标记
                logging.info("Read table has error: %s", e)
                startIndexList = []  # 附图标记说明和附图标记所在段落行数
                endIndexList = []     # 具体实施方式所在段落行数
                for i in range(len(doc.paragraphs)):
                    paraString = str(doc.paragraphs[i]).strip()
                    if paraString.find(targetText1) != -1 or paraString.find(targetText12) != -1 or paraString.find(targetText13) != -1:
                        startIndexList.append(i)
                    if paraString == targetText2:
                        endIndexList.append(i)

                # 段落中存在多个附图标记/附图标记说明/具体实施方式
                startIndex = endIndex = 0
                if len(startIndexList) != 0 and len(endIndexList) != 0:
                    endIndex = endIndexList[0]  # 第一个出现“具体实施方式”的段落
                    # 寻找比具体实施方式段落小，且最靠近具体实施方式的附图标记说明/附图标记段落
                    numberList = []
                    for number in startIndexList:
                        if number < endIndex:
                            numberList.append(number)
                    logging.info("numberList:%s",numberList)
                    numberList.sort(reverse=True)
                    startIndex = numberList[0]

                
                if startIndex != 0 and endIndex != 0:
                    j = startIndex + 1
                    while startIndex <= j < endIndex:
                        text = str(doc.paragraphs[j]).strip()
                        if text != '':
                            new_text = text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                            ftbjList.append(new_text)
                        j = j + 1

                logging.info("ftbjList：%s", ftbjList)
            else:
                numSymbolStringList = []  # 标号和含义处于同一个表格里（eg  12:功率单元， 此字符串占用一个小格子）
                numColumnList = []    # 数字所在的列
                stringColumnList = [] # 词语所在的列
                sumColumnsList = []   # 所有列数
                
                # 所有的列数
                sumColumns = table.Columns.Count  #总共的列数
                sumRows = table.Rows.Count  # 总共的行数

                logging.info("sumRows %s", sumRows)
                logging.info("sumColumns： %s", sumColumns)

                for colum in range(1, sumColumns + 1):
                    sumColumnsList.append(colum)
                logging.info("sumColumnsList %s", sumColumnsList)

                # 找到是标号的所在的列
                for row in table.Rows:  #遍历表格每行
                    colum = 0
                    for cell in row.Cells:  #遍历每行中的有效列
                        colum = colum + 1
                        tableText = cell.Range.Text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                        if tableText != "":
                            numSymbolStringList.append(tableText)
                            if str(tableText).encode('UTF-8').isalnum(): # 所有字段仅是数字或者英文字母
                                numColumnList.append(colum) 
                numColumnList = list(set(numColumnList)) # 删除重复的列数
                numColumnList.sort()

                if len(numColumnList) > 0:
                    for num in sumColumnsList:
                        if num not in numColumnList:
                            stringColumnList.append(num)
                    logging.info("numColumnList:%s",numColumnList)
                    logging.info("stringColumnList:%s",stringColumnList)
                    
                    if len(numColumnList) == len(stringColumnList):
                        for i in range(len(numColumnList)):
                            j = 1
                            while j <= sumRows:
                                num = table.Rows(j).Cells(numColumnList[i]).Range.Text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                                string = table.Rows(j).Cells(stringColumnList[i]).Range.Text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                                j = j + 1
                                new_num = re.sub(r"~|-", "",num)  # 特殊情况，例如 101-106：纸卷 或者 101~106：纸卷， 为了方便判断，暂时先去掉-和~
                                if new_num.encode('UTF-8').isalnum(): # 字段仅是数字或者英文字母 （有些表格第一行的列中的字符分别是“标号”和“含义”，而不是 数字和词语，所以这种情况下要考虑到）
                                    combineStr = num + '：' + string # 组合后的字段（ 数字：词语）
                                    ftbjList.append(combineStr)
                else:
                    ftbjList = numSymbolStringList  # 没有任何仅有数字或者字母的小表格，说明表格里的形式是：数字（字母）+标点符号+词语，或者词语+标点符号+数字（字母），不做任何处理

                logging.info("table info:%s",ftbjList)

            # 对不同的附图标记说明格式进行格式化
            for text in ftbjList:
                tempList = re.split(r';|；|。', text)
                # print(tempList)
                for strValue in tempList:
                    if strValue != '':
                        formatftbj.append(strValue)
            logging.info("formatftbj：%s",formatftbj)
            
            
            if funcNum == 2:
                #Function3: 判断说明书部分
                ftNum = int(entry2.get())   # 附图数目
                zhaiyaoName = ""
                firstIndex = smszyContent.find("一种")
    
                if firstIndex != -1:
                    symbolIndex = [substr.start() for substr in re.finditer(r"，|。|,|；|;", smszyContent)]
                    logging.info("symbolIndex:%s", symbolIndex)    
                    if len(symbolIndex) >= 1:
                        zhaiyaoName = smszyContent[firstIndex + 2:symbolIndex[0]]
                logging.info("zhaiyaoName:%s", zhaiyaoName)    
                # sms.main_sms(smsContent, ftNum, zhaiyaoName)


                # Function4: 判断错别字，重复字、实用新型和发明
                allString = ''
                for i in range(len(doc.paragraphs)):
                    allString = allString + str(doc.paragraphs[i]).strip()
                logging.info("allStr: %s", allString)

    
            ##function:只检查权利要求部分
            if funcNum == 1:
                # qlyq2.checkQLYQ(qlyqContent,wordLen,formatftbj)
                t1 = threading.Thread(target=QLYQ,args=(qlyqContent,wordLen,formatftbj))
                t1.start()

            ##function:全部检查
            if funcNum == 2:
                # 使用threading模块，threading.Thread()创建线程，其中target参数值为需要调用的方法，args参数值为要传递的参数，同样将其他多个线程放在一个列表中，遍历这个列表就能同时执行里面的函数了
                t1 = threading.Thread(target=QLYQ,args=(qlyqContent,wordLen,formatftbj))
                t2 = threading.Thread(target=SMS,args=(smsContent, ftNum, zhaiyaoName,formatftbj))
                t3 = threading.Thread(target=OTHER,args=(allString, wordLenth, hasEnglishWord))
                # 启动线程
                t1.start()
                t2.start()
                t3.start()
               
                


            # qlyq2.checkQLYQ(qlyqContent,wordLen,formatftbj)
            # sms.main_sms(smsContent, ftNum, zhaiyaoName,formatftbj)
            # other.main_other(allString, wordLenth, hasEnglishWord)


            
    ###################################################

    # 1.按钮：引用关系
    def getResult():
        # global resJson
        # logging.info("XXXXXXX:%s", resJson)
        root1 = Tk(className="引用关系")
        root1.wm_attributes('-topmost', 1)
        screenwidth, screenheight = root1.maxsize()
        width = 600
        height = 400
        size = '%dx%d+%d+%d' % (width, height, (screenwidth -
                                width) / 4, (screenheight - height) / 4)
        root1.geometry(size)
        root1.resizable(0, 0)

        lable = Label(root1, height=2)
        lable['text'] = ""
        lable.pack()

        texty = Text(root1)
        texty.pack()

        global qlyqContent
        stringValue = qlyqContent  # 获取上面文本框中的全部内容
        logging.info("stringValue: %s", stringValue)
        indexList = [substr.start() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]       # 段落编号首个数字所在位置(有全角和半角的英文句号，所以使用(\.|\．|\、| ))
        paraNumList = [substr.group().strip() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]    # 段落编号（数字+.或者数字+．或者数字+、或者数字+空格）
        logging.info("paraNumList: %s", paraNumList)   
        logging.info("indexList: %s", indexList)

        ### 正文内容中出现过这样的内容("2.34 - 5的常规变量")， finditer会将正文中的'2.'当成段落编号，所以需要进行处理
        sPos = 1.0
        allPostion = []
        texty.insert(1.0,qlyqContent)
        for paraNum in paraNumList:
            pos = texty.search(paraNum,sPos,END)
            logging.info("pos: %s" , pos)
            if pos == '':
                continue
            else:
                allPostion.append(pos)
                sPos = '%s+%dc' % (pos, len(paraNum))
        logging.info("allPostion: %s" ,allPostion)

        #首先判断段落编号是否存在重复，若有重复，则进行特殊处理,若无重复，则无需进行特殊处理
        
        deleteIndexList = []
        for i in range(len(allPostion)):
            rowColumList = allPostion[i].split(".")
            if rowColumList[1] != '0':
                deleteIndexList.append(i)
        deleteIndexList.reverse()
        logging.info("deleteIndex: %s" ,deleteIndexList)     
        
        for deleteIndex in deleteIndexList:
            del indexList[deleteIndex]
            del paraNumList[deleteIndex]

        paracodeList= [re.sub(r"\.|\．|\、| ","",element) for element in paraNumList]
        ##################
        

        numList = []        # 存放对应段落中的引用
        targetStrList = []  # 保存所有的权利要求和逗号之间的字符串
        paragList = []      # 保存所有的段落
        for i in range(len(indexList)):
            if i == len(indexList) - 1:
                parag = stringValue[indexList[i]:]
            else:
                parag = stringValue[indexList[i]:indexList[i+1]]
            logging.info("Para Text: %s",parag)
            paragList.append(parag)
            targetIndex = 0  # 初始化
            qtzzyIndex = str(parag).find("其特征在于")  # 找到第一个“其特征在于”字符所在位置
            if qtzzyIndex == -1:
                douHaoIndex = str(parag).find("，")  # 如果没找到“其特征在于”，就找第一个逗号
                if douHaoIndex != -1:
                    targetIndex = douHaoIndex
            else:
                targetIndex = qtzzyIndex

            # 段落编号和“其特征在于”或者“逗号”之间的字符串
            targetStr = stringValue[indexList[i]:indexList[i] + targetIndex]
            logging.info("targetStr: %s", targetStr)  
            resList = getNum(targetStr)
            logging.info("resList: %s", resList)
            numList.append(resList)
        logging.info("getRefList paracodeList: %s",paracodeList)  # 段落编号
        logging.info("getRefList numList: %s", numList)  
        
        res = getFormat(paracodeList, numList)
        resJson = json.loads(res)
        # 把引用数值排序一下
        for key, value in resJson.items():
            resJson[key] = sorted(value)
        logging.info("Key-Value: %s", resJson)
        
        # resJson = getRefJson()
        res = ''
        for k, v in resJson.items():
            if len(v) == 0:
                res = res + "权利" + k + "为独立权利！" + '\n'
                res = res + '\n'
            else:
                all = ''
                for i in v:
                    if i != -1:
                        all = all + str(i) + '、'
                if len(all) == 0:   
                    res = res + "权利" + k + "引用关系混乱" + '\n' 
                else:
                    all = all[:-1]
                    res = res + "权利" + k + "引用:" + all + '\n'
                res = res + '\n'
        logging.info("res: %s",res) 
        texty.delete(1.0,END)
        texty.insert(1.0, res)  # 1.0 代表文本框中第一行，第1列（行是从1开始，列是从0开始）
        texty.configure(font=("微软雅黑", 14))  # 设置文本框中的字体
        # root1.protocol("WM_DELETE_WINDOW", close_callback)
        root1.mainloop()
        # getRef("引用关系","", res)
        

    # 2.按钮：引用基础词语
    def refBase():
        # 添加新的名称
        def insert():
            # 获取Entry文本框中的文字
            data = entryEdit.get()
            table.insert('', END, values=data)  # 添加数据到末尾
            # 添加到json文件中，以永久保存数据
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                yyjcList = jsonInfo['YYJC']
                logging.info("json data: %s", yyjcList)
                yyjcList.append(data)   # 添加数据之后的list
                axis = {"YYJC":yyjcList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中

        def delete():
            global rowId
            table.delete(rowId)
            logging.info("Original ID: %s",rowId)
            idInt = int(rowId[0][1:],16)  # rowId为16进制，转换为10进制
            logging.info("Int ID: %s", idInt)
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                yyjcList = jsonInfo['YYJC']
                logging.info("json data: %s", yyjcList)
                del yyjcList[idInt - 1]   # 删除数据之后的list
                axis = {"YYJC":yyjcList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中
        
           
       
        win = tkinter.Tk()  # 窗口
        win.title('引用基础词语')  # 标题

        # 添加新的引用基础词语
        btn_frame = Frame(win)
        btn_frame.pack(side = TOP)

        entryEdit = Entry(btn_frame, width=20)
        entryEdit.grid(row=0,column=0,padx=(2, 0),pady=(2, 2))
        
        bAdd = Button(btn_frame, text='添加', width=8, command=insert)
        bAdd.grid(row=0,column=1,padx=(2, 0),pady=(2, 2))

        bDel = Button(btn_frame, text='删除', width=8, command=delete)
        bDel.grid(row=0,column=2,padx=(4, 0),pady=(2, 2))

         # 创建表格
        tabel_frame = Frame(win)
        tabel_frame.pack()

        yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

        # columns = ['系统无领域主题名称', '用户自定义']
        columns = ['引用基础词语']
        table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=15,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                yscrollcommand=yscroll.set,  # y轴滚动条
                )
        for column in columns:
            table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
            table.column(column=column, width=300, minwidth=300, anchor=CENTER, )  # 定义列
        yscroll.config(command=table.yview)
        yscroll.pack(side=RIGHT, fill=Y)

        table.pack(fill=BOTH, expand=True)

        # 添加json文件中的数据到表格
        allInfo = readJsonFile()
        if allInfo != "":
            info = allInfo['YYJC']
            logging.info(info)

            for index, data in enumerate(info):
                table.insert('', END, values=data)  # 添加数据到末尾

        # 选中表格中的某行数据并获取到该行的ID
        def selectjob(event):
            global rowId
            rowId = table.selection()
            # items = table.set(rowId[0])
        table.bind("<<TreeviewSelect>>",selectjob)
        # win.protocol("WM_DELETE_WINDOW", close_callback)
        win.mainloop()

    # 3. 按钮： 小标题管理
    def subTitle():

        # 添加新的名称
        def insert():
            # 获取Entry文本框中的文字
            data = entryEdit.get()
            table.insert('', END, values=data)  # 添加数据到末尾
            # 添加到json文件中，以永久保存数据
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                SUBTITLEList = jsonInfo['SUBTITLE']
                logging.info("subTitle json data: %s", SUBTITLEList)
                SUBTITLEList.append(data)   # 添加数据之后的list
                axis = {"SUBTITLE":SUBTITLEList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中

        def delete():
            global rowId
            table.delete(rowId)
            logging.info("Original ID: %s",rowId)
            idInt = int(rowId[0][1:],16)  # rowId为16进制，转换为10进制
            logging.info("Int ID: %s", idInt)
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                SUBTITLEList = jsonInfo['SUBTITLE']
                logging.info("subTitle json data: %s", SUBTITLEList)
                del SUBTITLEList[idInt - 1]   # 删除数据之后的list
                axis = {"subTitle":SUBTITLEList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中
        
           
       
        win1 = tkinter.Tk()  # 窗口
        win1.title('小标题')  # 标题

        # 添加新的敏感词
        btn_frame = Frame(win1)
        btn_frame.pack(side = TOP)

        entryEdit = Entry(btn_frame, width=20)
        entryEdit.grid(row=0,column=0,padx=(2, 0),pady=(2, 2))
        
        bAdd = Button(btn_frame, text='添加', width=8, command=insert)
        bAdd.grid(row=0,column=1,padx=(2, 0),pady=(2, 2))

        bDel = Button(btn_frame, text='删除', width=8, command=delete)
        bDel.grid(row=0,column=2,padx=(4, 0),pady=(2, 2))

         # 创建表格
        tabel_frame = Frame(win1)
        tabel_frame.pack()

        yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

        columns = ['小标题']
        table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=15,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                yscrollcommand=yscroll.set,  # y轴滚动条
                )
        for column in columns:
            table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
            table.column(column=column, width=300, minwidth=300, anchor=CENTER, )  # 定义列
        yscroll.config(command=table.yview)
        yscroll.pack(side=RIGHT, fill=Y)

        table.pack(fill=BOTH, expand=True)

        # 添加json文件中的数据到表格
        allInfo = readJsonFile()
        if allInfo != "":
            info = allInfo['SUBTITLE']
            logging.info(info)

            for index, data in enumerate(info):
                table.insert('', END, values=data)  # 添加数据到末尾

        # 选中表格中的某行数据并获取到该行的ID
        def selectjob(event):
            global rowId
            rowId = table.selection()
            # items = table.set(rowId[0])
        table.bind("<<TreeviewSelect>>",selectjob)
        # win1.protocol("WM_DELETE_WINDOW", close_callback)
        win1.mainloop()

    # 4. 按钮： 敏感词
    def warn():

        # bWarn['state'] = "disabled"

        # 添加新的名称
        def insert():
            # 获取Entry文本框中的文字
            data = entryEdit.get()
            table.insert('', END, values=data)  # 添加数据到末尾
            # 添加到json文件中，以永久保存数据
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                MGCList = jsonInfo['MGC']
                logging.info("json data: %s", MGCList)
                MGCList.append(data)   # 添加数据之后的list
                axis = {"MGC":MGCList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中

        def delete():
            global rowId
            table.delete(rowId)
            logging.info("Original ID: %s",rowId)
            idInt = int(rowId[0][1:],16)  # rowId为16进制，转换为10进制
            logging.info("Int ID: %s", idInt)
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                MGCList = jsonInfo['MGC']
                logging.info("json data: %s", MGCList)
                del MGCList[idInt - 1]   # 删除数据之后的list
                axis = {"MGC":MGCList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中
        
           
       
        win = tkinter.Tk()  # 窗口
        win.title('敏感词')  # 标题

        # 添加新的敏感词
        btn_frame = Frame(win)
        btn_frame.pack(side = TOP)

        entryEdit = Entry(btn_frame, width=20)
        entryEdit.grid(row=0,column=0,padx=(2, 0),pady=(2, 2))
        
        bAdd = Button(btn_frame, text='添加', width=8, command=insert)
        bAdd.grid(row=0,column=1,padx=(2, 0),pady=(2, 2))

        bDel = Button(btn_frame, text='删除', width=8, command=delete)
        bDel.grid(row=0,column=2,padx=(4, 0),pady=(2, 2))

         # 创建表格
        tabel_frame = Frame(win)
        tabel_frame.pack()

        yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

        # columns = ['系统无领域主题名称', '用户自定义']
        columns = ['敏感词']
        table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=15,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                yscrollcommand=yscroll.set,  # y轴滚动条
                )
        for column in columns:
            table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
            table.column(column=column, width=300, minwidth=300, anchor=CENTER, )  # 定义列
        yscroll.config(command=table.yview)
        yscroll.pack(side=RIGHT, fill=Y)

        table.pack(fill=BOTH, expand=True)

        # 添加json文件中的数据到表格
        allInfo = readJsonFile()
        if allInfo != "":
            info = allInfo['MGC']
            logging.info(info)

            for index, data in enumerate(info):
                table.insert('', END, values=data)  # 添加数据到末尾

        # 选中表格中的某行数据并获取到该行的ID
        def selectjob(event):
            global rowId
            rowId = table.selection()
            # items = table.set(rowId[0])
        table.bind("<<TreeviewSelect>>",selectjob)
        # win.protocol("WM_DELETE_WINDOW", close_callback)
        win.mainloop()

    # 5. 按钮： 重复词
    def dumWord():
        
        # bdumpWord['state'] = "disabled"

        # 添加新的重复词
        def insert():
            # 获取Entry文本框中的文字
            data = entryEdit.get()
            table.insert('', END, values=data)  # 添加数据到末尾
            # 添加到json文件中，以永久保存数据
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                DUMPList = jsonInfo['DUMP']
                logging.info("json data DUMPList: %s", DUMPList)
                DUMPList.append(data)   # 添加数据之后的list
                axis = {"DUMP":DUMPList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中

        def delete():
            global rowId
            table.delete(rowId)
            logging.info("Original ID: %s",rowId)
            idInt = int(rowId[0][1:],16)  # rowId为16进制，转换为10进制
            logging.info("Int ID: %s", idInt)
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                DUMPList = jsonInfo['DUMP']
                logging.info("json data: %s", DUMPList)
                del DUMPList[idInt - 1]   # 删除数据之后的list
                axis = {"DUMP":DUMPList}  # 添加数据之后的字典
                jsonInfo.update(axis)   # 将新的字典更新到json数据中
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中
        
        win = Tk()  # 窗口
        win.title('重复词')  # 标题
        win.resizable(0, 0)

        # 添加新的重复词
        btn_frame = Frame(win)
        btn_frame.pack(side = TOP)

        entryEdit = Entry(btn_frame, width=20)
        entryEdit.grid(row=0,column=0,padx=(2, 0),pady=(2, 2))
        
        bAdd = Button(btn_frame, text='添加', width=8, command=insert)
        bAdd.grid(row=0,column=1,padx=(2, 0),pady=(2, 2))

        bDel = Button(btn_frame, text='删除', width=8, command=delete)
        bDel.grid(row=0,column=2,padx=(4, 0),pady=(2, 2))

         # 创建表格
        tabel_frame = Frame(win)
        tabel_frame.pack()

        yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

        # columns = ['系统无领域主题名称', '用户自定义']
        columns = ['重复词']
        table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=15,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                yscrollcommand=yscroll.set,  # y轴滚动条
                )
        for column in columns:
            table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
            table.column(column=column, width=300, minwidth=300, anchor=CENTER, )  # 定义列
        yscroll.config(command=table.yview)
        yscroll.pack(side=RIGHT, fill=Y)

        table.pack(fill=BOTH, expand=True)

        # 添加json文件中的数据到表格
        allInfo = readJsonFile()
        if allInfo != "":
            info = allInfo['DUMP']
            logging.info("info: %s", info)

            for index, data in enumerate(info):
                table.insert('', END, values=data)  # 添加数据到末尾

        # 选中表格中的某行数据并获取到该行的ID
        def selectjob(event):
            global rowId
            rowId = table.selection()
            # items = table.set(rowId[0])
        table.bind("<<TreeviewSelect>>",selectjob)
        # win.protocol("WM_DELETE_WINDOW", close_callback)
        win.mainloop()

    # 6.按钮：错别字
    def errorWord():
        # berrorWord['state'] = "disabled"

        # def close_callback():
        #     berrorWord['state'] = 'normal'
        #     win.destroy()

        # 添加新的重复词
        def insert():
            # 获取Entry文本框中的文字  
            data1 = entryEdit1.get()   # 错别字
            data2 = entryEdit2.get()   # 正确字

            table.insert('', END, values=(data1, data2))  # 添加数据到末尾
            # 添加到json文件中，以永久保存数据
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                ERRORList = jsonInfo['ERROR']
                RIGHTList = jsonInfo['RIGHT']
                logging.info("json data ERRORList: %s", ERRORList)
                logging.info("json data RIGHTList: %s", RIGHTList)
                ERRORList.append(data1)   # 添加数据之后的list
                RIGHTList.append(data2)
                axis1 = {"ERROR":ERRORList}  # 添加数据之后的字典
                axis2 = {"RIGHT":RIGHTList}  # 添加数据之后的字典
                jsonInfo.update(axis1)   # 将新的字典更新到json数据中
                jsonInfo.update(axis2) 
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中

        def delete():
            global rowId
            table.delete(rowId)
            logging.info("Original ID: %s",rowId)
            idInt = int(rowId[0][1:],16)  # rowId为16进制，转换为10进制
            logging.info("Int ID: %s", idInt)
            jsonInfo = readJsonFile()
            if jsonInfo != "":
                ERRORList = jsonInfo['ERROR']
                RIGHTList = jsonInfo['RIGHT']
                logging.info("delete json data ERRORList: %s", ERRORList)
                logging.info("delete json data RIGHTList: %s", RIGHTList)
                del ERRORList[idInt - 1]   # 删除数据之后的list
                del RIGHTList[idInt - 1]
                axis1 = {"ERROR":ERRORList}  # 添加数据之后的字典
                axis2 = {"RIGHT":RIGHTList}  # 添加数据之后的字典
                jsonInfo.update(axis1)   # 将新的字典更新到json数据中
                jsonInfo.update(axis2)
                writeJsonFile(jsonInfo)  # 将json数据写入到json文件中
        
        win = Tk()  # 窗口
        win.title('错别字')  # 标题
        win.resizable(0, 0)

        # 添加新的重复词
        btn_frame = Frame(win)
        btn_frame.pack(side = TOP)

        #错别字字符
        label1 = Label(btn_frame)
        label1['text'] = "错别字:"
        label1.grid(row=0,column=0,padx=(5, 5),pady=(15, 15))

        # 错别字输入框
        entryEdit1 = Entry(btn_frame, width=20)
        entryEdit1.grid(row=0,column=1,padx=(5, 5),pady=(15, 15))

        #正确字字符
        label2 = Label(btn_frame)
        label2['text'] = "正确字:"
        label2.grid(row=0,column=2,padx=(5, 5),pady=(15, 15))

        # 正确字输入框
        entryEdit2 = Entry(btn_frame, width=20)
        entryEdit2.grid(row=0,column=3, padx=(5, 5),pady=(15, 15))
        
        bAdd = Button(btn_frame, text='添加', width=8, command=insert)
        bAdd.grid(row=0,column=4,padx=(5, 5),pady=(15, 15))

        bDel = Button(btn_frame, text='删除', width=8, command=delete)
        bDel.grid(row=0,column=5,padx=(5, 5),pady=(15, 15))

         # 创建表格
        tabel_frame = Frame(win)
        tabel_frame.pack()

        yscroll = Scrollbar(tabel_frame, orient=VERTICAL)

        # columns = ['系统无领域主题名称', '用户自定义']
        columns = ['错别字', '正确字']
        table = ttk.Treeview(
                master=tabel_frame,  # 父容器
                height=15,  # 表格显示的行数,height行
                columns=columns,  # 显示的列
                show='headings',  # 隐藏首列
                yscrollcommand=yscroll.set,  # y轴滚动条
                )
        for column in columns:
            table.heading(column=column, text=column, anchor=CENTER)  # 定义表头
            table.column(column=column, width=300, minwidth=300, anchor=CENTER, )  # 定义列
        yscroll.config(command=table.yview)
        yscroll.pack(side=RIGHT, fill=Y)

        table.pack(fill=BOTH, expand=True)

        # 添加json文件中的数据到表格
        allInfo = readJsonFile()
        if allInfo != "":
            info1 = allInfo['ERROR']
            info2 = allInfo['RIGHT']
            logging.info("info1: %s", info1)
            logging.info("info2: %s", info2)

            if len(info1) == len(info2):
                for i in range(len(info1)):
                    table.insert('', END, values=(info1[i], info2[i]))  # 添加数据到末尾


        # 选中表格中的某行数据并获取到该行的ID
        def selectjob(event):
            global rowId
            rowId = table.selection()
            # items = table.set(rowId[0])
        table.bind("<<TreeviewSelect>>",selectjob)
        # win.protocol("WM_DELETE_WINDOW", close_callback)
        win.mainloop()

           

    ######## 主按钮： 开始检查
    def batchFileCheck(funcNum):
        if funcNum == 1:
            qlyaCheck['fg'] = 'Turquoise'
        if funcNum == 2:
            bStartCheck['fg'] = 'Turquoise'

        word = wc.Dispatch("Word.Application")
        word.Visible = 0  # 0:后台运行，不显示； 1:打开文档，直接显示
        word.DisplayAlerts = 0  # 不警告

        wordPath = text.get(0.0,END)  # 文本框中的word路径
        wordPath = wordPath.strip('\n')
        logging.info("wordPath: %s", wordPath)

        if wordPath != "":
            wordReplacePath = wordPath.replace("\\","/")
            logging.info("wordReplacePath: %s", wordReplacePath)
            if wordReplacePath.endswith('.docx') or wordReplacePath.endswith('.doc'):
                checkQuality(word, wordReplacePath,funcNum)
                # tkinter.messagebox.showinfo('提示','检查完毕！！')



    # 拖放word文件
    def drag_word(files):
        text.delete(0.0,END)
        word_path = '\n'.join((item.decode('gbk') for item in files))
        text.insert(0.0,word_path)


    root = Tk(className="文字质量检查")

    # #调用api设置成由应用程序缩放
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    #调用api获得当前的缩放因子
    ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
    logging.info("ScaleFactor:%s", ScaleFactor)

    root.wm_attributes('-topmost', 0)
    screenwidth, screenheight = root.maxsize()
    scale = (ScaleFactor - 100)/100
    screenwidth = int(screenwidth*(1+scale))
    screenheight = int(screenheight*(1+scale))
    w = screenwidth/2
    h = screenheight/2 + 10
    size = '%dx%d+%d+%d' % (w, h, (screenwidth - w) / 2, (screenheight - h) / 2)
    root.geometry(size)
    root.resizable(0, 0)

    bWidth = 15
    lwidth = 10
    pad_x = 60 
    pad_y = 40

    # 第一行区域
    frame1 = Frame(root)
    frame1.pack(pady=(pad_y, pad_y))

    button11 = Button(frame1, width=bWidth,text="引用关系", command=lambda: getResult())
    button11.grid(row=0, column=0, padx=(pad_x, 0))


    button12 = Button(frame1, width=bWidth,text="引用基础词语", command=lambda: refBase())
    button12.grid(row=0, column=1, padx=(pad_x, 0))

    button13 = Button(frame1, width=bWidth, text="小标题管理", command=lambda:subTitle())
    button13.grid(row=0, column=2, padx=(pad_x, 0))


    # 截词长度文字
    label = Label(frame1, width=lwidth, height=2, text="截词长度",fg="#FF0000",font=("微软雅黑", 11, "bold"))
    label.grid(row=0, column=4, padx=(pad_x, 0))
    # 截词长度输入框
    entry1 = Entry(frame1, width=3,fg="#FF0000")
    entry1.insert(0,2)
    entry1.grid(row=0, column=5)


    # 第二行区域
    frame2 = Frame(root)
    frame2.pack()

    button21 = Button(frame2, width=bWidth, text="敏感词", command= lambda: warn())
    button21.grid(row=0, column=0, padx=(pad_x, 0))


    button22 = Button(frame2, width=bWidth,text="重复词", command= lambda: dumWord())
    button22.grid(row=0, column=1, padx=(pad_x, 0))

    button23 = Button(frame2, width=bWidth,  text="错别字", command= lambda: errorWord())
    button23.grid(row=0, column=2, padx=(pad_x, 0))


    # 附图数
    labe2 = Label(frame2, width=lwidth, height=2, text="附图数",fg="#FF0000",font=("微软雅黑", 11, "bold"))
    labe2.grid(row=0, column=4, padx=(pad_x, 0))

    # 截词长度输入框
    entry2 = Entry(frame2, width=3,fg="#FF0000")
    entry2.insert(0,2)
    entry2.grid(row=0, column=5)

    # 第三行区域
    frame3 = Frame(root)
    frame3.pack()

    # 设置标签1
    lable1 = Label(frame3, height=2)
    lable1['text'] = "请将Word拖放到下面文本框中"
    lable1.pack(pady=(30, 0))
    
    #设置文本框
    text = Text(frame3, height= 10, width = 80)
    text.pack(padx=(20,0))

    # 拖放Word文件
    windnd.hook_dropfiles(text, func=drag_word)

    # 第四行区域
    frame4 = Frame(root)
    frame4.pack(pady=(15,0))

    qlyaCheck = Button(frame4, width = 30, text='仅检查权利要求部分', command= lambda: batchFileCheck(1))
    qlyaCheck.grid(row=0, column=1,padx=(0,10))

    bStartCheck = Button(frame4, width = 30, text='全部检查', command= lambda: batchFileCheck(2))
    bStartCheck.grid(row=0, column=2,padx=(10,0))


    root.mainloop()


# 主入口函数
# if __name__ == '__main__':
#     main_qc()

