"""
Time:    2022/10/11
Author:  Li YuanJun
Version: V3.0
Func: None
"""

import json
import re
import copy
import ctypes
import logging
from tkinter import *
from tkinter import ttk
import tkinter.font as tf
import tkinter.messagebox



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


# @function: 多个权利要求和所述存在时，选择距离权利要求最近的那个所述
# @param value: 权利要求所在的位置
# @param sList: 所述所在的位置列表
# return 跟随在权利要求后面的所述位置
def getIndex(value, sList):
    i = value   # 初始化
    for v in sList:
        if v > value:
            i = v
            break
    return i


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


# @function: 获取“权利要求”和“所述”之间的数字
# @param string: 权利要求和逗号之间的字符串（考虑此类情况 权利要求1至3或者权利要求1~3，["至","-","到","~"]）
# return 提取到的数字，存放至列表中
def getNum2(string):
    index = string.find("所述")
    if index != -1:
        string = string[:index]
    numList = []
    intList = []
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
# @param
# return key-value


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
            if int(key) == v:
                tempList.remove(v)  # 删除掉超前引用权利数值
                # tempList.append() # 添加标记作为后面判断权利关系混乱的标记,不能直接添加到value列表中
        jsonValue[str(key)] = tempList  # 替换keyvalue列表

    logging.info("after-json: %s",jsonValue)

    # 将权利要求和其引用的所有权利放置在一起（eg:1:[],2:[1],3:[1,2],4:[1,2,3]）
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
    logging.info("newStr: %s", newStr)

    return newStr




# main入口函数

def checkQLYQ(Content,wordLen,formatftbjList):

    
    # 检查文章第一段是否缺少权利要求号码“1.”,只关注第一个段落即可
    def checkFirstParaNum():
        stringValue = text1.get(1.0,END) # 获取上面文本框中的全部内容
        logging.info("text1.get(1.0,END):%s",stringValue)
        firstParaIndex1 = stringValue.find("1.")
        firstParaIndex2 = stringValue.find("1．")
        firstParaIndex3 = stringValue.find("1、")
        logging.info("firsParaIndex: %s", firstParaIndex1)
        if firstParaIndex1 == -1 and firstParaIndex2 == -1 and firstParaIndex3 == -1:
            tkinter.messagebox.showinfo("提示","不存在权利要求1，请确保各权利要求编号格式为“1”、“2”、“3”....")
        # if (firstParaIndex1 != -1 and firstParaIndex1 != 0) or (firstParaIndex2 != -1 and firstParaIndex2 != 0) or (firstParaIndex3 != -1 and firstParaIndex3 != 0):
        #     tkinter.messagebox.showinfo("提示","权利要求1的编号前面存在空格等不需要的字符，请去Word文件中删除！")
    
    
    # 将获取每一段的权利要求引用数值
    def getRefList():
        stringValue = text1.get(1.0,END)  # 获取上面文本框中的全部内容
        logging.info("stringValueXX: %s", stringValue)
        indexList = [substr.start() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]       # 段落编号首个数字所在位置(有全角和半角的英文句号，所以使用(\.|\．|\、| ))
        paraNumList = [substr.group().strip() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]    # 段落编号（数字+.或者数字+．或者数字+、或者数字+空格）

        logging.info("paraNumList: %s", paraNumList)   
        logging.info("indexList: %s", indexList)

        ### 正文内容中出现过这样的内容("2.34 - 5的常规变量")， finditer会将正文中的'2.'当成段落编号，所以需要进行处理
        sPos = 1.0
        allPostion = []
        for paraNum in paraNumList:
            pos = text1.search(paraNum,sPos,END)
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

        numList1 = []        # 存放对应段落中的引用(引用关系只考虑每个段落中“其特征在于”关键字或者第一个逗号之前的字符串)
        numList2 = []       # 存放对应段落中的引用（引用关系考虑的是“权利要求”关键字和“所述”关键字之间的段落）
        targetStrList = []  # 保存所有的权利要求和逗号之间的字符串
        paragList = []      # 保存所有的段落
        for i in range(len(indexList)):
            if i == len(indexList) - 1:
                parag = stringValue[indexList[i]:]
            else:
                parag = stringValue[indexList[i]:indexList[i+1]]
            logging.info("Para Text: %s",parag)
            paragList.append(parag)

            # 引用关系中的权利要求只关注每个段落第一个逗号或者“其特征在于”关键字之前的段落
            targetIndex = 0  # 初始化
            qtzzyIndex = str(parag).find("其特征在于")  # 找到第一个“其特征在于”字符所在位置
            if qtzzyIndex == -1:
                douHaoIndex = str(parag).find("，")  # 如果没找到“其特征在于”，就找第一个逗号
                if douHaoIndex != -1:
                    targetIndex = douHaoIndex
            else:
                targetIndex = qtzzyIndex
            # 段落编号和“其特征在于”或者“逗号”之间的字符串
            targetStr1 = stringValue[indexList[i]:indexList[i] + targetIndex]
            logging.info("targetStr1: %s", targetStr1)  
            resList1 = getNum(targetStr1)
            logging.info("resList1: %s", resList1)
            numList1.append(resList1)

            # 判断引用权利是否存在择一的错误要考虑的是“权利要求”关键字和权利要求后的第一个标点符号之间的段落，“权利要求”关键字有可能存在自然段落中间位置。
            # “权利要求”字符所在的全部位置列表
            qlyqindexList = [substr.start() for substr in re.finditer(
                            "权利要求", str(parag))]  
            logging.info("qlyqindexList: %s",qlyqindexList)
            # “权利要求”字符后的第一个逗号所在index
            dhList2 = []   
            for index in qlyqindexList:
                for i in range(index,len(parag)):
                    if parag[i] == '，' or parag[i] == ',' or parag[i] == '；' or parag[i] == ';' or parag[i] == '：' or parag[i] == ':' or parag[i] == '、' or parag[i] == '。':
                        dhList2.append(i)
                        break
            interList = []
            for i in range(len(qlyqindexList)):
                # 权利要求和逗号之间的字符串
                targetStr = str(parag)[
                    qlyqindexList[i]+4:dhList2[i]]
                logging.info("targetStr: %s", targetStr)
                targetStrList.append(targetStr)
                resList = getNum2(targetStr)
                logging.info("interList: %s", resList)
                # 把同一个段落编号中所有的引用数值都放在一个列表中
                interList.extend(resList)
            interList = list(set(interList))  # 去掉重复的引用数值
            numList2.append(interList)

        logging.info("getRefList paracodeList: %s",paracodeList)  # 段落编号
        logging.info("getRefList numList1: %s", numList1)           # 对应段落编号中引用的权利要求书(只考虑“其特征在于”关键字之前或者第一个逗号之前的字符串)
        logging.info("getRefList numList2: %s", numList2)           # 对应段落编号中引用的权利要求书
        logging.info("getRefList targetStrList: %s", targetStrList) 
        logging.info("getRefList paragList: %s", paragList)

        return paracodeList, numList1, numList2, targetStrList, paragList

    def getRefJson():
        # global resJson
        paracodeList, numList1, numList2, targetStrList, paragList = getRefList()
        logging.info("paracodeList:%s",paracodeList)
        logging.info("numList1: %s", numList1)
        logging.info("numList2: %s", numList2)
        res1 = getFormat(paracodeList, numList1)
        resJson1 = json.loads(res1)
        # 把引用数值排序一下
        for key, value in resJson1.items():
            resJson1[key] = sorted(value)
        logging.info("Key-Value1: %s", resJson1)

        res2 = getFormat(paracodeList, numList2)
        resJson2 = json.loads(res2)
        # 把引用数值排序一下
        for key, value in resJson2.items():
            resJson2[key] = sorted(value)
        logging.info("Key-Value2: %s", resJson2)

        return resJson1, resJson2


    # 判断敏感词语
    def analysisWarn():
        
        hasWarn = False  # 判断是否出现敏感词
        # 获取所有的敏感词
        jsonInfo = readJsonFile()
        if jsonInfo != "":
            MGCList = jsonInfo['MGC']
            logging.info("json data: %s", MGCList)
            for key in MGCList:
                start = 1.0
                while True:
                    pos = text1.search(key,start,END)
                    logging.info("pos: %s" , pos)
                    if pos == '':
                        break
                    else:
                        hasWarn = True
                        start = '%s+%dc' % (pos, len(key))
                        text1.delete(pos, start)
                        text1.insert(pos, key, 'tag')
                        ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                        text1.tag_add('tag',pos) #申明一个tag,在a位置使用
                        text1.tag_config('tag',background='yellow') #设置tag即插入文字的大小,颜色等
        
        return hasWarn

    # 检查主题名称不一致
    def nameDiff():

        resJson, resJson2 = getRefJson()  # 获取权利要求之间的引用关系(eg:{'1': [], '2': [1], '3': [1, 2], '4': [1, 2], '5': [1, 2, 4], '6': [1, 2, 4, 5]})

        singleList = []  # 独权的段落号码
        singleDict = {}  # 独权：[引用独权的权利要求]
        for key, value in resJson.items():
            if len(value) == 0:
                singleList.append(int(key))

        for single in singleList:
            valueList = []
            for key,value in resJson.items():
                if single in value:
                    valueList.append(key)
            singleDict[str(single)] = valueList
        logging.info("singleDict : %s",singleDict)      

        stringValue = text1.get(1.0,END)  # 获取上面文本框中的全部内容
        indexList1 = [substr.start() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]     # 段落编号首个数字所在位置
        paraNumList = [substr.group().strip() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]    # 段落编号（数字+.）

        ### 正文内容中出现过这样的内容("2.34 - 5的常规变量")， finditer会将正文中的'2.'当成段落编号，所以需要进行处理
        sPos = 1.0
        allPostion = []
        for paraNum in paraNumList:
            pos = text1.search(paraNum,sPos,END)
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
            del indexList1[deleteIndex]
            del paraNumList[deleteIndex]

        paracodeList= [re.sub(r"\.|\．|\、| ","",element) for element in paraNumList]
        ##################

        indexList2 = []   # 每段第一个逗号所在的index
        for index in indexList1:
            for i in range(index,len(stringValue)):
                if stringValue[i] == '，' or stringValue[i] == ',':
                    indexList2.append(i)
                    break
        logging.info("nameDiff indexList1: %s", indexList1)
        logging.info("nameDiff indexList2: %s", indexList2)

        strList = []  # 存放段落和逗号之间的字符串
        for i in range(len(indexList1)):
            strList.append(stringValue[indexList1[i]:indexList2[i]])
        logging.info("string list: %s", strList)  

        strDict = {}  
        for i in range(len(paracodeList)):
            strDict[paracodeList[i]] = strList[i]
        logging.info("strDict : %s",strDict)
        
        return singleDict, strDict

    

    # 引用关系错误
    def refError():
        paracodeList, numList, numList2, targetStrList, paragList = getRefList()
        intList = list(map(int,paracodeList))
        logging.info("refError intList: %s", intList)
        maxParaCode = max(intList)   # 列表中最大的段落编码
 
        strValue = ""
        if len(paracodeList) == len(numList2):
            strValue = '{'
            for i in range(len(paracodeList)):
                strValue = strValue + '"' + \
                    str(paracodeList[i]) + '"' + ':' + str(numList2[i])
                if i != len(paracodeList) - 1:
                    strValue = strValue + ','
            strValue = strValue + '}'
        jsonValue = json.loads(strValue)
        logging.info("refError function jsonValue: %s", jsonValue)

        return jsonValue, maxParaCode, targetStrList, numList2
        
        
    # 检查标点符号
    def checkSymbol():

        stringValue = text1.get(1.0,END)  # 获取上面文本框中的全部内容
        logging.info("stringValueXX: %s", stringValue)
        indexList = [substr.start() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]       # 段落编号首个数字所在位置(有全角和半角的英文句号，所以使用(\.|\．|\、))
        paraNumList = [substr.group().strip() for substr in re.finditer(r"\d+(\.|\．|\、| )", stringValue)]    # 段落编号（数字+.或者数字+．）

        logging.info("paraNumList: %s", paraNumList)   
        logging.info("indexList: %s", indexList)

        ### 正文内容中出现过这样的内容("2.34 - 5的常规变量")， finditer会将正文中的'2.'当成段落编号，所以需要进行处理
        sPos = 1.0
        allPostion = []
        for paraNum in paraNumList:
            pos = text1.search(paraNum,sPos,END)
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

        return stringValue, indexList, paracodeList
    
    # 得到此时引用基础所在的权利要求编号
    def getqlCode(posList, paracodeList, posStr):

        n = 0
        i = 0

        RCList = posStr.split(".")
        R = int(RCList[0])     # 行数
        C = int(RCList[1])   # 列数

        for i in range(len(posList)):
            RowColumList1 = posList[i].split(".")
            Row1 = int(RowColumList1[0])     # 行数
           
            if i == len(posList) - 1:
                Row2 = 9999999
                Colum2 = 9999999
            else:
                RowColumList2 = posList[i+1].split(".")
                Row2 = int(RowColumList2[0])     # 行数
                Colum2 = int(RowColumList2[1])   # 列数

            if (R >= Row1 and R < Row2) or (R == Row2 and C < Colum2):
                n = paracodeList[i]
                break
        return n, i


                        
    ###################-----开始执行-------################

    title = "权利要求书质量检查结果"
    message = ""

    # 在文本框1中按鼠标右键弹窗搜索
    def find_keyWord(event):
        
        def getEntertText():
            keyW1 = entryText1.get() # 想要搜索的关键字1
            keyW2 = entryText2.get() # 想要搜索的关键字2
            keyW3 = entryText3.get() # 想要搜索的关键字3
            keyW4 = entryText4.get() # 想要搜索的关键字4

            keyWList = []
            if keyW1 != '':
                keyWList.append(keyW1)
            if keyW2 != '':
                keyWList.append(keyW2)
            if keyW3 != '':
                keyWList.append(keyW3)
            if keyW4 != '':
                keyWList.append(keyW4)    

            return keyWList

        # 高亮显示要搜索的关键字
        def searchKeyWord():
            keyW1 = entryText1.get() # 想要搜索的关键字1
            keyW2 = entryText2.get() # 想要搜索的关键字2
            keyW3 = entryText3.get() # 想要搜索的关键字3
            keyW4 = entryText4.get() # 想要搜索的关键字4

            keyWList = []
            if keyW1 != '':
                keyWList.append(keyW1)
            if keyW2 != '':
                keyWList.append(keyW2)
            if keyW3 != '':
                keyWList.append(keyW3)
            if keyW4 != '':
                keyWList.append(keyW4)

            color =['lightgreen', 'lightblue', 'bisque', 'pink']
            for i in range(len(keyWList)):
                s = 1.0
                while True:
                    p = text1.search(keyWList[i],s,END)
                    if p == '':
                        break
                    else:
                        s = '%s+%dc' % (p, len(keyWList[i]))
                        text1.delete(p, s)
                        tag = 'searchK' + str(i)
                        text1.insert(p, keyWList[i],tag )
                        ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                        text1.tag_add(tag,p) #申明一个tag,在a位置使用
                        text1.tag_config(tag,background=color[i]) #设置tag即插入文字的大小,颜色等
                        
                        
     
        # 清除已有的高亮关键字
        def clearColor():
            keyW1 = entryText1.get() # 想要搜索的关键字1
            keyW2 = entryText2.get() # 想要搜索的关键字2
            keyW3 = entryText3.get() # 想要搜索的关键字3
            keyW4 = entryText4.get() # 想要搜索的关键字4

            keyWList = []
            if keyW1 != '':
                keyWList.append(keyW1)
            if keyW2 != '':
                keyWList.append(keyW2)
            if keyW3 != '':
                keyWList.append(keyW3)
            if keyW4 != '':
                keyWList.append(keyW4)

            for i in range(len(keyWList)):
                s = 1.0
                while True:
                    p = text1.search(keyWList[i],s,END)
                    if p == '':
                        break
                    else:
                        s = '%s+%dc' % (p, len(keyWList[i]))
                        text1.delete(p, s)
                        tag = 'clearcolor' + str(i)
                        text1.insert(p, keyWList[i],tag )
                        ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                        text1.tag_add(tag,p) #申明一个tag,在a位置使用
                        text1.tag_config(tag,background='') #设置tag即插入文字的大小,颜色等
                       
  


        root2 = Tk(className="搜索关键字")
        root2.wm_attributes('-topmost', 1)
        screenwidth, screenheight = root2.maxsize()
        width = 350
        height = 200
        size = '%dx%d+%d+%d' % (width, height, (screenwidth -
                                width) / 4, (screenheight - height) / 4)
        root2.geometry(size)
        root2.resizable(0, 0)



        entryText1 = Entry(root2, bg='lightgreen')
        entryText1.configure(font=("微软雅黑", 10))  # 设置文本框中的字体
        entryText1.pack(padx=(10,10), pady=(10,0), fill='x')

        entryText2 = Entry(root2, bg='lightblue')
        entryText2.configure(font=("微软雅黑", 10))  # 设置文本框中的字体
        entryText2.pack(padx=(10,10), pady=(10,0), fill='x')

        entryText3 = Entry(root2, bg='bisque')
        entryText3.configure(font=("微软雅黑", 10))  # 设置文本框中的字体
        entryText3.pack(padx=(10,10), pady=(10,0), fill='x')

        entryText4 = Entry(root2, bg='pink')
        entryText4.configure(font=("微软雅黑", 10))  # 设置文本框中的字体
        entryText4.pack(padx=(10,10), pady=(10,0), fill='x')

        fra2 = Frame(root2)
        fra2.pack(side=BOTTOM)

        bsearch = Button(fra2, text="搜索文本框中的关键字",
        command=lambda: searchKeyWord())
        bsearch.grid(row=0, column=0, padx=(10, 0),pady=(8,15))

        bclearColor = Button(fra2, text="清除文本框中相关文字的高亮显示",
        command=lambda: clearColor())
        bclearColor.grid(row=0, column=1, padx=(10, 0),pady=(8,15))
        root2.mainloop()


     # 关闭窗口时的执行函数
    def close_callback():
        if tkinter.messagebox.askokcancel('信息提示', '您正在关闭“权利要求书质量检查结果”窗口！！'):
            root.destroy()


    # 设置弹窗尺寸
    root = Tk(className=title)
    root.wm_attributes('-topmost', 0)
    screenwidth, screenheight = root.maxsize()

    width = screenwidth
    height = screenheight
    logging.info("widthXXX : %s", width)
    logging.info("heightXXX : %s", height)
    size = '%dx%d+%d+%d' % (width, height,
                            (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)
    root.resizable(1, 1)  # 设置X和Y轴都可改变


    # 设置BUTTON按钮
    frame = Frame(root)
    frame.pack(side=TOP)

    # 设置输入框
    frame_1 = Frame(root)
    frame_1.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    text1 = Text(frame_1)
    text1.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)

    text1.configure(font=("微软雅黑", 10))   # 设置文本框中的字体
    text1.bind('<Button-3>', find_keyWord)
    text1.focus_set()


    frame_2 = Frame(root)
    frame_2.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    text2 = Text(frame_2)
    text2.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)

    text2.configure(font=("微软雅黑", 12))  # 设置文本框中的字体
    root.protocol("WM_DELETE_WINDOW", close_callback)
    


    ####### 执行质量检查函数 ####

    textString1 = Content
    logging.info("textString1:%s", textString1)
    text1.insert(1.0, textString1, 'AllClear')
    ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
    text1.tag_add('AllClear',1.0) #申明一个tag,在a位置使用
    text1.tag_config('AllClear',background='') #设置tag即插入文字的大小,颜色等
    

    result1 = [0,0,0,0,0]    # 记录总的结果出现错误的次数(主题名称不一致、引用关系错误、标点符号错误, 主题名称前面缺少所述，引用基础错误)
    result2 = []    # 列举每一条错误


    # 检查文章第一段是否缺少权利要求号码“1.”,只关注第一个段落即可
    checkFirstParaNum()

    # 1、检查敏感词
    res1 = analysisWarn()   
    if res1:
        result2.append("权利要求书中出现敏感词，已用黄色背景标记")

    # 2、检查主题名称不一致
    singleDict,  strDict = nameDiff()
    errorNum = 0  # 主题名称错误的次数
    for key, value in singleDict.items():
        str1 = strDict[key]  # 独权截止到第一个逗号的字符串
        index = str1.find("一种")  # 找到 返回第一个字符的下标，找不到 返回 -1
        if index != -1:
            targetStr1 = str1[index + 2:]  # 独权的主题名称
            targetStr2 = str1[-2:]         # 主题名称的最后两个字也算作主题名称
        else:
            targetStr1 = targetStr2 = str1
        logging.info("targetStr1: %s",targetStr1)
        for v in value:
            str2 = strDict[v]  # 从属权利截止到第一个逗号的字符串
            index1 = str2.find("所述")
            index2 = str2.find("所述的")
            isNotSame = False  # 是否出现不一致的情况标志
            if index1 != -1 and index2 == -1:
                testStr1 = str2[index1+2:]
                if not((targetStr1 in testStr1) or (targetStr2 in testStr1) or (testStr1 in targetStr1) or (testStr1 in targetStr2)):
                        isNotSame = True
            if index2 != -1:
                testStr2 = str2[index2+3:]
                if not((targetStr1 in testStr2) or (targetStr2 in testStr2) or (testStr2 in targetStr1) or (testStr2 in targetStr2)):
                        isNotSame = True
            if isNotSame:
                errorNum = errorNum + 1
                res = "权利要求" + v + "引用的主题名称与独立权利要求" + key + "的主题名称不一致。"
                result2.append(res)               
    result1[0] = errorNum           

    logging.info("nameDiff result1: %s", result1)
    logging.info("nameDiff result2: %s", result2)


    # 3、检查引用关系错误
    res3, maxNum, strList, allNumList = refError()
    logging.info("res3: %s",res3)
    logging.info("maxNum: %s",maxNum)
    logging.info("strList: %s",strList)
    logging.info("allNumList: %s",allNumList)

    # 对strList进行填充，保持与numList2的数量是一致的，以修改择一功能的bug
    noqlyqList = []  #在numList2中元素是空的序列组成的列表
    newstrList = [] # 填充后的新列表
    for m in range(len(allNumList)):
        if len(allNumList[m]) == 0:
            noqlyqList.append(m)
    logging.info("noqlyqList: %s",noqlyqList)
    n = 0
    m = 0
    while n < len(allNumList):
        if n not in noqlyqList and m < len(strList):
            newstrList.append(strList[m])
            m = m + 1
        if n in noqlyqList:
            newstrList.append("XX")  # 填充字符串
        n = n + 1
    logging.info("newstrList: %s",newstrList)


    errorNum = 0
    duoInduo = []   # 保存多引多的段落编码
    paraStrDict = {}  #存放段落编码和字符串的（eg: {'1': "xxxx", '2':'xxxxx'})
    i = 0
    for key, value in res3.items():
        for v in value:
            r = ""
            if int(key) == v:
                errorNum = errorNum + 1
                r = "权利" + key + "自己引用了自己." 
                result2.append(r)
            if int(key) < v:
                errorNum = errorNum + 1
                if v > maxNum:
                    r = "权利" + key + "引用了不存在的权利要求。" 
                else:
                    r = "权利" + key + "引用了在后权利要求" + str(v) + '。'
                result2.append(r)

        r = ""
        if len(value) >= 2:
            for v1 in duoInduo:
                if v1 in value:
                    errorNum = errorNum + 1
                    r = "权利" + key + "出现了多引多。" 
                    if r not in result2:
                        result2.append(r)
            duoInduo.append(int(key))
        logging.info("duoInduo: %s", duoInduo)
        # 检查择一引用
        if len(value) != 0:
            paraStrDict[key] = newstrList[i]
        i = i + 1
    logging.info("paraStrDict: %s", paraStrDict)

    for key, strValue in paraStrDict.items():
            index = strValue.find("所述")
            if index != -1:
                strValue = strValue[:index]
            indexList = [substr.start() for substr in re.finditer(r"\d", strValue)]       # 段落编号首个数字所在位置
            strList = [substr.group() for substr in re.finditer(r"\d", strValue)] 
            logging.info("refError indexList1: %s", indexList)
            new_indexList = copy.deepcopy(indexList)  # 深拷贝，不破坏indexList的结构
            # 处理“权利要求10所述XXXX”这种语句，这种情况下indexList会拿到1和0的下标，但是10属于一个数字，所以就不能要1的下标了
            for i in range(1, len(new_indexList)):
                if new_indexList[i] - new_indexList[i-1] == 1:
                    indexList.remove(new_indexList[i])

            logging.info("refError indexList2: %s", indexList)  
            logging.info("refError strList: %s", strList)

            hasDanger = True
            # 每个段落中有超过两个权利要求编号的才进行判断
            if len(indexList) >= 2:
                for i in range(len(indexList)):
                    if i == len(indexList) - 1:
                        break
                    strV = strValue[indexList[i] + 1:indexList[i + 1]]
                    logging.info("refError strV: %s", strV)
                    dangerWord = ["至","和","-","到","~", "、"]
                    word = ["任意一", "任一","之一","任何一","中的一项","中的一个","中一项","中一个"]
                    for w1 in dangerWord:
                        if strV.find(w1) != -1:
                            for w2 in word:
                                if strValue.find(w2) != -1:
                                    hasDanger = False
                                    break                
                        else:
                            if strValue.find("或者") != -1 or strValue.find("或") != -1:
                                hasDanger = False

                if hasDanger:
                    errorNum = errorNum + 1
                    r = "权利要求" + key + "没有以择一引用的方式引用权利要求"            
                    result2.append(r)

    result1[1] = errorNum
    logging.info("refError result1: %s", result1)
    logging.info("refError result2: %s", result2)


    # 4、检查标点符号
    stringValue, indexList, paracodeList = checkSymbol()
    errorNum = 0
    end = 1.0
    for i in range(len(indexList)):
        if i == len(indexList) - 1:
            parag = stringValue[indexList[i]:]
        else:
            parag = stringValue[indexList[i]:indexList[i+1]]
        logging.info("checkSymbol Para Text: %s",parag)

        paragText = parag
        logging.info("paragText : %s",paragText)

        # 每一个段落的起始和结束序列
        start = end
        end = '%s+%dc' % (start, len(paragText))
        logging.info("start: %s", start)
        logging.info("end : %s", end)
        logging.info("start,end : %s", text1.get(start,end))

        # 检查重复的标点符号
        symbolIndex = [substr.start() for substr in re.finditer(r"[，。、：！,.!&；;]{2,}", paragText)]   
        symbolCode = [substr.group() for substr in re.finditer(r"[，。、：！,.!&；;]{2,}", paragText)]
        if len(symbolIndex) != 0:
            errorNum = errorNum + 1
            r = ""
            r = "权利要求" + paracodeList[i] + "中存在重复相连的标点符号，已经用粉色背景标记" 
            result2.append(r)

        # 检查段落中句号多余问题
        symbolList = [substr.start() for substr in re.finditer(r"。", paragText)]  # 本段落中所有的句号序列
        logging.info("symbolList: %s",symbolList)

        r = ""
        if len(symbolList) == 0:
            errorNum = errorNum + 1
            r = "权利要求" + paracodeList[i] + "中没有句号。"
            result2.append(r)
        logging.info("len(paragText) - 1 : %s", len(paragText.strip()) - 1)
        if (len(symbolList) == 1) and (symbolList[0] != len(paragText.strip()) - 1):
            errorNum = errorNum + 1
            r = "权利要求" + paracodeList[i] + "中只有一个句号，但不在句尾，已用粉色背景标记。"
            result2.append(r)

            # 添加上背景颜色
            pos = text1.search("。",start,end)
            pos2 = '%s+%dc' % (pos, 1)
            text1.delete(pos,pos2)
            text1.insert(pos, "。", 'notEnd')
            ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
            text1.tag_add('notEnd',pos) #申明一个tag,在a位置使用
            text1.tag_config('notEnd',background='fuchsia') #设置tag即插入文字的大小,颜色等


        if len(symbolList) > 1:
            errorNum = errorNum + 1
            r = "权利要求" + paracodeList[i] + "中有多个句号，多余句号已用粉色背景标记。"
            result2.append(r)

            # 添加上背景颜色
            s = start
            logging.info("s: %s",s)
            logging.info("end: %s",end)
            while True:
                pos = text1.search("。",s,end)
                logging.info("symbol pos: %s" , pos)
                if pos == '':
                    break
                else:  
                    pos2 = '%s+%dc' % (pos, 1)
                    text1.delete(pos,pos2)
                    text1.insert(pos, "。", 'moreThanOne')
                    ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                    text1.tag_add('moreThanOne',pos) #申明一个tag,在a位置使用
                    text1.tag_config('moreThanOne',background='fuchsia') #设置tag即插入文字的大小,颜色等
                    s = pos2
    
    # 为重复的标点符号添加背景颜色
    symbolIndex = [substr.start() for substr in re.finditer(r"[，。、：！,.!&；;]{2,}", stringValue)]   
    symbolCode = [substr.group() for substr in re.finditer(r"[，。、：！,.!&；;]{2,}", stringValue)]
    logging.info("symbolIndex: %s", symbolIndex)
    logging.info("symbolCode: %s", symbolCode)
    s = 1.0
    for i in range(len(symbolIndex)):
        pos = '%s+%dc' % (s, symbolIndex[i])
        pos2 = '%s+%dc' % (pos, len(symbolCode[i]))
        text1.delete(pos,pos2)
        text1.insert(pos, symbolCode[i], 'dup')
        ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
        text1.tag_add('dup',pos) #申明一个tag,在a位置使用
        text1.tag_config('dup',background='fuchsia') #设置tag即插入文字的大小,颜色等
    
    
    result1[2] = errorNum

    # 5. 发明名称超过25个字
    paraDict, paraTextDict = nameDiff()
    logging.info("paraDict: %s", paraDict)
    logging.info("paraTextDict: %s", paraTextDict)
    for key1, value1 in paraDict.items():
        paraText = paraTextDict[key1]
        logging.info("paraText: %s", paraText)
        startIndex = paraText.find("一种")
        if startIndex != -1:
            string = paraText[startIndex+1:]
            print(string)
            if len(string) > 25:
                r = "权利" + key1 + "中的发明名称长度大于25，请检查！"
                result2.append(r)

    # 6. 检查缺少"所述"的问题
    paraDict, paraTextDict = nameDiff()
    errorNum = 0
    for key1, value1 in paraDict.items():
        logging.info("value1: %s", value1)
        for v in value1:
            string = paraTextDict[v]
            if string.find("所述") == -1:
                errorNum = errorNum + 1
                r = "权利" + v + "中的主题名称缺少‘所述’的引用词语！"
                if r not in result2:
                    result2.append(r)
    result1[3] = errorNum

    # 7.缺少引用基础

    # 获取所有的引用基础词语
    flagStrList = []   # 保存文本框2中需要被红色字体标记的关键字
    jsonInfo = readJsonFile()
    if jsonInfo != "":
        YYJCList = jsonInfo['YYJC']
        logging.info("json data: %s", YYJCList)
        keyValue1, keyValue2 = getRefJson()
        logging.info("keyValue2: %s", keyValue2)

        paraCodeStrDict = {}
        paracodeList, numList, numList2, targetStrList, paragList = getRefList()
        for i in range(len(paracodeList)):
            paraCodeStrDict[paracodeList[i]] = paragList[i]
        logging.info("paraCodeStrDict: %s", paraCodeStrDict)  # 权利要求编号： 对应的段落字符串


        # 获取每个段落的位置
        allNumPos = []
        for paracode in paracodeList:
            symCode = ['.','．','、',' ']
            for sym in symCode:
                numStr = paracode + sym
                pos = text1.search(numStr, 1.0,END)
                if pos != '':
                    p = pos.split('.')
                    if p[1] == '0':
                        allNumPos.append(pos)
                    break
        logging.info("allNumPos: %s", allNumPos)


        # 获取引用基础词语和它所在的行列位置
        rowColumnAndWordDict = {}  # 所有的引用基础词语所在行列位置信息和对应的引用基础词语(eg、2.13 : 所述)
        for yyjc in YYJCList:
            startCheck = 1.0
            while True:
                startPos = text1.search(yyjc,startCheck,END)
                if startPos == '':
                    break
                else:
                    rowColumnAndWordDict[startPos] = yyjc
                startCheck = '%s+%dc' % (startPos, len(yyjc))
        logging.info("rowColumnAndWordDict: %s", rowColumnAndWordDict)

        errorNum = 0
        deleteList = []  # 不需要在文本框2进进行标红的截词
        for k, v in rowColumnAndWordDict.items():
            qlyqNum, posListIndex = getqlCode(allNumPos, paracodeList, k)  # qlyqNum: 当前引用基础词语所在的权利要求编号  ,posListIndex :当前引用基础词语所在段落开头编号的行列位置在allNumPos列表中的位置
            beqlyqNumList = keyValue2[qlyqNum]      # beqlyqNumList : qlyqNum所引用的权利要求列表
            
            # 被引用的段落中所有字符串
            string = ""
            for num in beqlyqNumList:
                string = string + paraCodeStrDict[str(num)]
            # 再加上引用基础词语所在的段落开头至引用基础词语之间的字符串
            string = string + text1.get(allNumPos[posListIndex], k)
            logging.info("string: %s", string)

            p1 = '%s+%dc' % (k, len(v))
            p2 = '%s+%dc' % (p1, int(wordLen))
            keyStr = text1.get(p1, p2)
            logging.info("keyStr: %s", keyStr)
            # 带有标点符号，不能算作是所述后面的关键字，所以应该剔除
            isflag1 = False  # 带有标点符号的关键字不在文本框2中进行标红显示
            isflag2 = False
            
            symbList = [substr.group() for substr in re.finditer(r"(\&|\“|\”|\‘|\’|\《|\》|\{|\}|\]|\（|\）|\[)", keyStr)]   
            symbLen = len(symbList)
            while symbLen != 0:
                isflag1 = True
                for symb in symbList:
                    keyStr = keyStr.replace(symb,"")

                p3 = '%s+%dc' % (p2, symbLen)
                keyStr = keyStr + text1.get(p2, p3)
                symbList = [substr.group() for substr in re.finditer(r"(\&|\“|\”|\‘|\’|\《|\》|\{|\}|\]|\（|\）|\[)", keyStr)]   
                symbLen = len(symbList)
                p2 = p3

            logging.info("new_keyStr: %s", keyStr)


            if string.find(keyStr) == -1:
                qlyqNum, n = getqlCode(allNumPos, paracodeList, k)

                if n == len(allNumPos) - 1:
                    endPos = END
                else:
                    endPos = allNumPos[n + 1]

                douHaoPos = text1.search("，", p1, endPos)
                if douHaoPos == "":
                    douHaoPos = endPos
                partStr = text1.get(p1, douHaoPos)
                isflag2 = True
                r = "权利要求" + qlyqNum + "中" + "/*" + partStr.replace("\n","") + "*/" + "有缺乏引用基础的表述。"
                errorNum = errorNum + 1
                logging.info("partStr : %s",partStr)
                # 处理特殊情况，（例如 wordLen设置为4， KeyStr = "电机上，" 因为以逗号作为结束，所有，partStr = "电机上"， 以字符串长度短的为准 ）
                flagStr = partStr if len(keyStr) >= len(partStr) else keyStr
                flagStrList.append(flagStr.replace('\n',''))  
                if isflag1 and isflag2:
                    deleteList.append(len(flagStrList) - 1)

                text1.delete(k, p1)
                text1.insert(k, v, 'yyjccy')
                ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                text1.tag_add('yyjccy',k) #申明一个tag,在a位置使用
                text1.tag_config('yyjccy',background='red') #设置tag即插入文字的大小,颜色等
                result2.append(r)
            
            isflag1 = False
            isflag2 = False
        logging.info("deleteList1 : %s",deleteList)
        result1[4] = errorNum   

    # 8、判断附图标记说明的数字是否正确
    originStrList = [] # 附图标记的名字列表，不带数字的 
    numList = []       # 附图标记的数字列表

    for text in formatftbjList:
        wordList = re.split(r'[:、.．： \t]{1,}', text) # text包含此类情况（33a：纸卷）,为把33a识别出来，所以需要提前划分一下
        if len(wordList) == 2:
            if wordList[0].encode('UTF-8').isalnum():
                num = wordList[0]
                string = wordList[1]
            else:
                indexList = [substr.start() for substr in re.finditer(r"~|-|\(|\（", wordList[0])]    # 特殊情况，例如 101-106：纸卷 或者 101~106：纸卷或者11（12）：纸卷
                if len(indexList) == 0:
                    num = wordList[1]
                    string = wordList[0]
                else:
                    num = wordList[0]
                    string = wordList[1]
        else:
            num = ''.join([i for i in text if i.isdigit() or i.encode( 'UTF-8' ).isalpha()])
            string = ''.join([i for i in text if i.isalpha()])
        if string != "":
            originStrList.append(string)
        numStr = "（" + str(num) + "）"
        numList.append(numStr)
    logging.info("originStrList: %s", originStrList)
    logging.info("numList: %s", numList)

    # 开始判断附图标记的错误
    testList = []
    ishasError = FALSE
    for i in range(len(originStrList)):
        logging.info("originStrList[i]: %s", originStrList[i])
        startPos = 1.0
        while True:
            pos = text1.search(originStrList[i],startPos,END)
            if pos == "":
                break
            else:
                startPos = '%s+%dc' % (pos, len(originStrList[i]))
                # 获取originStrList关键字后面的字符
                compareNumStartPos = startPos 
                compareNumEndPos = '%s+%dc' % (compareNumStartPos, len(numList[i]))
                keyStr = text1.get(compareNumStartPos,compareNumEndPos)
                logging.info("keyStr: %s", keyStr)
                logging.info("numList: %s", numList[i])
                if numList[i] != keyStr: 
                    logging.info("originStrList[i]XXX: %s", originStrList[i])
                    testList.append(originStrList[i])
                    ishasError = True
                    #给错误的地方添加蓝色背景颜色
                    text1.delete(pos,startPos)
                    text1.insert(pos,originStrList[i], "notEqual")
                    ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                    text1.tag_add('notEqual',pos) #申明一个tag,在a位置使用
                    text1.tag_config('notEqual',background='dodgerblue') #设置tag即插入文字的大小,颜色等

    if ishasError:
        r = "权利要求书中出现附图标记数字不正确的情况，已经标记为蓝色背景，请做进一步检查"
        result2.append(r)




    #####最终全部检测结果###########
    logging.info("result1: %s", result1)
    logging.info("result2: %s", result2)

    # 最终的全部检查结果
    allResult = '检查报告：' + " "
    
    # 总体结果
    s1 = "主题名称不一致错误:" + str(result1[0]) + "、" + "引用关系错误:" + str(result1[1]) + "、" + "标点错误:" + str(result1[2])+ "、" + "主题名称前面缺少所述两个字错误:" + str(result1[3]) + "、"+ "引用基础错误：" + str(result1[4]) +"\n"
    s1= s1 + "\n"
    
    # 详细结果
    s2 = ""
    for i in range(len(result2)):
        s2 = s2 + str(i + 1) + '、' + str(result2[i]) + "\n"
        s2 = s2 + "\n"
    
    # 写入到文本框2中
    allResult = allResult + s1 + s2
    logging.info("text2: %s",allResult)
    text2.insert(1.0,allResult)

    # 标记文本框2中的引用基础关键字
    startP = []
    sP = 1.0
    while True:
        p = text2.search("/*",sP, END)
        if p == "":
            break
        else:
            sP = '%s+%dc' % (p, len("/*"))
        startP.append(p)
    logging.info("startP: %s", startP)
    logging.info("flagStrList: %s", flagStrList)

    logging.info("deleteList2: %s", deleteList)
    for deleteIndex in deleteList:
        del startP[deleteIndex]
        del flagStrList[deleteIndex]

    logging.info("new startP: %s", startP)
    logging.info("new flagStrList: %s", flagStrList)

    for i in range(len(startP)):
        p1 = '%s+%dc' % (startP[i], len("/*"))
        p2 = '%s+%dc' % (p1, len(flagStrList[i]))
        text2.delete(p1, p2)
        text2.insert(p1, flagStrList[i], 'fStr')
        ft = tf.Font(family = '微软雅黑', size=12, underline=True) ###有很多参数
        text2.tag_add('fStr',p1) #申明一个tag,在a位置使用
        text2.tag_config('fStr',foreground='red') #设置tag即插入文字的大小,颜色等


    ######最终全部检测结果###############        
    
    root.mainloop()
    

    # return resJson
    
   




# #主入口函数
# if __name__ == '__main__':
#     checkQLYQ()

