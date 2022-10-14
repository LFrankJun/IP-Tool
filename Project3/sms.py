
"""
Time:    2022/10/11
Author:  Li YuanJun
Version: V3.0
"""


import json
import re
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


####### ======主函数入口===========######

def main_sms(Content, ftNum, zhaiyaoName,formatftbjList):

    # 关闭窗口时的执行函数
    def close_callback():
        if tkinter.messagebox.askokcancel('信息提示', '您正在关闭“说明书质量检查结果”窗口！！'):
            root.destroy()


    # 设置弹窗尺寸
    root = Tk(className="说明书质量检查结果")
    root.wm_attributes('-topmost', 0)
    screenwidth, screenheight = root.maxsize()
    width = screenwidth
    height = screenheight
    size = '%dx%d+%d+%d' % (width, height,
                            (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)
    root.resizable(1, 1)  # 设置X和Y轴都可改变


    frame = Frame(root)
    frame.pack(side = TOP)

    # 文本框1
    frame1 = Frame(root)
    frame1.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    # HORIZONTAL 设置水平方向的滚动条，默认是竖直的
    s1= Scrollbar(frame1, orient=HORIZONTAL)
    s1.pack(side= BOTTOM, fill=X)

    s2 = Scrollbar(frame1)
    s2.pack(side=RIGHT, fill=Y)

    text1 = Text(frame1, xscrollcommand=s1.set, yscrollcommand=s2.set, wrap='none')   # wrap设置为不换行
    text1.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    text1.configure(font=("微软雅黑", 10))   # 设置文本框中的字体
    text1.focus_set()


    s1.config(command=text1.xview)
    s2.config(command=text1.yview)

    # 文本框2
    frame2 = Frame(root)
    frame2.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    s3 = Scrollbar(frame2)

    s3.pack(side=RIGHT, fill=Y)
    text2 = Text(frame2,yscrollcommand=s3.set)
    text2.pack(padx=5, pady=5, fill=BOTH, expand=TRUE)
    text2.configure(font=("微软雅黑", 12))  # 设置文本框中的字体
    s3.config(command=text2.yview)

    root.protocol("WM_DELETE_WINDOW", close_callback)




    ###############  开始执行函数  ###############
    textString1 = Content
    text1.insert(1.0, textString1, 'AllClear')
    ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
    text1.tag_add('AllClear',1.0) #申明一个tag,在a位置使用
    text1.tag_config('AllClear',background='') #设置tag即插入文字的大小,颜色等


    result1 = [0,0,0]  # 记录总的结果出现错误的次数(标点符号错误, 附图数错误，小标题错误)
    result2 = []         # 列举每一条错误
    
    
    allString = text1.get(1.0,END)
    if allString != "":
        # 第二次按”检查基本错误“按钮时候，无须再添加行号了（因为在首次按此按钮的时候，已经添加过了）
        if allString[0:2] != "1)":
            # 添加行号
            allString = "1)" + allString
            logging.info("All Text: %s",allString)

            symbolIndexList = [substr.start() for substr in re.finditer(r"\n", allString)] 
            allStringList = list(allString)
            logging.info("symbolIndexList : %s", symbolIndexList)
            del (symbolIndexList[-1])
        
            l = len(symbolIndexList)
            logging.info("l : %s", l)
            for i in range(l-2,-1,-1):
                    symbol = str(i+2) + ')' 
                    allStringList.insert(symbolIndexList[i]+1,symbol)
            newAllString = ''.join(allStringList)
            logging.info(newAllString)

            text1.delete(1.0,END)      # 先删除文本框中原先的内容
            text1.insert(1.0,newAllString)  # 添加到文本框中

        newAllString = text1.get(1.0,END)  # 加上行号之后新的文本框内容


        # 1、对比说明书中的名称和摘要中的名称是否一致
        logging.info("zhaiyaoName %s", zhaiyaoName)
        if zhaiyaoName != "":
            # 获取说明书中的主题名称
            firstIndex = newAllString.find("\n")
            if firstIndex != -1 and len(newAllString) > 2:
                titleName = newAllString[2:firstIndex]
                logging.info("titleName： %s", titleName)
                if titleName != '' and zhaiyaoName.strip() != titleName.strip():
                    result2.append("说明书中的题目与摘要中的主题名称不一致")


        # 2、查看标点符号错误(重复标点，段落末尾没有句号)
        errorNum = 0
        text_content = (text1.get("0.0","end").replace(" ","")).split("\n")   # 每一行内容放到列表中
        logging.info("checkSymbol allString: %s", text_content)

        notJuHao = "" # 存放所有末尾不是句号的段落编号

        for string in text_content:
            if len(string) != 0:
                # 检查重复的标点符号
                logging.info("string: %s",string)
                dumpsymbolIndexList = [substr.start() for substr in re.finditer(r"[，。、：！,.!&；;？?]{2,}", string)]
                dumpsymbolCodeList = [substr.group() for substr in re.finditer(r"[，。、：！,.!&；;？?]{2,}", string)]
                logging.info("dumpsymbolIndexList: %s", dumpsymbolIndexList)
                logging.info("dumpsymbolCodeList: %s", dumpsymbolCodeList)
                if len(dumpsymbolIndexList) != 0:
                    errorNum = errorNum + 1
                    index = string.find(")")
                    if index != -1:
                        r = "自动段落" + string[:index] + "中存在重复相连的标点符号，已经用粉色背景标记" 
                        result2.append(r)

                        # 给重复的标点符号添加粉色背景颜色
                        logging.info("newAllString: %s",newAllString)
                        dumpsymbolIndexList = [substr.start() for substr in re.finditer(r"[，。、：！,.!&；;？?]{2,}", newAllString)]
                        dumpsymbolCodeList = [substr.group() for substr in re.finditer(r"[，。、：！,.!&；;？?]{2,}", newAllString)]
                        s = 1.0
                        for i in range(len(dumpsymbolIndexList)):
                            pos = '%s+%dc' % (s, dumpsymbolIndexList[i])
                            pos2 = '%s+%dc' % (pos, len(dumpsymbolCodeList[i]))
                            text1.delete(pos,pos2)
                            text1.insert(pos, dumpsymbolCodeList[i], 'dup')
                            ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                            text1.tag_add('dup',pos) #申明一个tag,在a位置使用
                            text1.tag_config('dup',background='fuchsia') #设置tag即插入文字的大小,颜色等



                ### 检查段落中末尾没有句号的错误

                # 获取小标题
                jsonInfo = readJsonFile()
                if jsonInfo != "":
                    SUBTITLEList = jsonInfo['SUBTITLE']
                    newSubTitleList = []
                    for title in SUBTITLEList:
                        if title.find('/') != -1:
                            titleList = title.split('/')
                            for t in titleList:
                                newSubTitleList.append(t)
                        else:
                            newSubTitleList.append(title)
                    logging.info("newSubTitleList: %s", newSubTitleList)


                index = string.find(')')
                if index != -1: 
                    if len(string) > index + 1:  # 去掉只有行号的那些行
                        string1 = string[index + 1:] 
                        logging.info("string1: %s", string1)
                        if string1.strip() not in newSubTitleList and string[:2] != '1)':  # 小标题和第一个题目不进行末尾句号的判断
                            lastsymbol = string1[-1]  # 本行最后一个字符
                            logging.info("lastsymbol: %s", lastsymbol)
                            if lastsymbol != '。':
                                errorNum = errorNum + 1
                                notJuHao = notJuHao + string[:index] + "、"
                                
        if notJuHao[:-1] != "":                
            r = "自动段落" + notJuHao[:-1] + "末尾不是句号" 
            result2.append(r)
        result1[0] = errorNum

        # 3、敏感词判断

        # 获取所有的敏感词
        hasWarn = False  # 判断是否出现敏感词
        jsonInfo = readJsonFile()
        if jsonInfo != "":
            MGCList = jsonInfo['MGC']
            logging.info("MGC json data: %s", MGCList)
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
        if hasWarn:
            result2.append('说明书中出现敏感词，已用黄色背景标记')
        
        # 4、附图数错误
        errorNum = 0
        logging.info("ftNum: %s", ftNum)
        tuIndexList = [substr.start() for substr in re.finditer("图", newAllString)]
        numStrList = []
        for index in tuIndexList:
            i = 1
            numStr = ""
            while index + i < len(newAllString) and newAllString[index + i].isdigit():   # 避免出现超过全部内容的长度
                numStr = numStr + newAllString[index+i] # 获取“图”后面的数字，之所以加循环，是因为数字可能是两位数
                i = i + 1
            if numStr != "":
                numStrList.append(numStr)
        logging.info("numStrList: %s", numStrList) 

        numList = list(map(int,numStrList))  # 字符串元素转整型元素
        newNumList = list(set(numList))  # 去掉重复的数字
        logging.info("newNumList: %s", newNumList)

        # 判断说明书中多了图X
        for num in newNumList:
            if num > ftNum:
                errorNum = errorNum + 1
                r = "说明书中多了" + "“" + "图" + str(num) + "”" + "," + "导致说明书和附图不一致"
                result2.append(r)

        #判断说明书中少了图X
        for i in range(1,ftNum + 1):
            if i not in newNumList:
                errorNum = errorNum + 1
                r = "说明书中少了" + "“" + "图" + str(i) + "”"  + "," + "导致说明书和附图不一致"
                result2.append(r)

        result1[1] = errorNum


        # 5、小标题错误 （缺少小标题，小标题顺序不对）
        # 获取小标题
        erroNum = 0
        jsonInfo = readJsonFile()
        if jsonInfo != "":
            SUBTITLEList = jsonInfo['SUBTITLE']  # 直接从json文件中读取的小标题列表
            logging.info("SUBTITLEList: %s", SUBTITLEList)
            newSubTitleList = []   # 去除“/”之后的小标题列表
            for title in SUBTITLEList:
                if title.find('/') != -1:
                    titleList = title.split('/')
                    for t in titleList:
                        newSubTitleList.append(t)
                else:
                    newSubTitleList.append(title)
            logging.info("newSubTitleList: %s", newSubTitleList)

            textTitleList = []  # 说明书中的小标题
            for rowStr in text_content:
                newRowStr = rowStr.strip()
                # logging.info("newRowStr: %s", newRowStr)
                index = newRowStr.find(")")
                if newRowStr != "" and newRowStr[index + 1:] in newSubTitleList:
                    textTitleList.append(newRowStr[index + 1:])
            logging.info("textTitleList: %s", textTitleList)

            titleNameList = []
            # 判断是否缺少某个小标题
            for subTitle in SUBTITLEList:
                sTList = subTitle.split("/")
                logging.info("sTList: %s", sTList)
                titleNameList.append(sTList)
                k = 0 # 计算当前文件中是否缺少json文件列表中的全部元素（为了应对“/”的情况）
                for sT in sTList:
                    if sT not in textTitleList:
                        k = k + 1
                if k == len(sTList):
                    stringName = sT
                    if len(sTList) >=2:
                        stringName = ""
                        for sT in sTList:
                            stringName = stringName + sT + "或"
                        stringName = stringName[:-1]
        
                    r = "缺少小标题" + "“" + stringName + "”"
                    erroNum = erroNum + 1
                    result2.append(r)

            # 判断小标题的顺序是不是正确
            logging.info("titleNameList: %s", titleNameList)
            isorderError = False
            for i in range(len(textTitleList)):
                if textTitleList[i] not in titleNameList[i]:
                    isorderError = True
                    break
            if isorderError:
                r = "小标题顺序不正确，请做进一步检查"
                erroNum = erroNum + 1
                result2.append(r)

        result1[2] = erroNum


        # 6、判断附图标记说明的标记是否准确  
        originStrList = [] # 附图标记的名字列表，不带数字的 
        numList = []       # 附图标记的数字列表
        for text in formatftbjList:
            wordList = re.split(r'[:、.．： \t]{1,}', text) # text包含此类情况（33a：纸卷）,为把33a识别出来，所以需要提前划分一下
            if len(wordList) == 2:
                if wordList[0].isalnum():
                    num = wordList[0]
                    string = wordList[1]
                else:
                    indexList = [substr.start() for substr in re.finditer(r"~|-", wordList[0])]    # 特殊情况，例如 101-106：纸卷 或者 101~106：纸卷
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
            numList.append(num)
        logging.info("originStrList: %s", originStrList)
        logging.info("numList: %s", numList)
        # logging.info("replaceList1: %s", replaceList1)
        
        jtssfsfirstPos = text1.search("具体实施方式",1.0,END)  # 获取“具体实施方式”第一次出现的行列位置
        ishasError = FALSE
        for i in range(len(originStrList)):
            startPos = jtssfsfirstPos
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
                        ishasError = True
                        #给错误的地方添加红色背景颜色
                        text1.delete(pos,startPos)
                        text1.insert(pos,originStrList[i], "notEqual")
                        ft = tf.Font(family = '微软雅黑', size=14) ###有很多参数
                        text1.tag_add('notEqual',pos) #申明一个tag,在a位置使用
                        text1.tag_config('notEqual',background='dodgerblue') #设置tag即插入文字的大小,颜色等
        if ishasError:
            r = "“具体实施方式”中出现附图标记数字不正确的情况，已经标记为蓝色背景，请做进一步检查"
            result2.append(r)


    # 最终所有的结果汇总
    logging.info("result1: %s", result1)
    logging.info("result2: %s", result2)            

    # 最终的全部检查结果
    allResult = '检查报告：' + " "
    # 总体结果
    s1 = "标点符号错误:" + str(result1[0]) + "、" + "附图数错误:" + str(result1[1]) + "、" + "小标题错误:" + str(result1[2]) +"\n"
    s1= s1 + "\n"
    
    # 详细结果
    s2 = ""
    for i in range(len(result2)):
        s2 = s2 + str(i + 1) + '、' + str(result2[i]) + "\n"
        s2 = s2 + "\n"

    # 写入到文本框2中
    allResult = allResult + s1 + s2
    text2.insert(1.0,allResult)


    root.mainloop()

 

# if __name__ == '__main__':
#     main_sms()