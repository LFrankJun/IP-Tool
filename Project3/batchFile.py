"""
Time:    2022/10/11
Author:  Li YuanJun
Version: V3.0
File:   batchFile.py
Describe: Auto replace text in word document new UI
"""


from tkinter import *
import tkinter.messagebox
from win32com import client as wc
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import logging
import os
import io
import re
import pdfplumber
import windnd
import ctypes




pdfmetrics.registerFont(TTFont('SimSun', './SimSun.ttf'))  # 注册字体
logging.getLogger().setLevel(logging.INFO)

def main_bF():

    def close_callback():
        if tkinter.messagebox.askokcancel('信息提示', '您正在关闭“文字批量替换”窗口！！'):
            root.destroy()

    def writerText(bfilePath, afilePath, numList, strList):

        logging.info("numList:%s", numList)
        logging.info("strList:%s", strList)
        deleteIndexList = []  
        # 预处理一下标号（特殊情况，例如 101-106：纸卷 或者 101~106：纸卷）
        for i in range(len(numList)):
            twonumList = re.split(r'~|-', numList[i])
            if len(twonumList) == 2:
                maxNum = max(int(twonumList[0]),int(twonumList[1]))
                minNum = min(int(twonumList[0]),int(twonumList[1]))

                if str(minNum).isdigit() and str(maxNum).isdigit():
                    strWord = strList[i]
                    deleteIndexList.append(i)
                    while minNum <= maxNum:
                        numList.append(str(minNum))
                        strList.append(strWord)
                        minNum =  minNum + 1
        for index in deleteIndexList:
            del numList[index]  # 删除 101- 106
            del strList[index]  # 删除对应的附图词语
        logging.info("new numList:%s", numList)
        logging.info("new strList:%s", strList)


        # read your existing PDF
        file = open(bfilePath, "rb")
        existing_pdf = PdfFileReader(file)
        output = PdfFileWriter()
        pageNum = existing_pdf.getNumPages()  # pdf总页数
        logging.info("pageNum : %s",pageNum)
        for i in range(0, pageNum):
            # 计算pdf页面尺寸
            pdf = pdfplumber.open(bfilePath)  # 打开pdf
            page = pdf.pages[i]  # 每一页的尺寸相同，所以选择第一页
            # pageWidth = page.width    # 页面的宽度
            pageHeight = page.height  # 页面的高度
            # 存放所有要写的数字和名称,以及对应的位置坐标
            xList = []
            yList = []
            nameList = []
            # 提取每页pdf的文本
            words = pdf.pages[i].extract_words()
            for word in words:
                print("pdf中的内容:%s", word)

                # word['text']中也可能含有xx(xx)的数字形式
                valueList = []
                intList = []

                # 处理出现cid字符的情况
                textString = word['text']  # {'text': '(cid:25688)(cid:35201)(cid:38468)(cid:22270)', XXXXXXXXXXX}，textString格式类似这样：'(cid:25688)(cid:35201)(cid:38468)(cid:22270)'
                logging.info("textString: %s",textString)
                if 'cid' in textString:
                    leftIndexList = [substr.start() for substr in re.finditer(r"\(", textString)]  # 所有左边括号的位置
                    rightIndexList = [substr.start() for substr in re.finditer(r"\)", textString)]  # 所有右边括号的位置
                    combineStr = ""
                    if len(leftIndexList) == len(rightIndexList):
                        for m in range(len(leftIndexList)):
                            tempStr = textString[leftIndexList[m]:rightIndexList[m]]
                            numStr = str(tempStr).split(':')  # numStr格式：[cid, 数字]
                            if len(numStr) == 2:
                                num = numStr[1]
                                realChar = chr(int(num))  # cid字符转换为中文字符
                                combineStr = combineStr + realChar
                        textString = combineStr
                logging.info(" new textString: %s", textString)      

                for s in textString:
                    if s.isdigit() or s.encode( 'UTF-8' ).isalpha():
                        intList.append(s)
                    else:
                        if len(intList) != 0:
                            n = ''.join(intList)
                            valueList.append(n)
                            intList.clear()
                if len(intList) != 0:
                    n = ''.join(intList)
                    valueList.append(n)
                print("pdf上的同一个text里面的数字列表：%s",valueList)

                for value in valueList:
                    if value in numList:
                        wordX = word['x0']  # 以页面的左下角为原点，数字的X轴坐标
                        xList.append(wordX)
                        wordY = pageHeight - word['top']  # 以页面的左下角为原点，数字的Y轴坐标
                        yList.append(wordY)
                        numIndex = numList.index(value)  # 这个数字在附图标记数字列表中的索引,同样在附图标记汉字列表中也是这个索引
                        num_string = value + '.' + strList[numIndex]  # 要标记的信息(数字.汉字)
                        nameList.append(num_string)

            print("X轴:%s", xList)
            print("y轴:%s", yList)
            print("标记的信息:%s", nameList)
            pdf.close()  # 关闭pdf

            packet = io.BytesIO()
            # create a new PDF with Reportlab
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("SimSun", 9)  # 设置字体大小
            can.setFillColorRGB(255, 0, 0)  # 设置字体颜色
            # 把所有的信息都写入到第i页的pdf中
            # 避免Y坐标相同的几个标记互相覆盖
            tempy = 0
            sameNum = 0  # 相同的Y坐标个数
            for j in range(len(xList)):
                signText = nameList[j]
                textLength = len(signText)
                x = xList[j] - 7 * textLength  # 为避免添加的标记覆盖到原先页面上的数字
                if yList[j] != tempy:
                    y = yList[j]
                    sameNum = 0
                else:
                    sameNum = sameNum + 1
                    y = yList[j] - 9*sameNum
                tempy = yList[j]
                can.drawString(x, y, signText)

            can.save()
            # move to the beginning of the StringIO buffer
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            page = existing_pdf.getPage(i)
            # logging.info("page: %s",page)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)

            # finally, write "output" to a real file
            outputStream = open(afilePath, "wb")
            output.write(outputStream)
            outputStream.close()  # 关闭新文件
            
        file.close()  # 关闭原文件
        os.remove(bfilePath)  # 删除原文件
        os.rename(afilePath, bfilePath)  # 重命名新文件为原文件


    def auto_change_word(wordhandle, file_path, funcNum):
        replaceList3 = []  # 要标记到pdf上的字符串
        try:
            doc = wordhandle.Documents.Open(file_path)
            doc.TrackRevisions = True  # word文档进入修订模式
        except Exception as e:
            logging.info("Open Document has error!!")
        else:
            '''
            @function: 获取要替换的内容
            '''
            targetText1 = "附图标记说明"
            targetText12 = "附图标记"
            targetText13 = "标号说明"
            targetText2 = "具体实施方式"

            ftbjList = []  # 附图标记说明
            formatftbj = []  # 对不同格式的附图标记说明进行格式化
            origin_string = []
            replaceList1 = []  # 具体实施方式
            replaceList2 = []  # 权力要求部分
            notReplaceWordList = [] #不需要被替换的词语

            startIndex = endIndex = 0
        
            # 附图标记有可能是以表格的形式存在，所以先判断文章中是否有表格
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

                    # 特殊情况下，有些“附图标记”关键字和附图标号和字段处于一个自然段（eg 附图标记：1、柜体；11、开口；111、第一侧壁；）
                    text = str(doc.paragraphs[startIndex]).strip()
                    if (targetText1 in text or targetText12 in text or targetText13 in text) and bool(re.search(r'\d', text)):
                        indexList = [substr.start() for substr in re.finditer(r"：|:| ", text)] 
                        if len(indexList) != 0:
                            text = text[indexList[0] + 1:]
                            if text != '':
                                new_text = text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                                if new_text != '':
                                    ftbjList.append(new_text)
                    else:
                        j = startIndex + 1
                        while startIndex <= j < endIndex:
                            text = str(doc.paragraphs[j]).strip()
                            logging.info("text: %s", text)
                            
                            if text != '':
                                new_text = text.replace('\x07',' ').replace('\r',' ').replace('\t',' ').replace('\b',' ').strip()
                                if new_text != '':
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
            logging.info("formatftbj: %s",formatftbj)


            '''
            @function: 被替换的内容
            '''
            pdfNumList = []  # 标记到pdf上的数字
            pdfStringList = []  # 标记到pdf上的汉字
            for text in formatftbj:
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

                string_num = string + num
                string_symbol_num = string + '（' + num + '）'
                num_point_string = num + '.' + string
                origin_string.append(string)
                replaceList1.append(string_num)
                replaceList2.append(string_symbol_num)
                replaceList3.append(num_point_string)

                pdfNumList.append(num)
                pdfStringList.append(string)

            print("被替换的内容:%s", origin_string)
            print("具体实施方式替换的内容:%s", replaceList1)
            print("权利要求书替换的内容:%s", replaceList2)
            print("标记到pdf上的字符串内容:%s", replaceList3)

            if len(origin_string) > 0:
                '''
                @function:找出附图标记中有包含的词语组（eg:功率单元柜, 功率单元 ， 功率单元柜包含功率单元）
                ''' 
                # 判断includeMoreList是否已经包含此个词语
                def isdumpWord(keyStr, wLlist):
                    isFlag = False
                    for list in wLlist:
                        if keyStr in list:
                            isFlag = True
                            break
                    return isFlag

                includeMoreList= []
                for i in range(len(origin_string)):
                    includeList = []
                    if not isdumpWord(origin_string[i],includeMoreList):
                        includeList.append(origin_string[i])
                        for k in range(len(origin_string)):
                            if k != i:
                                index = origin_string[k].find(origin_string[i])
                                if index != -1:
                                    includeList.append(origin_string[k])
                        if len(includeList) > 1:
                            includeMoreList.append(includeList)

                logging.info("includeMoreList:%s", includeMoreList)


                '''
                @function:包含关系的词语无法做到全部替换，被包含的词语不能进行替换
                '''
                for inWordList in includeMoreList:
                    for m in range(len(inWordList)):
                        for n in range(len(inWordList)):
                            if n != m:
                                index = inWordList[n].find(inWordList[m])
                                if index != -1:
                                    notReplaceWordList.append(inWordList[m])
                                    break
                logging.info("notReplaceWordList: %s", notReplaceWordList)
                notReplaceWordList = list(set(notReplaceWordList))  # 避免有重复的词语出现

                '''
                @function:删除不需要替换的词语
                '''
                for w in notReplaceWordList:
                    index = origin_string.index(w)
                    del origin_string[index]
                    del replaceList1[index]
                    del replaceList2[index]



                '''
                @function:获取权利要求书和说明书所在的页眉序列
                '''
                headNameList = []
                sections = wordhandle.ActiveDocument.Sections  # 所有页眉
                for i in range(len(sections)):
                    name = wordhandle.ActiveDocument.Sections[i].Headers[0]
                    logging.info("name: %s",str(name))
                    spName = ''.join([char for char in str(name) if u'\u4e00' <= char <= u'\u9fa5'])  # 提取段落中的页眉
                    if spName != "":
                        headNameList.append(spName)
                print("所有的页眉:%s",headNameList)

                qlyqsIndex = []
                smsIndex = []
                for i in range(len(headNameList)):
                        string = headNameList[i]
                        if "权利要求书" == str(string):
                            qlyqsIndex.append(i)
                        if "说明书" == str(string):
                            smsIndex.append(i)
                logging.info("qlyqsIndex:%s",qlyqsIndex)
                logging.info("smsIndex:%s",smsIndex)


                '''
                @function:只替换权利要求书的内容或者全部替换
                '''
                if funcNum == 2 or funcNum == 4:
                    if len(qlyqsIndex) != 0:
                        # 第3节的段落
                        for index in qlyqsIndex:
                            for i in range(len(origin_string)):
                                wordhandle.ActiveDocument.Sections[index].Range.Find.Execute(origin_string[i], True, True, False, False, False,
                                                                                        True, 0, False, replaceList2[i], 2)
                    else:
                        qlyqs = "权利要求书"
                        sms = "说明书"
                        isFlag11 = False
                        isFlag12 = True
                        for i in range(len(doc.paragraphs)):
                            logging.info("doc.paragraphs[i]XX: %s",str(doc.paragraphs[i]))
                            paragraphString = ''.join([char for char in str(doc.Paragraphs[i]) if u'\u4e00' <= char <= u'\u9fa5'])  # “权 利 要 求 书”或者“，权 利 要 求 书”的格式，需要将逗号和空格去掉
                            logging.info("paragraphString: %s",paragraphString)
                            if paragraphString.strip() == qlyqs:
                                isFlag11 = True   
                            if paragraphString.strip() == sms:
                                isFlag12 = False
                            # print(str(doc.paragraphs[i]).strip())
                            if isFlag11 and isFlag12:
                                for j in range(len(origin_string)):
                                    doc.paragraphs[i].Range.Find.Execute(origin_string[j], True, True, False, False, False, True, 0, False, replaceList2[j], 2)
                        


                '''
                # @function:只替换具体实施方式中的内容或者全部内容
                '''
                if funcNum == 3 or funcNum == 4:
                    # 获取具体实施方式所在的段落行数
                    if len(smsIndex) != 0:
                        for index in smsIndex:
                            startIndex = -1
                            for i in range(len(wordhandle.ActiveDocument.Sections[index].Range.Paragraphs)):
                                if str(wordhandle.ActiveDocument.Sections[index].Range.Paragraphs[i]).strip() == '具体实施方式':
                                    startIndex = i
                                if startIndex != -1:
                                    for j in range(len(origin_string)):
                                        wordhandle.ActiveDocument.Sections[index].Range.Paragraphs[i].Range.Find.Execute(origin_string[j], True,
                                                                                                                    True, False, False, False,
                                                                                                                    True, 0, False,
                                                                                                                    replaceList1[j], 2)
                    else:
                        jtssfs = '具体实施方式'
                        isFlag21 = False
                        for i in range(endIndex, len(doc.paragraphs)):
                            if str(doc.paragraphs[i]).strip() == jtssfs:
                                isFlag21 = True
                            if isFlag21:
                                for j in range(len(origin_string)):
                                        doc.paragraphs[i].Range.Find.Execute(origin_string[j], True, True, False, False, False, True, 0, False, replaceList1[j], 2)

            else:
                tkinter.messagebox.showinfo('提示','未找到附图标记，请检查是否存在附图标记的内容')
            # '''
            # @function:退出文档
            # '''
            # doc.Close()
            # wordhandle.Quit()

        return pdfNumList, pdfStringList, notReplaceWordList


    # 进行批量替换操作
    def batchFileReplace(funcNum):

        word = wc.Dispatch("Word.Application")
        word.Visible = 1  # 0:后台运行，不显示； 1:打开文档，直接显示
        word.DisplayAlerts = 0  # 不警告

        folderP = text1.get(0.0,END)  # 文本框1中的word路径
        folderP = folderP.strip('\n')
        logging.info("batchFileReplace folderP: %s", folderP)

        folderPdfP = text2.get(0.0,END)  # 文本框2中的pdf路径
        folderPdfP = folderPdfP.strip('\n')
        logging.info("batchFileReplace folderPdfP: %s", folderPdfP)

        if folderP != "":
            folderPath = folderP.replace("\\","/")
            logging.info("batchFileReplace folderPath: %s", folderPath)
            
            if folderPath.endswith('.docx') or folderPath.endswith('.doc'):
                endHint = '批量替换处理完毕啦啦！！！'
                pdfNumList, pdfStringList, inList = auto_change_word(word, folderPath, funcNum)  # strList为要标记到pdf上的字符串内容
                if len(inList) > 0:
                    endHint = endHint + "附图标记中存在有包含关系的名词：" + str(inList) + "未在Word文件中进行替换,请注意一下！"
                if funcNum != 1 and len(pdfNumList) >= 1 and len(pdfStringList) >= 1:
                    tkinter.messagebox.showinfo('提示',endHint)

                # 在PDF上进行标记
                if funcNum == 1:
                    if folderPdfP == "":
                        tkinter.messagebox.showinfo('提示',"请拖放一个pdf文件！")
                    if folderPdfP != "" and len(pdfNumList) != 0 and len(pdfStringList) != 0:
                        folderPdfPath = folderPdfP.replace("\\","/")
                        logging.info("batchFileReplace folderPdfPath: %s", folderPdfPath)
                        if folderPdfPath.endswith('.pdf'):
                            before_filePath = folderPdfPath
                            after_filePath = folderPdfPath[:-4] + '_new' + '.pdf'
                            writerText(before_filePath, after_filePath, pdfNumList, pdfStringList)   
                            tkinter.messagebox.showinfo('提示','PDF附图标记完毕！')
        else:
            tkinter.messagebox.showinfo('提示','您未拖放一个word文件！')


   
    def drag_word(files):
        text1.delete(0.0,END)
        word_path = '\n'.join((item.decode('gbk') for item in files))
        text1.insert(0.0,word_path)
        # tkinter.messagebox.showinfo('您拖放的Word文件',word_path)


    
    def drag_pdf(files):
        text2.delete(0.0,END)
        pdf_path = '\n'.join((item.decode('gbk') for item in files))
        text2.insert(0.0,pdf_path)
        # tkinter.messagebox.showinfo('您拖放的pdf文件',pdf_path)



    # 弹窗
    root = Tk(className="文字批量替换")

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
    h = screenheight/2
    size = '%dx%d+%d+%d' % (w, h, (screenwidth - w) / 2, (screenheight - h) / 2)
    root.geometry(size)
    root.resizable(0, 0)


    # 第一行区域
    fram1 = Frame(root)
    fram1.pack(pady=(0,10))

    # 设置标签1
    lable1 = Label(fram1, height=2)
    lable1['text'] = "请将Word拖放到下面文本框中(必须)"
    lable1.grid(row=0, column=0, padx=(0,40))

    # 设置标签2
    lable2 = Label(fram1, height=2)
    lable2['text'] = "请将pdf拖放到下面文本框中(非必须)"
    lable2.grid(row=0, column=1,padx=(40,0))


    #第二行区域
    frame2 = Frame(root)
    frame2.pack()
    #设置文本框1 
    text1 = Text(frame2, height= 18, width = 30)
    text1.grid(row=0, column=0, padx=(0,40))
    # 拖放Word文件
    windnd.hook_dropfiles(text1, func=drag_word)

    #设置文本框2
    text2 = Text(frame2 , height= 18, width = 30)
    text2.grid(row=0, column=1,padx=(40,0))
    # 拖放Pdf文件
    windnd.hook_dropfiles(text2, func=drag_pdf)

    #第三行区域(1、按钮：只替换权利要求；2、按钮：只替换说明书；3、按钮：全部替换；4、按钮：智能附图)
    frame3 = Frame(root)
    frame3.pack(pady=(20,0))
    # 开始替换按钮
    writePDF = Button(frame3, width = 16, text='附图标记PDF', command= lambda: batchFileReplace(1))
    writePDF.grid(row=0, column=1)
    qlyqCheck = Button(frame3, width = 16, text='仅替换权利要求部分', command= lambda: batchFileReplace(2))
    qlyqCheck.grid(row=0, column=2, padx=(20,10))
    smsCheck = Button(frame3, width = 16, text='仅替换说明书部分', command= lambda: batchFileReplace(3))
    smsCheck.grid(row=0, column=3, padx=(10,20))
    allCheck = Button(frame3,  width = 16, text='全部替换', command= lambda: batchFileReplace(4))
    allCheck.grid(row=0, column=4)
    

    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()

    

# if __name__ == '__main__':
#     main_bF()
