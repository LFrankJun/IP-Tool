"""
Time:    2022/10/11
Author:  Li YuanJun
Version: V3.0
File:    main.py
Describe: function(batch file replace、zhaiyao、quanliyaoqiu、other and help)
"""

import logging
from tkinter import *
from tkinter import ttk
import tkinter.font as tf
import tkinter.messagebox
import qlyq2
import batchFile
import other
import sms
import os
import io
import re
import pdfplumber
import json
import copy
import windnd
import threading
import ctypes
import qualityCheck
from win32com import client as wc
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import concurrent_log_handler


logger = logging.getLogger()
logging.getLogger().setLevel(logging.INFO)
# logger.setLevel(level=logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(filename)s[line:%(lineno)d] - %(process)s - %(threadName)s%(thread)d: '
                              '%(message)s')

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
time_rotating_file_handler = concurrent_log_handler.ConcurrentRotatingFileHandler('mylog.log', "a", 1024*1024*1024, 20)
time_rotating_file_handler.setFormatter(formatter)

logger.addHandler(stream_handler)
logger.addHandler(time_rotating_file_handler)


def mainUI(title, message):

    # 按钮： 文字批量替换
    def batchFilePart():
        batchFile.main_bF()

    # 按钮： 文字质量检查
    def checkQualityPart():
        qualityCheck.main_qc()

    # # 按钮：检查权利要求部分
    # def checkQLYQPart():
    #     qlyq.checkQLYQ()
    
    # # 按钮： 查看说明书
    # def checkSMS():
    #     sms.main_sms()


    # # 按钮： 其他
    # def checkOther():
    #     other.main_other()
    
    #按钮： 帮助
    def help():
        
        root1 = Tk(className='帮助')

         # #调用api设置成由应用程序缩放
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        #调用api获得当前的缩放因子
        ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
        logging.info("ScaleFactor:%s", ScaleFactor)

        root1.wm_attributes('-topmost', 0)
        screenwidth, screenheight = root1.maxsize()
        scale = (ScaleFactor - 100)/100
        screenwidth = int(screenwidth*(1+scale))
        screenheight = int(screenheight*(1+scale))
        width = 580*(1+scale)
        height = 400*(1+scale)
        size = '%dx%d+%d+%d' % (width, height,
                                (screenwidth - width)/2, (screenheight - height)/2)
        root1.geometry(size)
        root1.resizable(0, 0)  # 设置X和Y轴都可改变

        # 设置标签
        fontStyle = tf.Font(size=10)

        lable1 = Label(root1, justify= CENTER, font=fontStyle, height=2)  # 标签的大小
        lable1['text'] = "功能介绍"
        lable1.pack()  # 标签的位置

        lable2 = Label(root1)  # 标签的大小
        lable2['text'] = "       本软件主要包含5个功能：关键字批量替换、检查说明书摘要文字的数量、检查权利要求书内容格式、\n 检查说明书内容格式、其他错误检查（文件中同时包含“发明”和“实用新型”关键字、重复词语、错别字）。"
        lable2.pack()  # 标签位置

        lable3 = Label(root1, justify= LEFT)  # 标签的大小
        lable3['text'] = "1、文字批量替换按钮： 自动替换word文件中权利要求书和说明书中的关键字，同时也可以对pdf格式\n的说明书附图进行标记。"
        lable3.pack(pady=(5,5))  # 标签位置

        lable4 = Label(root1, justify= LEFT)  # 标签的大小
        lable4['text'] = "2、文字质量检查按钮： 文字质量检查按钮包含多个功能，此按钮提供以下4个功能：                      "
        lable4.pack(pady=(5,5))  # 标签位置

        lable5 = Label(root1, justify= LEFT)  # 标签的大小
        lable5['text'] = "   1)、检查摘要字数功能： 自动检查word文件中说明书摘要的字数,当字数超过标准字数后，就会提示\n此问题。"
        lable5.pack(pady=(5,5))  # 标签位置

        lable6 = Label(root1, justify= LEFT)  # 标签的大小
        lable6['text'] = "   2)、检查权利要求书功能： 检查权利要求书中的格式问题，包括 查找敏感词、查找引用基础错误、生\n成引用关系、核对标点符号、判断择一引用、核对主题名称是否一致、清空内容、检查附图标记数字\n是否正确"
        lable6.pack(pady=(5,5))  # 标签位置

        lable7 = Label(root1, justify= LEFT)  # 标签的大小
        lable7['text'] = "   3)、检查说明书功能： 检查说明书格式问题，包括 查找敏感词、核对标点符号、说明书主题名称和摘\n要名称不一致、小标题缺失和顺序错误问题、检查附图标记数字是否正确。"
        lable7.pack(pady=(5,5))  # 标签位置

        lable8 = Label(root1, justify= LEFT)  # 标签的大小
        lable8['text'] = "   4)、其他检查功能： 包括 查看整个文件中是否同时包含“发明”和“实用新型”关键字、是否存在重复词\n语和错别字。"
        lable8.pack(pady=(5,5))  # 标签位置




    # 关闭窗口时的执行函数
    def close_callback():
        if tkinter.messagebox.askokcancel('信息提示', '您正在关闭主窗口！！'):
            root.destroy()


    

    # 设置弹窗尺寸
    root = Tk(className=title)
    # root.iconphoto(False, tkinter.PhotoImage(file='标题.png'))  # 设置标题图标
    # #调用api设置成由应用程序缩放
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    #调用api获得当前的缩放因子
    ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
    logging.info("ScaleFactor:%s", ScaleFactor)
    #设置缩放因子
    # root.tk.call('tk', 'scaling', ScaleFactor/120)

    root.wm_attributes('-topmost', 0)
    screenwidth, screenheight = root.maxsize()
    scale = (ScaleFactor - 100)/100
    screenwidth = int(screenwidth*(1+scale))
    screenheight = int(screenheight*(1+scale))
    width = screenwidth/3
    height = screenheight/3
    size = '%dx%d+%d+%d' % (width, height,
                            (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)
    root.resizable(0, 0)  # 设置X和Y轴都可改变

    
    # 设置标签
    # lable = Label(root, height=2)  # 标签的大小
    # lable['text'] = message
    # lable.pack()  # 标签的位置，默认是上面中间位置

   
    # 设置按钮（文字批量替换，质量检查，帮助）
    button_width = 15
    button_xpad = 30
    # 第一行区域
    frame1 = Frame(root)
    frame1.pack(side=TOP,pady=(height/3,0))

    bBtachFile = Button(frame1, height = 2, width = button_width, text ="文字批量替换", command= lambda: batchFilePart())
    bBtachFile.grid(row=0, column=0)

    bBtachFile = Button(frame1, height = 2, width = button_width, text ="文字质量检查", command= lambda: checkQualityPart())
    bBtachFile.grid(row=0, column=1, padx=(button_xpad,button_xpad))

    bHelp = Button(frame1, height = 2, width = button_width, text ="帮助", command= lambda: help())
    bHelp.grid(row=0, column=2)


    root.protocol("WM_DELETE_WINDOW", close_callback)

    

    root.mainloop()





if __name__ == '__main__':
    mainUI("主界面","")