#!/bin/env python 

notes='''
编程目标:
    电梯保养计划
        输入: 一个文件: 文件内容包含：电梯所在地项目名,首次维保日期
        输出:excel 表格: 文件内容包含: 电梯项目地址，电梯维保计划(月度,季度,半年,年度)
'''
import os
import fileinput
from datetime import datetime,timedelta
from collections import defaultdict
import xlwings as xw
import string

times=25    # 计算多少次日期，
start_date_list=[]  # 需要计算的初始日期列表
tdelta=timedelta(days=14)   # 两次保养之间的时间间隔
height=15
width=20
red=(255, 0, 0)
green=(0, 255, 0)
blue=(0, 0, 255)
yellow=(255, 255, 0)
brown=(128,64,0)
center=-4108

def get_date_dict(filename:str="address_date.txt",times:int=times): # 从 address_date.txt 获得数据，并将这个数据转化成 python 内部的数据
    """
    本函数功能:
        从 文件 address_date.txt (默认值) 中获取两列数据: 电梯项目地址 首次维保日期,并转化成 默认字典对象存储  ===》 请确保编写 这个文件时，不要数据错误
            数据举例:
吉隆税务局(1)	2024,2,16
海星湾8-11栋(8)	2024,2,8
...
        参数:
            filename: 遵守上面格式文件的文件名
            time: 需要计算的维保计划日期次数
        返回值:
            defaultdict 默认字典对象：  {"项目名1":["维保日期","维保日期2",...],"项目名2":["维保日期","维保日期2",...],...}
    """
    with fileinput.input(filename,encoding="utf-8")as f:
        d=defaultdict(list)
        for line in f:
            address_str,date_str=line.split("\t")
            try:date=datetime(*[int(i) for i in date_str.strip().split(",")])
            except Exception as e:
                raise Exception("error: address_date.txt 文件中第二列(日期数据)有问题，日期数据必须用英文逗号隔开,第二列必须与第一列数据用 tab 键盘符号隔开")
            d[address_str].append(date)
        
        for address,dt in d.items():   # 要计算的起始日期
            try:
                temp_date=dt[0]
                d[address][0]=temp_date.strftime("%Y-%m-%d")
            except Exception as e:
                raise Exception("error: address_date.txt 文件中的 日期数据 不符合日期规范，请重新检查！")
            for i in range(times):  # 列输出: 每一台的多少次日期
                temp_date+=tdelta
                d[address].append(temp_date.strftime("%Y-%m-%d"))
        return d


def set_colorAndHeader(index:int,sheet:object,syn:str,length):
    """
    本函数功能  
        将 index(代表每一列数据的编号变量)，对应的 excel 单元格 进行染色，同时给 excel 表格第一行定义 保养命名(季度，月度，年度？)
        
        参数:
            index: 略
            sheet: excel 单元格对象
            syn: excel 字母行代号变量
            length: 项目数据总数+1 值
        返回值:
            None
    """
    _range1=sheet.range(f"{syn}1")
    _range2=sheet.range(f"{syn}2:{syn}{length}")
    if index%6==1:
        _range1.value="季度保养"
        _range2.color=green  
    elif (index+1)%12==0:
        if index+1== 24:
            _range1.value="年度保养"
            _range2.color=red    
        else:
            _range1.value="半年保养"
            _range2.color=blue    
    else:
        _range1.value="月度保养"
        _range2.color=yellow      
    

def save_to_excel(filename:str="维保计划表.xlsx",data:defaultdict=None): 
    """
    本函数功能:
        保存数据到 excel 文件中
            格式: 略
        参数:
            filename 为保存的文件名
            data : 前面 get_date_dict 返回的 defaultdict 对象
        返回值:
            None
    """
    with xw.App(visible=True)as app:
        book=app.books[0]
        sheet=book.sheets[0]
        leng=len(data)+1     
        
        # 单元格格式
        _range=sheet.range          # 此处为代码优化，减少 . 访问消耗
        range_len=_range(f"A1:Z{leng}")
        range_api=range_len.api
        
        _range("A1").value="项目地址(台数)"
        _range(f"A2:A{leng}").color=_range(f"A1:Z1").color=brown
        range_api.VerticalAlignment=center       # 水平居中
        range_api.HorizontalAlignment=center     # 垂直居中
        range_api.Font.Bold=True                # 字体加粗
        range_len.row_height=height             # 设置单元格行高
        range_len.column_width =width           # 设置单元格列宽
        
        # index 表示列变量，line 表示行变量，syn 表示 横坐标的字母
        for index,syn in enumerate(string.ascii_uppercase[1:],1):
            set_colorAndHeader(index,sheet,syn,leng)
        
        # 表数据填充
        for line,tu_data in enumerate(data.items(),2):   # {'吉隆税务局(1)': ['2024-02-16', '2024-03-02', '2024-03-17'   
            _range(f"A{line}").value=tu_data[0]     # 填充最左边项目地点数据
            for index,syn in enumerate(string.ascii_uppercase[1:],1):
                _range(f"{syn}{line}").value=tu_data[1][index]      # 填充项目点电梯保养时间
        book.save(filename)
        
    
    
if __name__=="__main__":
    if os.path.exists("address_date.txt"):
        dd_data=get_date_dict(times=times)
        save_to_excel(filename="维保计划表.xlsx",data=dd_data)
    else:raise IOError("File not exits: you must create address_date.txt and input address,first_date in it")

# print((datetime(2024,2,29)+timedelta(days=15)).strftime("%Y/%m/%d"))   # 2024/03/15 



