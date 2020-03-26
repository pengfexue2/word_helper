#!/usr/bin/env python
# encoding: utf-8
# @Time : 2020-03-20 16:05

__author__ = 'Ted'

from docx import Document
import docx
import pandas as pd


# 一、读取 xlsx 表格数据
excel_path="词频更新表.xlsx"
data = pd.read_excel(excel_path,sheet_name='Sheet1')

excel_dict={}
for i,item in enumerate(data["Column3"]):
    if isinstance(item,str):
        excel_dict[item]=data["Column2"][i]

print(excel_dict)

# 二、读取 word 文档
path="高考词频表.docx"
document = Document(path)

# 读取文档中的所有表格
tables = document.tables
# 获取所有表格数
table_num = len(tables)

# 为所有单词建立对应的词频字典
f_dict = {}


for table_index in range(table_num):
    table = tables[table_index]
    for i in range(1,len(table.rows)):
        word_text = table.cell(i, 0).text
        frequency = table.cell(i, 4).text
        f_dict[word_text] = frequency

        if frequency!="":
            #print(f"word中{word_text}的频率为{frequency}")
            update_num = int(excel_dict.get(word_text,0))
            #print(f"excel中{word_text}的频率为{update_num}")
            count = int(frequency)+update_num
            table.cell(i, 4).text = str(count)
            #print(f"更新完，词频为{count}")
    print(f"已完成 {table_index+1}/{table_num+1}")

print("更新词频完成！")
document.save("result.docx")









