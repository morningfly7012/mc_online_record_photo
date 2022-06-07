'''
Author: Travis CI travis@travis-ci.org
Date: 2022-06-07 11:55:02
LastEditors: Travis CI travis@travis-ci.org
LastEditTime: 2022-06-07 20:06:23
FilePath: \c:Users\morni\OneDrive - 晨飛\桌面\做成圖表\main.py
Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'''
from tracemalloc import start
import numpy as np
import matplotlib.pyplot as plt
import json
from openpyxl import load_workbook

with open ("config.json",mode="r",encoding="utf-8") as filt:
    data = json.load(filt)
startw = data["start"]
end = data["end"]
name = data["name"]
titles = data["title"]

# 讀取 Excel 檔案
wb = load_workbook('status.xlsx') #載入status.xlsx這個檔案
sheet = wb['狀態'] #選擇工作列 狀態

x = []#橫軸
y = [] #縱軸
count = startw
for i in range(end - startw): #設置x軸的東東
    a = sheet["A"+str(count)]
    c = sheet["C"+str(count)]
    ass = a.value
    x.append(ass[11:13])
    y.append(c.value)
    count = count + 1

#設定線條版權所有morningfly
line = plt.plot(x, y, color='#AE81FF', linestyle='dashed', marker='o', label='Online')

#設置圖例
plt.legend(title='Status')
plt.title(titles)
plt.savefig(name+'.png')