# encoding: utf-8
import os
import pandas as pd
from docxtpl import DocxTemplate
import xlrd

filepath = "待调整格式信息.xls"

gstype = input("请选择要调整格式的字号类型，1为小三号字体，2为小四号字体：")
if gstype == "1":
    tplfilepath = 'file\调整格式模板（小三号）.docx'
else:
    tplfilepath = 'file\调整格式模板（小四号）.docx'
while not os.path.exists(filepath):
    txt = input("未在当前文件夹下找到文件“" + filepath + "”。回车键继续查找")
while not os.path.exists(tplfilepath):
    txt = input("未在当前文件夹下找到文件“" + tplfilepath + "”。回车键继续查找")


print("程序正在执行……")
contents = {'results': []}
df = pd.read_excel(filepath)
i = 1
mytype = 0

def unilen(txt):
    lenTxt = len(txt)
    lenTxt_utf8 = len(txt.encode('utf-8'))
    return int((lenTxt_utf8 - lenTxt) / 2 + lenTxt)

for index, row in df.iterrows():
    biaoti = str(i) + "." + str(row['标题'])
    danwei = str(row['单位'])
    txtlen = unilen(biaoti + danwei + "（）")
    if gstype == "1":
        if unilen(danwei) >= 52:
            mytype = 1  # 单位一整行或超过一行
        elif txtlen <= 56:
            mytype = 0   # 单位和标题在一行
            biaoti = biaoti.ljust(56-unilen(danwei)-4-len(biaoti)+3)
        else:
            mytype = 0  # 单位和标题超过一行
            if txtlen <= 112:
                biaoti = biaoti.ljust(112-unilen(danwei)-4-len(biaoti)+2)
            else:
                biaoti = biaoti.ljust(166 - unilen(danwei)-4 - len(biaoti)+2)
    else:
        if unilen(danwei) >= 66:
            mytype = 1  # 单位一整行或超过一行
        elif unilen(biaoti) <=70 and txtlen > 70 and txtlen < 140:
            mytype = 2  # 标题和单位各占一行
        elif unilen(biaoti) <=138 and txtlen > 140 :
            mytype = 2  # 标题和单位各占一行
        elif txtlen <= 70:
            mytype = 0   # 单位和标题在一行
            biaoti = biaoti.ljust(70-unilen(danwei)-4-len(biaoti)+3)
        else:
            mytype = 0  # 单位和标题超过一行
            if txtlen <= 140:
                biaoti = biaoti.ljust(140-unilen(danwei)-4-len(biaoti)+2)
            else:
                biaoti = biaoti.ljust(280 - unilen(danwei)-4-len(biaoti)+2)


    i += 1

    contents['results'].append({'biaoti': biaoti, 'danwei': danwei, 'type': mytype})
# print(contents)

# print(tplfilepath)
tpl = DocxTemplate(tplfilepath)
tpl.render(contents)
tpl.save('格式调整结果.docx')
print("已完成格式调整。\n结果请查看“格式调整结果.docx”。\n按回车关闭")