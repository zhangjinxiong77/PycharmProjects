# -*- coding: utf-8 -*-

import time
import requests
import json
import re
import uuid
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
import openpyxl.styles as sty

# 本脚本使用已生成的智能催收多轮对话测试集，调用多轮对话接口，来测试算法的准确性，同时给出对应的错误类型
# 自定义检测Excel的A列最大行函数。openpyxl的worksheet.max_row并不准
def find_maxrow(worksheet):
    i = 1
    while worksheet['A' + str(i)].value != None:
        i += 1
    return i-1

# 打开Excel
wb = load_workbook('D:/智能催收测试集0713.xlsx')
ws = wb['Sheet1']
rows = find_maxrow(ws)

# 判断哪些Call_ID是对话完整的——将带有 “再见”的行ID加入List
completeLi = []
for row in range(2, rows + 1):
    # 对话数据必须放在D列.A列放CALL_ID
    for col in range(4, 5):
        if "再见" in ws.cell(column=col, row=row).value:
            completeLi.append(ws.cell(column=1, row=row).value)

print('完整的对话共'+str(len(completeLi))+'条')
# 开始执行循环
resultDic={}
for row in range(2, 100):
    for col in range(1, 2):
        print(row)
        if ws.cell(column=col,row=row).value in completeLi:
            # 如果是第一句话，则创建新的UUID，并发送''给接口，以开启多轮对话
            dialogue = ws.cell(column=col+3,row=row).value
            if dialogue == '机器人：您好，请问您是机主本人吗':
                uuid_stamp = uuid.uuid4().hex
                dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format('', uuid_stamp)
                dialogue_res = requests.get(dialogueAPI_url).text
                ws.cell(column=col+5, row=row).value = dialogue
                print(dialogue)
            else:
                if dialogue[:2] == '客户':
                    if resultDic[ws.cell(column=col,row=row).value] != "错误":
                        dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format(dialogue, uuid_stamp)
                        dialogue_res = requests.get(dialogueAPI_url).text
                        dialogue_res = json.loads(dialogue_res)
                        robotAnswer = dialogue_res['data']['answer']
                        if robotAnswer != ws.cell(column=col+3, row=row+1).value:
                            resultDic[ws.cell(column=col,row=row).value] = "错误"
                            print(dialogue)
                            print('错误回答：'+ robotAnswer)
                        else:
                            print(dialogue)
                else:
                    if resultDic[ws.cell(column=col, row=row).value] != "错误":
                        print(dialogue)




