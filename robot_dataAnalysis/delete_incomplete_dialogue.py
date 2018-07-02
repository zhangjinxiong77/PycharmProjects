# -*- coding: utf-8 -*-

import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
import datetime
import openpyxl.styles as sty
from openpyxl import Workbook

def delete_incomplete_dialogue():
    # 自定义检测Excel的A列最大行函数。openpyxl的worksheet.max_row并不准
    def find_maxrow(worksheet):
        i = 1
        while worksheet['A' + str(i)].value != None:
            i += 1
        return i - 1

    # 打开Excel
    wb = load_workbook('D:/智能催收数据标注0622-0626.xlsx')
    ws = wb['Sheet1']
    rows = find_maxrow(ws)

    # 将带有 “再见”的行ID加入List
    completeLi= []
    for row in range(2, rows+1):
        # 对话数据必须放在D列.
        for col in range(4, 5):
            if "再见" in ws.cell(column=col,row=row).value:
                completeLi.append(ws.cell(column=1,row=row).value)

    # 遍历Excel，如果对话ID在completeLi中，则将对话拼接入以对话ID为Key，List[对话时间，姓名，对话语句拼接]为值的字典中。
    labelDataDic = {}
    for row in range(2, 1000):
        for col in range(1, 2):
            if ws.cell(column=col,row=row).value in completeLi:
                if ws.cell(column=col,row=row).value in labelDataDic.keys():
                    labelDataDic[ws.cell(column=col,row=row).value][2] += '\n' + ws.cell(column=col+3 ,row=row).value
                else:
                    labelDataDic[ws.cell(column=col,row=row).value] = [ws.cell(column=col+1, row=row).value.strftime('%Y-%m-%d %H:%M:%S'), ws.cell(column=col+2 ,row=row).value,ws.cell(column=col+3 ,row=row).value]

    print(labelDataDic)

    # 将字典存入Excel。因为字典无序不能循环，所以先将字典转换为List中的两位子tuple
    sumLi = []
    for kv in labelDataDic.items():
        sumLi.append(kv)

    wb2 = Workbook()
    ws2 = wb2.active

    for row in range(1, len(labelDataDic) + 1):
        for col in range(1, 2):
            ws2.cell(column=col, row=row).value = sumLi[row - 1][0]
            ws2.cell(column=col + 1, row=row).value = sumLi[row - 1][1][0]
            ws2.cell(column=col + 2, row=row).value = sumLi[row - 1][1][1]
            ws2.cell(column=col + 3, row=row).value = sumLi[row - 1][1][2]

    wb2.save(filename='D:/完整的对话语句拼接.xlsx')


if __name__=='__main__':
    delete_incomplete_dialogue()