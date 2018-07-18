# -*- coding: utf-8 -*-

import requests
import json
import uuid
import openpyxl
from openpyxl import load_workbook
from openpyxl.compat import range
import openpyxl.styles as sty
from bears import find_maxrow,integrated_callid,request_dialogue,request_intention

# 本脚本使用已生成的智能催收多轮对话测试集，调用多轮对话接口，来测试算法的准确性，同时给出对应的错误类型
# 本脚本使用模块引入、封装、配置文件的方法

# 配置文件
configDic = {
    'WORKBOOK_PATH': 'D:/智能催收测试集0713.xlsx',
    'WORKSHEET_NAME': 'Sheet1',
    'CALL_ID': 1,
    'DIALOGUE': 4,
    'DIALOGUE_API': 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}',
    'INTENTION_API': 'http://47.95.36.52:13927/getdata?callId={}',
}

# 将配置文件中的Value存到变量中待用
wb = load_workbook(configDic['WORKBOOK_PATH'])
ws = wb[configDic['WORKSHEET_NAME']]
callid_col = configDic['CALL_ID']
dialogue_col = configDic['DIALOGUE']
dialogue_api_url = configDic['DIALOGUE_API']
intention_api_url = configDic['INTENTION_API']

# 寻找Excel的最后一行
rows = find_maxrow(ws)

# 找出对话完整的CALL_ID列表
callid_list = integrated_callid(ws, rows, callid_col, dialogue_col)

# 启动多轮对话接口
uuid_stamp = uuid.uuid4().hex
start_dialogue = request_dialogue(dialogue_api_url, '', uuid_stamp)
print(start_dialogue)

robot_answer = request_dialogue(dialogue_api_url, '我是本人', uuid_stamp)
print(robot_answer)
intention_answer = request_intention(intention_api_url, uuid_stamp)
print(intention_answer)

robot_answer1 = request_dialogue(dialogue_api_url, '忘了', uuid_stamp)
print(robot_answer1)
intention_answer = request_intention(intention_api_url, uuid_stamp)
print(intention_answer)

robot_answer2 = request_dialogue(dialogue_api_url, '我已经还了', uuid_stamp)
print(robot_answer2)
intention_answer = request_intention(intention_api_url, uuid_stamp)
print(intention_answer)

