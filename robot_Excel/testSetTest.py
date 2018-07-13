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


# 本脚本用来测试每一句话一行（而不是一通电话一行）的多轮对话测试集
str1='机器人：您好，请问您是机主本人吗'
newstr=str1[:3]
print(newstr)

# uuid_stamp = uuid.uuid4().hex
# ask1 = ''
# dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format(ask1, uuid_stamp)
# dialogue_res = requests.get(dialogueAPI_url).text
# dialogue_res = json.loads(dialogue_res)
# print(dialogue_res['data']['answer'])
#
# ask2 = '是的'
# dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format(ask2, uuid_stamp)
# dialogue_res = requests.get(dialogueAPI_url).text
# dialogue_res = json.loads(dialogue_res)
# intention_url = 'http://47.95.36.52:13927/getdata?callId={}'.format(uuid_stamp)
# intention_res = requests.get(intention_url).text
# intention_res = json.loads(intention_res)
# intention_res = intention_res['data']['intention']
# print(intention_res)
# print(dialogue_res['data']['answer'])
#
# ask3 = '忘了'
# dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format(ask3, uuid_stamp)
# dialogue_res = requests.get(dialogueAPI_url).text
# dialogue_res = json.loads(dialogue_res)
# intention_res = requests.get(intention_url).text
# intention_res = json.loads(intention_res)
# intention_res = intention_res['data']['intention']
# print(intention_res)
# print(dialogue_res['data']['answer'])
#
# ask4 = '我已经还了'
# dialogueAPI_url = 'http://47.94.52.113:5061/api/v0.1/ask?project_id=48334234b5544627aa1cc977eec6cec6&query={}&uid={}'.format(ask4, uuid_stamp)
# dialogue_res = requests.get(dialogueAPI_url).text
# dialogue_res = json.loads(dialogue_res)
# intention_res = requests.get(intention_url).text
# intention_res = json.loads(intention_res)
# intention_res = intention_res['data']['intention']
# print(intention_res)
# print(dialogue_res['data']['answer'])
