import json
import re
import pprint

import openpyxl

# path = r"C:\Users\asuka\Desktop"
# os.chdir(path)  # 修改工作路径

workbook = openpyxl.load_workbook('46届ICPC(昆明)中文问卷报名统计表.xlsx')  # 返回一个workbook数据类型的值
sheet = workbook.active  # 获取活动表
print('当前活动表是：' + str(sheet))

with open('rating.json','r',encoding='utf8')as fp:
    json_data = json.load(fp)

json_data = dict(json_data)

# pprint.pprint(dict(json_data))

teams = {}

# pprint.pprint(json_data[0])



for i in json_data:
    team = json_data[i]
    for h in team['history']:
        teams[h['teamName']] = {'history': [], 'best_rank': 9999}


for i in json_data:
    team = json_data[i]
    best_rank = 99999
    for h in team['history']:
        teams[h['teamName']]['history'].append({
            'rank': h['rank'],
            'contestName': h['contestName']
        })
        teams[h['teamName']]['best_rank'] = min(teams[h['teamName']]['best_rank'], int(h['rank']))

# pprint.pprint(teams)

res = []

for i in sheet.iter_rows(min_row=2, max_row=884, min_col=4, max_col=4):
    for j in i:
        if j.value in teams.keys():
            res.append([{
                j.value: teams[j.value]['history']
            }, teams[j.value]['best_rank']])
        else:
            res.append([j.value, 999999])


# pprint.pprint()

res.sort(key=lambda x: x[1])
#
#
pprint.pprint(res)