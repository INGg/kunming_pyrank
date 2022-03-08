import json
import re
import pprint

import openpyxl

# path = r"C:\Users\asuka\Desktop"
# os.chdir(path)  # 修改工作路径

workbook = openpyxl.load_workbook('46届ICPC(昆明)中文问卷报名统计表.xlsx')  # 返回一个workbook数据类型的值
sheet = workbook.active  # 获取活动表

with open('rating.json','r',encoding='utf8')as fp:
    json_data = json.load(fp)

json_data = dict(json_data)

teams = {}

for i in json_data:
    team = json_data[i]
    ranks = []
    for h in team['history']:
        ranks.append(h['rank'])
    for h in team['history']:
        teams[h['teamName']] = [team["rating"], team["handle"], ranks]

res = []
seen=set()
for i in sheet.iter_rows(min_row=2, max_row=884, min_col=4, max_col=4):
    for j in i:
        if j.value in teams.keys():
            if j.value not in seen:
                res.append([j.value, teams[j.value]])
                seen.add(j.value)
        else:
            res.append([j.value, [0, 0, None]])

def get_har_mean(ranks):
    if not ranks:
        return 999999
    sum = 0
    for i in ranks:
        sum += 1/i
    return len(ranks)/sum

res.sort(key=lambda x: get_har_mean(x[1][2]))
# res.sort(key=lambda x: -x[1][0])

def get_aligned_string(string,width):
    string = "{:{width}}".format(string,width=width)
    bts = bytes(string,'utf-8')
    string = str(bts[0:width],encoding='utf-8',errors='backslashreplace')
    new_width = len(string) + int((width - len(string))/2)
    if new_width!=0:
        string = '{:{width}}'.format(str(string),width=new_width)
    return string[:-4]

for i in range(1, 21):
    cur =get_aligned_string(res[i-1][0], 32)
    if i<=2:
        print(f"rank ??: {cur}ranks: {res[i-1][1][2]}  member: {res[i-1][1][1]}")
    else:
        print(f"rank {i:2}: {cur}ranks: {res[i-1][1][2]}   member: {res[i-1][1][1]}")
    # print(f"{res[i-1][1][2]}")
