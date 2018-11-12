# -*- coding: utf-8 -*-
"""
Created on Mon Nov 12 16:46:25 2018

@author: DELL
"""

fi = open('学生未刷卡报表.txt','r')
context = fi.readlines()
lines = list()
for line in context:
    lines.append(line.split())
schools = list(set([i[2] for i in lines]))
di = {}
for line in lines:
    if line[2] not in di:
        di[line[2]] = []
    else:
        di[line[2]].append(line)
head = ['学号', '姓名', '院系', '楼栋', '宿舍', '最后刷卡日期', '最后刷卡时间', '没刷卡天数']
for key in di:
    name = 'new\未刷卡报表'+'_'+str(key)+'.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in di[key]:
            line = ','.join(item)
            f.write(line+'\n')
fi.close()