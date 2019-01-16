# -*- coding: utf-8 -*-
"""
Created on Mon Nov 12 16:46:25 2018

@author: DELL
#进修生名单
'''
==两名单共有研究生===
魏浩然	进修生	博士	201880000050	医学院
张运帷	进修生	硕士	201880000006	医学院
龙杰	进修生	博士	201880000019	医学院	
杨淑涵	进修生	硕士	201880000007	医学院
晏凯	进修生	博士	201880000021	医学院	
沙锐	进修生	博士	201880000022	医学院
王连生	进修生	博士	201880000023	医学院
赵志斌	进修生	博士	201888000024	医学院
王燕	进修生	硕士	201880000001	医学院
刘清枝	进修生	博士	201880000002	医学院
高采月	进修生	博士	201880000003	医学院
刘榴	进修生	博士	201880000004	医学院
范亚楠	进修生	硕士	201880000011	医学院
罗英丽	进修生	博士	201880000012	医学院
张玥	进修生	硕士	201880000013	医学院
赵贵	进修生	硕士	201880000027	医学院	
陈之尧	进修生	硕士	201880000028	医学院
陈凯歌	进修生	博士	201880000031	医学院	
曹志婷	进修生	博士	201880000014	医学院
王继龙	进修生	博士	201880000029	医学院	
张厚兵	进修生	硕士	201880000030	医学院
=====两名单共有本科生=======
王颖辉	进修生	本科生	201880000052	医学院 ===> 现为研究生2019.01.07
徐雅菲	进修生	本科	201880000008	医学院
蒋荣华	进修生	本科	201880000018	环境与能源学院
=======第一版名单有的人（记）=====
杨冬冬	进修生	博士	320103198805012523+杨冬冬	医学院		
李晶	进修生	本科	201892990903	经济与贸易学院	
李军华	进修生	本科		经济与贸易学院
甘蕴久	进修生	博士	420606199409070013+甘蕴久	医学院
====第二版名单有的人（不记）=====
李彭轩 学生 本科 201830320188 材料科学与工程学院
黄薇娴 其他 硕士 201880000036 研究生院（研究生工作部）
"""

import datetime
import pandas as pd

#删除的学号
deleted_num = ('201820137853','201630741688','201720145044','201530642115','201636599375','201721045749','201630834731','201630855354','201820140855','201821038681','201536761179','201720148229','201820140666','201820140873')

#进修生本科
jinxiu_ben = ('201892990903','201880000008','201880000018')

#进修生研究生
jinxiu_yan = ('201880000052','320103198805012523+杨冬冬','201880000050','201880000006','201880000019','201880000007','201880000021','420606199409070013+甘蕴久','201880000022','201880000023','201888000024','201880000001','201880000002','201880000003','201880000004','201880000011','201880000012','201880000013','201880000028','201880000031','201880000014','201880000029','201880000030','201880000027')

#所有学院名单
schools_all = [
'机械与汽车工程学院',
'建筑学院',
'土木与交通学院',
'电子与信息学院',
'材料科学与工程学院',
'化学与化工学院',
'轻工科学与工程学院',
'食品科学与工程学院',
'数学学院',
'物理与光电学院',
'经济与贸易学院',
'自动化科学与工程学院',
'计算机科学与工程学院',
'电力学院',
'生物科学与工程学院',
'环境与能源学院',
'软件学院',
'工商管理学院',
'公共管理学院',
'马克思主义学院',
'外国语学院',
'法学院（知识产权学院）',
'新闻与传播学院',
'艺术学院',
'体育学院',
'设计学院',
'医学院',
'国际教育学院',
'继续教育学院',
'生物医学科学与工程学院',
'吴贤铭智能工程学院']

#20181227 23:00-20181228 6:00 星期四 晚删除 合唱团学生名单
hechangtuan = [
'唐海琳',
'邓相林',
'彭慧娴',
'张文敏',
'杨凯恩',
'雷斯雨',
'彭晓丽',
'张艺菲',
'李明诗',
'杜涵涵',
'吴双瑾儿',
'周雨薇',
'周芷茵',
'温佳妮',
'何双欣',
'肖泊秋',
'周岐荣',
'尹艺蓓',
'黄若宾',
'汤清茜',
'温翠娜',
'刘议谦',
'肖牧青',
'陈秋彤',
'许露云',
'王紫溶',
'张舒婷',
'范余奥',
'曹隽然',
'李枝蔓',
'刘莎',
'邹馨雨',
'宋曼嘉',
'陈滢',
'江俞莹',
'陈思曼',
'王嘉咏',
'刘思嘉',
'黄淑美',
'胡蝶',
'彭梦梦',
'李鸿康',
'傅宇尧',
'李易轩',
'刘欢',
'范睿宇',
'李聪',
'范俞辰',
'杜沛林',
'徐啟铭',
'高超渝',
'龙煦涛',
'张扬',
'何绍玮',
'侯镔宸',
'林俊良',
'胡宜鹏',
'郝晓光',
'王悄',
'陈瑞森',
'刘嘉昊',
'陈涵',
'罗博文',
'房志宏',
'何亚',
'吴宏宇',
'杨晨',
'余铮逸',
'朱宇鹏']



def split_weishuaka_table():
    fi = open('学生未刷卡报表.txt','r')
    context = fi.readlines()
    lines = list()
    for line in context:
        lines.append(line.split())
    #schools = list(set([i[2] for i in lines]))
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
    
def split_table(folder,source_file,postfix):
    sf = open(source_file,'r',encoding='gbk')
    context = sf.readlines()
    lines = list()
    for line in context:
        lines.append(line.strip().split(','))
    #schools = list(set([i[2] for i in lines]))
    di = {}
    for line in lines:
        if line[2] not in di:
            di[line[2]] = [line]
        else:
            di[line[2]].append(line)
    head = ['学号', '姓名', '院系', '楼栋', '宿舍', '通道', '进出记录', '日期','时间','本科生\研究生']
    title = "学生进出流水记录报表"+postfix
    for key in di:
        name = folder+"\\"+key[1:-1]+"晚归记录报表"+postfix+".csv"
        with open(name,'w') as f:
            f.write(title+'\n')
            f.write(','.join(head)+'\n')
            for item in di[key]:
                line = ','.join(item)
                f.write(line+'\n')
            f.close()
    sf.close()
    
def gen(filename='1125_1202.txt',fri=(2018,11,30)):
    data = open(filename,'r',encoding='utf-8')
    context = data.readlines()
    data.close()
    lines = list()
    #设置周末时间节点
    fri_2300 = datetime.datetime(fri[0],fri[1],fri[2],23,0,0)
    fri_2330 = fri_2300 + datetime.timedelta(minutes=30)
    sat_0000 = fri_2300 + datetime.timedelta(hours=1)
    sat_2300 = fri_2300 + datetime.timedelta(days=1)
    sat_2330 = sat_2300 + datetime.timedelta(minutes=30)
    sun_0000 = sat_2300 + datetime.timedelta(hours=1)
    
    ttt1 = datetime.datetime(1900,1,1,23,0,0)
    ttt2 = datetime.datetime(1900,1,1,23,30,0)
	
	#设置每天的时间戳23:00到次日6:00，本周日23点-下一个周日6点
	#例1125 23:00-1202 6:00 timeList[1125-26,1126-27,1127-28,1128-29,1129-30,1130-31,1201-02]共7条
    first_day = fri_2300 + datetime.timedelta(days=-5)
    timeList = list()
    for i in range(7):
        print(i)
        time_start = first_day + datetime.timedelta(days=i)
        time_end = time_start + datetime.timedelta(hours=8)
        timeList.append((time_start,time_end))
    
    for line in context:
        temp = line.strip()
        tt = temp.split()
        num = tt[0]
        name = tt[1]
        school = tt[2]
        dorm = tt[3] #楼栋
        #删掉学号不为12位的和非学院人士
        if len(num) != 12 and num not in ('320103198805012523+杨冬冬','420606199409070013+甘蕴久')or school in ('学生工作部（处）','珠海葆力物业有限公司广州分公司华工服务中心'):
            print(tt)
            continue
        #删掉部分晚归记录，包括自己
        if num in deleted_num:
            print(tt)
            continue
        #### 2018.12.20修改，只要c8西的记录
#        if dorm not in ['C8西']:
#            continue


        #删掉日期不符晚归的人士：周末23:00-23:30
        dt_str = tt[7]+' '+tt[8]
        dt = datetime.datetime.strptime(dt_str,"%Y/%m/%d %H:%M:%S")
        if dt >= fri_2300 and dt <= fri_2330:
            print(dt)
            continue
        if dt >= sat_2300 and dt <= sat_2330:
            print(dt)
            continue
		
		#### 2018.12.27修改 删除合唱团当晚晚归记录
#        if name in hechangtuan and dt >= timeList[4][0] and dt <= timeList[4][1]:
#            print(name)
#            print(dt)
#            continue
        
        #### 2019.01.06修改 删除元旦节三天记录
        #if dt >= dt >= timeList[0][0] and dt <= timeList[0][1]:
            #continue
        #if dt >= dt >= timeList[1][0] and dt <= timeList[1][1]:
            #continue
        #if dt >= dt >= timeList[2][0] and dt <= timeList[2][1]:
            #continue
            
        
		#设置身份
        if int(num[4]) < 3:
            mark = '研究生'
        if int(num[4]) >= 3:
            mark = '本科生'
        if num[:5] == '20188' and num not in ('201880000027', '201880000020', '201880000008','201880000005'):
            mark = '研究生'
            
        #标记进修生
        if num in jinxiu_ben:
            mark = '进修生(本科生)'
        if num in jinxiu_yan:
            mark = '进修生(研究生)'
            
        if name == '李军华' and school == '经济与贸易学院':
            mark = '进修生(本科生)'

        #删掉日期不符晚归的研究生：周末24:00前
        if mark in ('研究生','进修生(研究生)') and dt >= fri_2300 and dt <= sat_0000:
            print(dt)
            continue
        if mark in ('研究生','进修生(研究生)') and dt >= sat_2300 and dt <= sun_0000:
            print(dt)
            continue
        
        ### 20190113修改 提取军训学生记录
        if not mark in ('本科生'):
            continue
        
        if not num[:4] == '2018':
            continue
        
        tt.append(mark)
        lines.append(tt)
    
    newLines = list()
    for line in lines:
        time = datetime.datetime.strptime(line[8],"%H:%M:%S")
        mark = line[9]
        #研究生周一到周五23:30之前回来不算晚归
        if mark in ('研究生','进修生(研究生)') and time >= ttt1 and time <= ttt2:
            print(line)
            continue
        newLines.append(line)
        
    #按照学号划分条目
    num_dic = {}
    for line in newLines:
        if line[0] not in num_dic:
            num_dic[line[0]] = [line]
        else:
            num_dic[line[0]].append(line)
            

        

    ddic = []
    for num in num_dic:
        values = num_dic[num]
        time_dic = {}
        #for time in timeList:
        #    time_dic[time[0]] = []
        for line in values:
            dt = datetime.datetime.strptime(line[7]+' '+line[8],"%Y/%m/%d %H:%M:%S")
            for time in timeList:
                if dt >= time[0] and dt <= time[1]:
                    if time[0] not in time_dic:
                        time_dic[time[0]] = [line]
                    else:
                        time_dic[time[0]].append(line)
        ddic.append(time_dic)
        
    return ddic
    
def gen1129_nqc(ddic,desc='ffff'):
    wangui = []
    likai = []
    jinchu = []
    for person in ddic:
        #每一个person是一个dict
        for day in person:
            records = person[day]
            actions = [i[6] for i in records]
            if actions[0] == '离开宿舍':
                likai.extend(records)
            else:
                count = 0
                for i in range(len(records)):
                    if actions[i] != '进入宿舍':
                        count = i
                if count == 0:
                    wangui.extend(records)
                else:
                    jinchu.extend(records)
                    
    head=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间']
    name = desc+'\wangui.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in wangui:
            line = ','.join(item)
            f.write(line+'\n')
            
        
    name = desc+'\likai.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in likai:
            line = ','.join(item)
            f.write(line+'\n')
            
        
    name = desc+'\jinchu.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in jinchu:
            line = ','.join(item)
            f.write(line+'\n')
  
    return wangui,likai,jinchu

def gen1129_quchong(ddic,desc='ffff'):
    wangui = []
    likai = []
    jinchu = []
    for person in ddic:
        #每一个person是一个dict
        for day in person:
            records = person[day]
            actions = [i[6] for i in records]
            if actions[0] == '离开宿舍':
                likai.append(records[0])
            else:
                count = 0
                for i in range(len(records)):
                    if actions[i] != '进入宿舍':
                        count = i
                if count == 0:
                    wangui.append(records[0])
                else:
                    jinchu.append(records[i])
                    
    head=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间']
    name = desc+'\wanguiq.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in wangui:
            line = ','.join(item)
            f.write(line+'\n')
            
        
    name = desc+'\likaiq.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in likai:
            line = ','.join(item)
            f.write(line+'\n')
            
        
    name = desc+'\jinchuq.csv'
    with open(name,'w') as f:
        f.write(','.join(head)+'\n')
        for item in jinchu:
            line = ','.join(item)
            f.write(line+'\n')
  
    return wangui,likai,jinchu
    
def main(postfix='(2018年11月25日-2018年12月02日)',filename='1125_1202.txt',fri=(2018,11,30),desc='ff',output='学院总表1125_1202.xlsx'):
    '''
    参数：
    postfix:学院分表的日期后缀
    filename:读入的txt形式的excel报表，以utf-8编码
    fri:一周中周五的日期，如11-25-12-02周五就是1130
    desc:输出学院分表存放的目标文件夹
    '''
    ddic = gen(filename=filename,fri=fri)
    wan,li,jc = gen1129_quchong(ddic,desc=desc)
    wangui,likai,jinchu = gen1129_nqc(ddic,desc=desc)
    
    writer = pd.ExcelWriter(output)
    #统计各学院的人数
    wand = pd.DataFrame(wan,columns=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间','本科生/研究生'])
    wand = wand[(wand['本科生/研究生'] == '本科生')|(wand['本科生/研究生'] == '进修生(本科生)') ]
    wandd = pd.DataFrame(wand['院系'].value_counts())
    #增加统计12点之后晚归的人数
    wand['time'] = pd.to_datetime(wand['时间'])
    wanda12 = wand[wand['time'] < pd.to_datetime('22:00:00')]
    wanda12d = pd.DataFrame(wanda12['院系'].value_counts())
    #离开，进出
    lid = pd.DataFrame(li,columns=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间','本科生/研究生'])
    lid = lid[(lid['本科生/研究生'] == '本科生')|(lid['本科生/研究生'] == '进修生(本科生)') ]
    lidd = pd.DataFrame(lid['院系'].value_counts())
    jcd = pd.DataFrame(jc,columns=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间','本科生/研究生'])
    jcd = jcd[(jcd['本科生/研究生'] == '本科生')|(jcd['本科生/研究生'] == '本科生')]
    jcdd = pd.DataFrame(jcd['院系'].value_counts())
    #修改格式
    jcdd['school'] = jcdd.index
    lidd['school'] = lidd.index
    wandd['school'] = wandd.index
    wanda12d['school'] = wanda12d.index
    
    jcdd.columns=['门禁后曾离开过宿舍人数','院系']
    lidd.columns=['夜不归宿人数','院系']
    wandd.columns=['晚归人数','院系']
    wanda12d.columns = ['12点后晚归人数','院系']
    
    new = pd.merge(wandd,jcdd,on=['院系'],how='outer')
    new = pd.merge(new,lidd,on=['院系'],how='outer')
    new = pd.merge(new,wanda12d,on=['院系'],how='outer')
    new = new[['院系','晚归人数','12点后晚归人数','夜不归宿人数','门禁后曾离开过宿舍人数']]
    new = new.fillna(0)
    
    #自定义学院顺序
    new['院系'] = new['院系'].astype('category')
    new['院系'].cat.set_categories(schools_all, inplace=True)
    new = new.sort_values('院系', ascending=True)
        
    
    new.to_excel(excel_writer=writer, sheet_name = '分院系人数汇总',header=True,index=False)
    #晚归、离开、进出分开的以学院为key的dict
    di = {}
    for line in wangui:
        if line[2] not in di:
            di[line[2]] = [line]
        else:
            di[line[2]].append(line)
    dij = {}
    for line in jinchu:
        if line[2] not in dij:
            dij[line[2]] = [line]
        else:
            dij[line[2]].append(line)
    dil = {}
    
    for line in likai:
        if line[2] not in dil:
            dil[line[2]] = [line]
        else:
            dil[line[2]].append(line)
    #统计数据中出现过的所有学院
    schools_real = list(set(list(di.keys())+list(dij.keys())+list(dil.keys())))
    #看看总学院表中是否有遗漏
    for school in schools_real:
        if school not in schools_all:
            print(school)
    #按学院制作分sheet
    for school in schools_all:
        if school not in schools_real:
            print(school)
            continue
        sheet = []
        head_ben=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间','本科生/研究生','原因','学院处理情况反馈（口头批评教育、拟书面通报批评、拟处分等）']
        head_yan=['学号','姓名','院系','楼栋','宿舍','通道','进出记录','日期','时间','本科生/研究生']
        values_wangui = []
        values_jinchu = []
        values_likai = []
        if school in di:
            values_wangui = di[school]
        if school in dij:
            values_jinchu = dij[school]
        if school in dil:
            values_likai = dil[school]
        #改成4部分
        benke_wangui = []
        benke_jinchu = []
        benke_likai = []
        yanjiusheng = []
        
        for value in values_wangui:
            if value[-1] in ('本科生','进修生(本科生)'):
                benke_wangui.append(value)
            else:
                yanjiusheng.append(value)
        for value in values_jinchu:
            if value[-1] in ('本科生','进修生(本科生)'):
                benke_jinchu.append(value)
            else:
                yanjiusheng.append(value)
                pass
        for value in values_likai:
            if value[-1] in ('本科生','进修生(本科生)'):
                benke_likai.append(value)
            else:
                yanjiusheng.append(value)
                pass

        sheet.extend([])
        sheet.append(head_ben)
        sheet.extend(benke_wangui)
        sheet.extend([])
        sheet.append(head_ben)
        sheet.extend(benke_likai)
        sheet.extend([])
        sheet.append(head_ben)
        sheet.extend(benke_jinchu)
        sheet.append(head_yan)
        sheet.extend(yanjiusheng)
        df = pd.DataFrame(sheet)
        df.to_excel(excel_writer=writer, sheet_name = school, index=False,header=False)
## 删除分表导出
#        temp_name = desc+"\\"+school+"晚归记录报表"+postfix+".xlsx"
#        temp_writer = pd.ExcelWriter(temp_name)
#        df.to_excel(excel_writer=temp_writer, sheet_name = school, index=False,header=False)
#        temp_writer.save()
#        temp_writer.close()
#        all_data.append(df)
    writer.save()
    writer.close()
    
def compair(txt1='new.txt',txt2='old.txt'):
    new = open(txt1,'r',encoding='utf-8')
    old = open(txt2,'r',encoding='utf-8')
    context_new = new.readlines()
    context_old = old.readlines()
    return context_new,context_old


def process_xideng(filepath='未熄灯宿舍检查汇总表.xlsx',output='(2018年12月14日-2018年12月22日)',desc='ff'):
    #先检查数据格式，尤其是日期要统一格式，否则转换会出问题
    data = pd.read_excel('未熄灯宿舍检查汇总表.xlsx',header=1,skiprows=0,names=['日期', '楼栋', '宿舍号', '学院', '检查人', '备注'])
    #统计各宿舍出现次数
    gr1=data.groupby(['楼栋','宿舍号','学院'])
    si = gr1.size().reset_index()
    si.rename(columns={0:'次数'}, inplace = True)
    order = ['楼栋','宿舍号','次数','学院']
    si = si[order]
    name1 = '各学院未熄灯宿舍记录报表'+output+'.xlsx'
    writer1 = pd.ExcelWriter(name1)
    si.to_excel(excel_writer=writer1, index=False)
    writer1.save()
    writer1.close()
    #按照学院制作sheet
    gr2 = data.groupby(['学院'])
    for name,group in gr2:
        filename = desc+'\\'+name + '未熄灯宿舍记录报表'+output+'.xlsx'
        writer = pd.ExcelWriter(filename)
        group = group.drop(['检查人'],axis=1)
        temp = group['日期']
        temp_str=temp.apply(lambda x: datetime.datetime.strftime(x,"%d/%m/%Y"))
        group['日期'] = temp_str
        group.to_excel(excel_writer=writer, index=False)
        writer.save()
        writer.close()

	

#if __name__ == '__main__':
#    ##提前在桌面创建好desc=ff文件夹，否则会报错
#    main(postfix='(2018年12月09日-2018年12月16日)',filename='1209_1216.txt',fri=(2018,12,14),desc='ff',output='学院总表1209_1216.xlsx')
	

    