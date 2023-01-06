import pandas as pd
import numpy as np
import os

io = r'..\Level 3 VS scorecard.xlsx'
VS_input='..\名字-ValueStream.xlsx'
PID_input_path = '../主管级PID审核分值统计表/'
ldl_input_path = '../主管领导力特质评分表/'



PID_output='./extract_主管级PID审核分值统计表.xlsx'
ldl_output='./extract_主管领导力特质评分表.xlsx'


print('\n############################# Start ################################\n\n')

with open("sum.txt","w") as f:
        f.write("############################### Start ##############################\n\n")  

month_list=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']






def strtomonth(month_str):
    for i in range(len(month_list)):
        if month_list[i] in month_str:
            return i+1
    return 0

def PIDtoexcel():
    PID_files = os.listdir(PID_input_path)
    df_PID=pd.DataFrame(columns=['名字','月份','得分'])
    for file in PID_files:
        #print(file.split('-')[-1].split('.')[0])
        if strtomonth(file)!=0:
            month=strtomonth(file)
        else:
            print('month error!!!')
        input=PID_input_path+file
        # print(input)
        sheet_month='主管级PID 审核分值统计-{}月'.format(month)
        data_PID = pd.read_excel(input,sheet_name=sheet_month,header=2,usecols=['主管','平均分'])
        data_PID=data_PID.dropna()
        # print(data_PID)
        #有些人没数据
        for i in data_PID.values:
            #print(i[0].split('/'))
            for j in i[0].split('/'):
                df_row=df_PID.shape[0]
                df_PID.loc[df_row]=[j,month,i[1]]
    df_PID.to_excel(PID_output)

def ldltoexcel():
    ldl_files = os.listdir(ldl_input_path)
    df_zl_ldl=pd.DataFrame(columns=['名字','月份','得分'])
    for file in ldl_files:
        #print(file.split('-')[-1].split('.')[0])
    #     index=month_list.index(file.split('-')[-1].split('.')[0])
        if strtomonth(file)!=0:
            month=strtomonth(file)
        else:
            print('month error!!!')
        input=ldl_input_path+file
        #print(input)
        df = pd.read_excel(input, sheet_name=None)
        list_df=list(df)
        for i in list_df:
            data = pd.read_excel(input, sheet_name = i)
            name=i.split('-')[-1]
            score=data[data['Unnamed: 1'].isin(['合计'])]['Unnamed: 6'].values[0]
            df_row=df_zl_ldl.shape[0]
            df_zl_ldl.loc[df_row]=[name,month,score]
            # print(i.split('-')[-1])
            # print(data[data['Unnamed: 1'].isin(['合计'])]['Unnamed: 6'].values[0])
            df_zl_ldl.to_excel(ldl_output)
    

def write_txt(list,list_1,zhibiao):
    with open("sum.txt","a") as f:
        f.write('\n(指标名称:{} )\n'.format(zhibiao))
        print('(指标名称:{} )'.format(zhibiao))
        for i in range(len(list_1)):
            f.write('[{}] read data:{}, 实际得分：{}\n'.format(month_list[i],list[i+1],list_1[i]))
            print(('[{}] read data:{}, 实际得分：{}'.format(month_list[i],list[i+1],list_1[i])))

def write_txt_s(string):
    with open("sum.txt","a") as f:
        f.write('\n{}'.format(string))


def RIF(list):
    print(('----Start RIF----'))
    write_txt_s('----Start RIF----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #目标0伤害: 达标100分/每增加一起可记录伤害扣30分
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if 0<=list[i]<=3:
                    list_1_n=100-list[i]*30
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            elif isinstance(list[i],str):
                if 0<=list[i]<=3:
                    list_1_n=100-float(list[i])*30
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'RIF(可记录伤害）')
    return list_1

def LTCFR(list):
    write_txt_s('----Start LTCFR（损失工作日）----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #目标0损失: 达标100分/每增加一天工作日损失扣20分
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if 0<=list[i]<=5:
                    list_1_n=100-list[i]*20
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            elif isinstance(list[i],str):
                if 0<=list[i]<=5:
                    list_1_n=100-float(list[i])*20
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'LTCFR（损失工作日）')
    return list_1

def ABBS(list):
    write_txt_s('----Start  ABBS基于现场安全巡查行动的完成率（每周一次）----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
           #目标要求每周进行一次，为100分达标； 少于目标一次扣25分，
            #如一次都未进行即没有ABBS 巡查，为0分；如多于4次巡查不加分
            
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                list_1.append(list[i]*100)
            elif isinstance(list[i],str):
                list_1.append(float(list[i])*100)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,' ABBS基于现场安全巡查行动的完成率（每周一次）')
    return list_1

def anquan(list):
    write_txt_s('----Start  安全改进措施的完成率95%----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
           #目标95%: 达标为100分，超标不加分/不达标 相应百分比转换成百分数分值
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],float) or isinstance(list[i],int):
                if list[i]>=0.95:
                    list_1.append(100)
                else:
                    list_1_n=list[i]/0.95*100
                    list_1.append(list_1_n)
            elif isinstance(list[i],str):
                if float(list[i])>=0.95:
                    list_1.append(100)
                else:
                    list_1_n=float(list[i])/0.95*100
                    list_1.append(list_1_n)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,' 安全改进措施的完成率95%')
    return list_1

def OOT(list):
    write_txt_s('----Start # of operators OT > 36H/M----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #目标0人：达标100分/加班超36小时发生1人扣50分，以此类推
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],float) or isinstance(list[i],int):
                if 0<=list[i]<=2:
                    list_1_n=100-list[i]*50
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            elif isinstance(list[i],str):
                stf=float(list[i])
                if 0<=stf<=2:
                    list_1_n=100-stf*50
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'# of operators OT > 36H/M')
    return list_1

def YE(list):
    write_txt_s('----Start Missed defect (YE)----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #目标0逃逸缺陷：达标100分/逃逸缺陷发生1个扣10分，以此类推
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if 0<=list[i]<=10:
                    list_1_n=100-list[i]*10
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            elif isinstance(list[i],str):
                if 0<=list[i]<=10:
                    list_1_n=100-float(list[i])*10
                    list_1.append(list_1_n)
                else :
                    list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'Missed defect (YE)')
    return list_1

# #
# eff_list=[0,0,2,2,2,4,4,4,6,8,8,10]
# def Eff(list):
#     list_1=[]
#     for i in range(len(list)):
# #         print(list[i],type(list[i]))
#         if i==0:
#             pass
#         else:
#             #目标6%：达标100分，超标不加分/ 
#             #未达标，相应百分比转换成百分数分值，用实际数➗6%✖100
#             if pd.isnull(list[i]):
#                 if eff_list[i-1]>=0.06:
#                     print('Null!')
#                     list_1.append(0)
#                 else:
#                     list_1_n=eff_list[i-1]/0.06*100
#                     list_1.append(list_1_n)
#             elif isinstance(list[i],float) or isinstance(list[i],int):
#                 if list[i]>=0.06:
#                     list_1.append(100)
#                 else:
#                     list_1_n=list[i]/0.06*100
#                     list_1.append(list_1_n)
#             elif isinstance(list[i],str):
#                 if float(list[i])>=0.06:
#                     list_1.append(100)
#                 else:
#                     list_1_n=float(list[i])/0.06*100
#                     list_1.append(list_1_n)
#             else: 
#                 print('Eff num error')
#                 list_1.append(0)
#     return list_1

#与plan 值对比
eff_count=False
eff_list_plan=[]
def Eff(list):
    write_txt_s('----Start Efficiency improvement----')
    global eff_count,eff_list_plan
    list_1=[]
    if not eff_count:
        eff_list_plan=list
        eff_count=True
    else:
        for i in range(len(list)):
    #         print(list[i],type(list[i]))
            if i==0:
                pass
            else:
                #达到目标值：达标100分，超标不加分/ 未达标；
                # 减少10%以内，扣10分，10%-20%之间，扣20分，以此类推
                if pd.isnull(list[i]):
                    print('{} Data is Null!'.format(month_list[i-1]))
                    write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                    list_1.append(0)
                elif isinstance(list[i],float) or isinstance(list[i],int):
                    if list[i]>=eff_list_plan[i]:
                        list_1.append(100)
                    elif 0<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.1:
                        list_1.append(90)
                    elif 0.1<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.2:
                        list_1.append(80)
                    elif 0.2<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.3:
                        list_1.append(70)
                    elif 0.3<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.4:
                        list_1.append(60)
                    elif 0.4<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.5:
                        list_1.append(50)
                    elif 0.5<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.6:
                        list_1.append(40)
                    elif 0.6<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.7:
                        list_1.append(30)
                    elif 0.7<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.8:
                        list_1.append(20)
                    elif 0.8<(eff_list_plan[i]-list[i])/eff_list_plan[i]<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)

                elif isinstance(list[i],str):
                    if float(list[i])>=eff_list_plan[i]:
                        list_1.append(100)
                    elif 0<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.1:
                        list_1.append(90)
                    elif 0.1<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.2:
                        list_1.append(80)
                    elif 0.2<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.3:
                        list_1.append(70)
                    elif 0.3<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.4:
                        list_1.append(60)
                    elif 0.4<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.5:
                        list_1.append(50)
                    elif 0.5<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.6:
                        list_1.append(40)
                    elif 0.6<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.7:
                        list_1.append(30)
                    elif 0.7<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.8:
                        list_1.append(20)
                    elif 0.8<(eff_list_plan[i]-float(list[i]))/eff_list_plan[i]<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)
                else: 
                    print('{} num error!'.format(month_list[i-1]))
                    write_txt_s('{} Num Error!'.format(month_list[i-1]))
                    list_1.append(0)
        eff_count=False
        write_txt(eff_list_plan,list_1,'Efficiency improvement')
        write_txt(list,list_1,'Efficiency improvement')
        eff_list_plan=[]
        
    return list_1



def MS(list):
    write_txt_s('----Start Material Scrap 材料报废金额----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #报废控制在$1000内为达标100分；超10%以内，扣10分，10%-20%之间，扣20分，以此类推
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if list[i]<=1000:
                    list_1.append(100)
                else:
                    if 0<(list[i]-1000)/1000<=0.1:
                        list_1.append(90)
                    elif 0.1<(list[i]-1000)/1000<=0.2:
                        list_1.append(80)
                    elif 0.2<(list[i]-1000)/1000<=0.3:
                        list_1.append(70)
                    elif 0.3<(list[i]-1000)/1000<=0.4:
                        list_1.append(60)
                    elif 0.4<(list[i]-1000)/1000<=0.5:
                        list_1.append(50)
                    elif 0.5<(list[i]-1000)/1000<=0.6:
                        list_1.append(40)
                    elif 0.6<(list[i]-1000)/1000<=0.7:
                        list_1.append(30)
                    elif 0.7<(list[i]-1000)/1000<=0.8:
                        list_1.append(20)
                    elif 0.8<(list[i]-1000)/1000<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)
            elif isinstance(list[i],str):
                if float(list[i])<=1000:
                    list_1.append(100)
                else:
                    if 0<(float(list[i])-1000)/1000<=0.1:
                        list_1.append(90)
                    elif 0.1<(float(list[i])-1000)/1000<=0.2:
                        list_1.append(80)
                    elif 0.2<(float(list[i])-1000)/1000<=0.3:
                        list_1.append(70)
                    elif 0.3<(float(list[i])-1000)/1000<=0.4:
                        list_1.append(60)
                    elif 0.4<(float(list[i])-1000)/1000<=0.5:
                        list_1.append(50)
                    elif 0.5<(float(list[i])-1000)/1000<=0.6:
                        list_1.append(40)
                    elif 0.6<(float(list[i])-1000)/1000<=0.7:
                        list_1.append(30)
                    elif 0.7<(float(list[i])-1000)/1000<=0.8:
                        list_1.append(20)
                    elif 0.8<(float(list[i])-1000)/1000<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'Material Scrap 材料报废金额')
    return list_1


def PID(name):
    write_txt_s('----Start PID:Maturity and effectiveness  Evaluation ----')
    global df_PID
    list_1=[]
    for i in range(12):
        #print(df_PID[df_PID['名字'].isin([name]) & df_PID['月份'].isin([i+1])]['得分'].values)
        if len(df_PID[df_PID['名字'].isin([name]) & df_PID['月份'].isin([i+1])]['得分'].values)!=0:
            #1.目标85分： 达标为100分
            #2. 未达标或超标完成，用相应得分转换成百分数分值，即用审核分数➗85✖100
            value_PID=df_PID[df_PID['名字'].isin([name]) & df_PID['月份'].isin([i+1])]['得分'].values[0]
            # print(type(value_PID))
            if pd.isnull(value_PID):
                print('{} Data is Null!'.format(month_list[i]))
                write_txt_s('{} Data is Null!'.format(month_list[i]))
                list_1.append(0)
            elif isinstance(value_PID,float) or isinstance(value_PID,int):
                if value_PID>=85:
                    list_1.append(100)
                    write_txt_s('[{}] read data:{}, 实际得分：{}'.format(month_list[i],value_PID,100))
                else:
                    list_1_n=value_PID/85*100
                    list_1.append(list_1_n)
                    write_txt_s('[{}] read data:{}, 实际得分：{}'.format(month_list[i],value_PID,list_1_n))
            else:
                if float(value_PID)>=85:
                    list_1.append(100)
                    write_txt_s('[{}] read data:{}, 实际得分：{}'.format(month_list[i],value_PID,100))
                else:
                    list_1_n=float(value_PID)/85*100
                    list_1.append(list_1_n)
                    write_txt_s('[{}] read data:{}, 实际得分：{}'.format(month_list[i],value_PID,list_1_n))
            # else:
            #     print('PID num error')
            #     list_1.append(0)
        else:
            list_1.append(0)
            print('{} num error!'.format(month_list[i]))
            write_txt_s('{} Num Error!'.format(month_list[i]))
    return list_1


#小问题
def CI(list):
    write_txt_s('----Start Valid CI per Month----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            #1. 目标人均有效CI条数1： 达标为100分 ；
            #2. 不达标扣分，每少于目标1%-10%，扣20分；11%-20%，扣40分，以此类推；
            #3.人均CI 个数在0.5以下该项不得分
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if list[i]>=1:
                    list_1.append(100)
                else:
                    if 0<(1-list[i])<=0.1:
                        list_1.append(80)
                    elif 0.1<(1-list[i])<=0.2:
                        list_1.append(60)
                    elif 0.2<(1-list[i])<=0.3:
                        list_1.append(40)
                    elif 0.3<(1-list[i])<=0.4:
                        list_1.append(20)
                    else:
                        list_1.append(0)
            elif isinstance(list[i],str):
                if float(list[i])>=1:
                    list_1.append(100)
                else:
                    if 0<(1-float(list[i]))<=0.1:
                        list_1.append(80)
                    elif 0.1<(1-float(list[i]))<=0.2:
                        list_1.append(60)
                    elif 0.2<(1-float(list[i]))<=0.3:
                        list_1.append(40)
                    elif 0.3<(1-float(list[i]))<=0.4:
                        list_1.append(20)
                    else:
                        list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'Valid CI per Month）')
    return list_1

#与plan 值对比
UOT_count=False
UOT_list_plan=[]
def UOT(list):
    write_txt_s('----Start Unplan OT(现在系统显示效率休假等）----')
    global UOT_count,UOT_list_plan
    list_1=[]
    if not UOT_count:
        UOT_list_plan=list
        print(UOT_list_plan)
        UOT_count=True
    else:
        print(UOT_list_plan)
        for i in range(len(list)):
    #         print(list[i],type(list[i]))
            if i==0:
                pass
            else:
                #1. 目标非计划加班小时数跟目标一致或低于目标，达标100分；
                #2. 超10%以内，扣10分，10%-20%之间，扣20分，以此类推
                if pd.isnull(list[i]):
                    print('{} Data is Null!'.format(month_list[i-1]))
                    write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                    list_1.append(0)
                elif isinstance(list[i],float) or isinstance(list[i],int):
                    if list[i]<=UOT_list_plan[i]:
                        list_1.append(100)
                    elif 0<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.1:
                        list_1.append(90)
                    elif 0.1<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.2:
                        list_1.append(80)
                    elif 0.2<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.3:
                        list_1.append(70)
                    elif 0.3<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.4:
                        list_1.append(60)
                    elif 0.4<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.5:
                        list_1.append(50)
                    elif 0.5<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.6:
                        list_1.append(40)
                    elif 0.6<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.7:
                        list_1.append(30)
                    elif 0.7<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.8:
                        list_1.append(20)
                    elif 0.8<(list[i]-UOT_list_plan[i])/UOT_list_plan[i]<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)

                elif isinstance(list[i],str):
                    if float(list[i])<=UOT_list_plan[i]:
                        list_1.append(100)
                    elif 0<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.1:
                        list_1.append(90)
                    elif 0.1<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.2:
                        list_1.append(80)
                    elif 0.2<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.3:
                        list_1.append(70)
                    elif 0.3<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.4:
                        list_1.append(60)
                    elif 0.4<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.5:
                        list_1.append(50)
                    elif 0.5<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.6:
                        list_1.append(40)
                    elif 0.6<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.7:
                        list_1.append(30)
                    elif 0.7<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.8:
                        list_1.append(20)
                    elif 0.8<(float(list[i])-UOT_list_plan[i])/UOT_list_plan[i]<=0.9:
                        list_1.append(10)
                    else:
                        list_1.append(0)
                else: 
                    print('{} num error!'.format(month_list[i-1]))
                    write_txt_s('{} Num Error!'.format(month_list[i-1]))
                    list_1.append(0)
        UOT_count=False
        UOT_list_plan=[]
    write_txt(list,list_1,'Unplan OT(现在系统显示效率休假等）')
    return list_1


def Gemba(list):
    write_txt_s('----Start Gemba 巡查活动的完成率和发现项>=2----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:
            # 1.每周Gamba 一次，每月发现项2条为达标：达标为100分；
            # 2.Gamba 次数少于一次扣20分，少于2次为0分
            if pd.isnull(list[i]):
                #未知
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],int) or isinstance(list[i],float):
                if list[i]>=2:
                    list_1.append(100)
                elif 1<=list[i]<2:
                    list_1.append(50)
                elif 0<=list[i]<1:
                    list_1.append(0)
                else:
                    list_1.append(0)
            elif isinstance(list[i],str):
                if float(list[i])>=2:
                    list_1.append(100)
                elif 1<=float(list[i])<2:
                    list_1.append(50)
                elif 0<=float(list[i])<1:
                    list_1.append(0)
                else:
                    list_1.append(0)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,'Gemba 巡查活动的完成率和发现项>=2')
    return list_1

def EI(list):
    write_txt_s('----Start  EI Engagement Index (%)----')
    list_1=[]
    for i in range(len(list)):
        if i==0:
            pass
        else:        
            #             print(type(list[i]),list[i])
            #             1. 目标95%达标： 达标为100分；
            #             2. 不达标或超标，用相应百分比转换成百分数分值，即用实际分数➗95%✖100
            if pd.isnull(list[i]):
                print('{} Data is Null!'.format(month_list[i-1]))
                write_txt_s('{} Data is Null!'.format(month_list[i-1]))
                list_1.append(0)
            elif isinstance(list[i],float) or isinstance(list[i],int):
                list_1_n=list[i]/0.95*100
                list_1.append(list_1_n)
            elif isinstance(list[i],str):
                list_1_n=float(list[i])/0.95*100
                list_1.append(list_1_n)
            else:
                print('{} num error!'.format(month_list[i-1]))
                write_txt_s('{} Num Error!'.format(month_list[i-1]))
                list_1.append(0)
    write_txt(list,list_1,' EI Engagement Index (%)')
    return list_1



def yongyv(name):
    write_txt_s('----Start 勇于挑战，勇于改变，充满好奇，发展团队----')
    global df_yongyv
    list_1=[]
    for i in range(12):
        #print(df_yongyv[df_yongyv['名字'].isin([name]) & df_yongyv['月份'].isin([i+1])]['得分'].values)
        if len(df_yongyv[df_yongyv['名字'].isin([name]) & df_yongyv['月份'].isin([i+1])]['得分'].values)!=0:
            #1.目标85分： 达标为100分
            #2. 未达标或超标完成，用相应得分转换成百分数分值，即用审核分数➗85✖100
            value_yongyv=df_yongyv[df_yongyv['名字'].isin([name]) & df_yongyv['月份'].isin([i+1])]['得分'].values[0]
            if pd.isnull(value_yongyv):
                print('{} Data is Null!'.format(month_list[i]))
                write_txt_s('{} Data is Null!'.format(month_list[i]))
                list_1.append(0)
            elif isinstance(value_yongyv,float) or isinstance(value_yongyv,int):
                list_1.append(value_yongyv)
            else:
                list_1.append(float(value_yongyv))
            
            # else:
            #     print('yongyv num error')
            #     list_1.append(0)
        else:
            list_1.append(0)
            print('{} num error!'.format(month_list[i]))
            write_txt_s('{} Num Error!'.format(month_list[i]))
        write_txt_s('[{}] read data:{}, 实际得分：{}'.format(month_list[i],value_yongyv,list_1[-1]))
    return list_1


# for list in sheet3.values:
#     if list[0]=='RIF(可记录伤害）':
#         print(RIF(list))
#     elif list[0]=='LTCFR（损失工作日）':
#         print(LTCFR(list))
#     elif list[0]=='ABBS完成率':
#         print(ABBS(list))
#     elif list[0]=='安全不符合项及时关闭率':
#         print(anquan(list))
#     elif list[0]=='# of operators OT > 36H/M':
#         print(OOT(list))
#     elif list[0]=='Missed defect (YE)- Assy':
#         print(YE(list))
#     elif list[0]=='Efficiency improvement-Assy':
#         print(Eff(list))
#     elif list[0]=='Material Scrap 材料报废金额':
#         print(MS(list))
#     elif list[0]=='PID:Maturity and effectiveness  Evaluation ':
#         print(PID(list))
#     elif list[0]=='Valid CI per Month':
#         print(CI(list))
#     elif list[0]=='UPOT-Assy（非计划加班）':
#         print(UOT(list))
#     elif list[0]=='Gemba 巡查活动的完成率和发现项>=2':
#         print(Gemba(list))
#     elif list[0]=='EI Engagement Index (%)':
#         print(EI(list))
#     else:
#         print('sheet3 error')
        
    
def people(name,first_index='Unknow'):
    global df_a,df_b
    with open("sum.txt","a") as f:
        f.write('\n######################################################################\n') 
        f.write('#####################< Start Calculate :{} >#####################\n'.format(name))  
        f.write('######################################################################\n') 
    #print('-------------------------------------------------------- ')
    print('######################################################################')
    print('#####################< Start Calculate :{} >#####################'.format(name))
    print('######################################################################')
    #print('-------------------------------------------------------- ')
    global RIF_list,LTCFR_list,ABBS_list,anquan_list,OOT_list,YE_list,Eff_list,MS_list,PID_list,CI_list,UOT_list,Gemba_list,EI_list,PID_series
        #截取一个表的规则
    sheet_index1=data[data['指标名称'].isin([name])].index.tolist()
    #print(sheet_index1)
    # sheet_index2=data[data['指标分类1'].isin([last_index])].index.tolist()
    #print(sheet_index2)
    # if sheet_index2==sheet_index1:
    #     sheet_index2=[sheet_index1[0]+69]
    sheet_index2=[sheet_index1[0]+69]
    if len(sheet_index1)!=1:
        print('index error')
    # else:
    #     print(sheet_index2[0]-sheet_index1[0])
    sheet1=data[:][sheet_index1[0]:sheet_index2[0]]
    sheet2=data_1[:][sheet_index1[0]:sheet_index2[0]]
    series_1=sheet1['指标名称'].isin(['RIF(可记录伤害）','LTCFR（损失工作日）','EI Engagement Index (%)','ABBS完成率','安全不符合项及时关闭率','Missed defect (YE)- Assy','# of operators OT > 36H/M'])
    series_2=series_1.copy()
    for key,value in enumerate(series_1,start=sheet_index1[0]):
        if value==True:
            series_2[key]=False
            series_2[key+1]=True
    series_a=sheet1['指标名称'].isin(['Material Scrap 材料报废金额','PID:Maturity and effectiveness  Evaluation ','Valid CI per Month','Gemba 巡查活动的完成率和发现项>=2'])
    for key,value in enumerate(series_a,start=sheet_index1[0]):
        if value==True:
            series_2[key]=True
    #eff UOT 与plan对比类
    series_b=sheet2['指标名称'].isin(['Efficiency improvement-Assy','UPOT-Assy（非计划加班）'])
    for key,value in enumerate(series_b,start=sheet_index1[0]):
        if value==True:
            series_2[key]=True
    sheet3=sheet2.loc[series_2,:]
    del sheet3['指标分类1']
    #print(sheet3.values)
    
    for list_v in sheet3.values:
        if list_v[0]=='RIF(可记录伤害）':
            #print(RIF(list_v))
            RIF_list=RIF(list_v)
        elif list_v[0]=='LTCFR（损失工作日）':
            #print(LTCFR(list_v))
            LTCFR_list=LTCFR(list_v)
        elif list_v[0]=='ABBS完成率':
            #print(ABBS(list_v))
            ABBS_list=ABBS(list_v)
        elif list_v[0]=='安全不符合项及时关闭率':
            #print(anquan(list_v))
            anquan_list=anquan(list_v)
        elif list_v[0]=='# of operators OT > 36H/M':
            #print(OOT(list_v))
            OOT_list=OOT(list_v)
        elif list_v[0]=='Missed defect (YE)- Assy':
            #print(YE(list_v))
            YE_list=YE(list_v)
        elif list_v[0]=='Efficiency improvement-Assy':
            #print(Eff(list_v))
            Eff_list=Eff(list_v)
        elif list_v[0]=='Material Scrap 材料报废金额':
            #print(MS(list_v))
            MS_list=MS(list_v)
        # elif list_v[0]=='PID:Maturity and effectiveness  Evaluation ':
        #     #print(PID(list_v))
        #     PID_list=PID(list_v)
        elif list_v[0]=='Valid CI per Month':
            #print(CI(list_v))
            CI_list=CI(list_v)
        elif list_v[0]=='UPOT-Assy（非计划加班）':
            #print(UOT(list_v))
            UOT_list=UOT(list_v)
        elif list_v[0]=='Gemba 巡查活动的完成率和发现项>=2':
            #print(Gemba(list_v))
            Gemba_list=Gemba(list_v)
        elif list_v[0]=='EI Engagement Index (%)':
            #print(EI(list_v))
            EI_list=EI(list_v)
        else:
            print('sheet3 error')
    


    PID_list=PID(name)
    yongyv_list=yongyv(name)
    
    
    
    for i in range(12):
        Performance=0.1*RIF_list[i]+0.1*LTCFR_list[i]+0.03*ABBS_list[i]+0.03*anquan_list[i]+0.1*OOT_list[i]+0.1*YE_list[i]+0.12*Eff_list[i]+0.02*MS_list[i]
        Culture=0.02*PID_list[i]+0.02*CI_list[i]+0.02*UOT_list[i]+0.02*Gemba_list[i]+0.02*EI_list[i]+0.3*yongyv_list[i]
        Performance_Culture=Performance+Culture
        df_row=df_a.shape[0]
        df_b.loc[df_row]=[name,first_index,'2021-{}'.format(i+1),
                          0.1*RIF_list[i],
                          0.1*LTCFR_list[i],
                          0.03*ABBS_list[i],
                          0.03*anquan_list[i],
                          0.1*OOT_list[i],
                          0.1*YE_list[i],
                          0.12*Eff_list[i],
                          0.02*MS_list[i],
                          0.02*PID_list[i],
                          0.02*CI_list[i],
                          0.02*UOT_list[i],
                          0.02*Gemba_list[i],
                          0.02*EI_list[i],
                          0.3*yongyv_list[i],
                          Performance,
                          Culture,
                          Performance_Culture]
        df_a.loc[df_row]=[name,first_index,'2021-{}'.format(i+1),Performance,Culture,Performance_Culture]
        print('<2021-{}> Performance:{}, Culture & attributes:{}, Total Score:{}'.format(i+1,round(Performance,2),round(Culture,2),round(Performance_Culture,2)))
        write_txt_s('<2021-{}> Performance:{}, Culture & attributes:{}, Total Score:{}'.format(i+1,round(Performance,2),round(Culture,2),round(Performance_Culture,2)))
    df_a.to_excel('sum_simple.xlsx')
    df_b.to_excel('sum.xlsx')
    #print('-------------------------------------------------------- ')
    print('######################################################################')
    print('######################< End Calculate :{} >#########################'.format(name))
    print('######################################################################')
    with open("sum.txt","a") as f:
        f.write('\n######################################################################\n') 
        f.write('######################< End Calculate :{} >#########################\n'.format(name))  
        f.write('######################################################################\n') 
    #print('-------------------------------------------------------- ')
    print(' ')

    
# PIDtoexcel()
# ldltoexcel()



#月份设置  
data = pd.read_excel(io, sheet_name = 'L3&VS-Assy',header=2,usecols=['指标分类1','指标名称', 'Jan', 'Feb', 'Mar', 'Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])
data_1=data.copy()
data_1['指标名称']=data_1['指标名称'].fillna(method="ffill",limit=1)
# data_PID = pd.read_excel(PID_output,sheet_name='主管级PID 审核分值统计-10月',header=2,usecols=['主管','平均分'])
# data_PID=data_PID.dropna()
# #有些人没数据
# df_PID=pd.DataFrame(columns=['主管','平均分'])
# for i in data_PID.values:
#     #print(i[0].split('/'))
#     for j in i[0].split('/'):
#         df_row=df_PID.shape[0]
#         df_PID.loc[df_row]=[j,i[1]]
# df_PID.to_excel('整理-主管级PID审核分值统计.xlsx')

df_PID = pd.read_excel(PID_output)
df_yongyv=pd.read_excel(ldl_output)



df_a=pd.DataFrame(columns=['名字','Value_Stream','月份','Performance','Culture & attributes','总分'])
df_b=pd.DataFrame(columns=['名字','Value_Stream','月份','RIF(可记录伤害）',
                                    'LTCFR（损失工作日）',
                                    'ABBS完成率','安全不符合项及时关闭率',
                                    '# of operators OT > 36H/M',
                                    'Missed defect (YE)- Assy',
                                   'Efficiency improvement-Assy',
                                    'Material Scrap 材料报废金额',
                                    'PID:Maturity and effectiveness  Evaluation ',
                                    'Valid CI per Month',
                                    'UPOT-Assy（非计划加班）',
                                    'Gemba 巡查活动的完成率和发现项>=2',
                                    'EI Engagement Index (%)',
                               '勇于挑战，勇于改变，充满好奇，发展团队',
                           'Performance','Culture & attributes','总分'
                                   ])

df_VS=pd.read_excel(VS_input)

print(df_VS.values)
for i in df_VS.values:
    people(i[0],i[1])
with open("sum.txt","a") as f:
        f.write("\n\n\n############################### Finish ##############################\n\n")  
print("############################### Finish ##############################\n\n")  


